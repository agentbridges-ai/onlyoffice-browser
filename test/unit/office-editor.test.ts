import { beforeEach, describe, expect, it, vi } from 'vitest';

import { OFFICE_HOST_PROTOCOL, type OfficeHostParentMessage } from '../../src/lib/office-host-protocol';
import { createOfficeEditor, loadOfficeEditorApi, mountOfficeEditor } from '../../src/lib/office-editor';
import { OFFICE_EDITOR_ORIGIN_SLOTS } from '../../src/lib/office-origin-pool';

const HOST_URL = 'http://127.0.0.1:5173/office-host.html';
const HOST_IDENTITY = {
  packageVersion: '0.3.29',
  hostBuildId: 'office-host-0.3.29-r1',
  assetManifestDigest: 'a'.repeat(64),
};

function flush(): Promise<void> {
  return Promise.resolve().then(() => undefined);
}

function waitForMessage(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function getSessionId(iframe: HTMLIFrameElement): string {
  const url = new URL(iframe.src);
  return new URLSearchParams(url.hash.slice(1)).get('sessionId') || url.searchParams.get('sessionId') || '';
}

async function waitForIframe(container: HTMLElement): Promise<HTMLIFrameElement> {
  for (let i = 0; i < 20; i += 1) {
    const iframe = container.querySelector<HTMLIFrameElement>('iframe');
    if (iframe) return iframe;
    await flush();
  }
  throw new Error('Timed out waiting for host iframe');
}

async function connectHost(
  container: HTMLElement,
  onChildMessage?: (message: OfficeHostParentMessage, childPort: MessagePort) => void,
): Promise<{ childPort: MessagePort; iframe: HTMLIFrameElement; messages: OfficeHostParentMessage[] }> {
  const iframe = await waitForIframe(container);
  const parentWindow = container.ownerDocument.defaultView || window;
  const sessionId = getSessionId(iframe);
  const hostOrigin = new URL(iframe.src).origin;
  const messages: OfficeHostParentMessage[] = [];
  let childPort: MessagePort | null = null;

  vi.spyOn(iframe.contentWindow! as any, 'postMessage').mockImplementation((...args: unknown[]) => {
    const transfer = args[2] as Transferable[] | undefined;
    childPort = transfer?.[0] as MessagePort;
    childPort.onmessage = (event: MessageEvent<OfficeHostParentMessage>) => {
      messages.push(event.data);
      onChildMessage?.(event.data, childPort!);
      if (event.data.type === 'INIT') {
        const sourceKind =
          event.data.options.source.kind === 'empty' ? 'new-document' : event.data.options.source.sourceKind;
        childPort!.postMessage({
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'READY',
          sessionId,
          state: {
            id: sessionId,
            fileName: event.data.options.fileName || 'document.docx',
            fileType: (event.data.options.fileName || 'document.docx').split('.').pop() || 'docx',
            mode: event.data.options.mode || (event.data.options.readonly ? 'readonly' : 'edit'),
            readonly: event.data.options.mode === 'preview' || Boolean(event.data.options.readonly),
            dirty: false,
            sourceKind,
            status: 'ready',
            destroyed: false,
          },
        });
      }
    };
    childPort.start();
  });

  parentWindow.dispatchEvent(
    new parentWindow.MessageEvent('message', {
      origin: hostOrigin,
      source: iframe.contentWindow,
      data: {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'HOST_READY',
        sessionId,
        identity: HOST_IDENTITY,
      },
    }),
  );

  await flush();
  return { childPort: childPort!, iframe, messages };
}

describe('office-editor parent proxy', () => {
  beforeEach(() => {
    vi.useRealTimers();
    document.head.innerHTML = '';
    document.body.innerHTML = '';
    vi.restoreAllMocks();
    Object.defineProperty(URL, 'createObjectURL', {
      configurable: true,
      value: vi.fn(() => 'blob:onlyoffice-browser-parent-download'),
    });
    Object.defineProperty(URL, 'revokeObjectURL', {
      configurable: true,
      value: vi.fn(),
    });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
  });

  it('keeps loadOfficeEditorApi as a parent-side no-op', async () => {
    await expect(loadOfficeEditorApi()).resolves.toBeUndefined();
    expect(document.querySelectorAll('script[data-office-editor-api="true"]')).toHaveLength(0);
  });

  it('mounts an isolated host synchronously and defers document activation', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const mount = mountOfficeEditor(container, {
      hostUrl: window.location.origin + '/office-host.html',
      file: new File(['a'], 'deferred.docx'),
      fileName: 'deferred.docx',
      destroyTimeoutMs: 1,
    });

    const iframe = container.querySelector<HTMLIFrameElement>('iframe');
    expect(iframe).not.toBeNull();
    expect(getSessionId(iframe!)).toBe(mount.id);
    expect(mount.getState()).toMatchObject({ phase: 'host-loading' });

    const connection = await connectHost(container);
    expect(connection.messages).toEqual([]);
    expect(mount.getState()).toMatchObject({ phase: 'waiting-for-activation' });

    const activation = mount.activate();
    await expect(activation).resolves.toMatchObject({ id: mount.id });
    expect(connection.messages[0]).toMatchObject({ type: 'INIT' });
    expect(mount.getState()).toMatchObject({ phase: 'ready' });
    const debug = (
      window as typeof window & {
        __officeHostDebug?: {
          activeHostPortCount: number;
          peakActiveHostPortCount: number;
          activeInstanceCount: number;
          activeOriginLeaseCount: number;
        };
      }
    ).__officeHostDebug;
    expect(debug).toMatchObject({
      activeHostPortCount: 1,
      activeInstanceCount: 1,
      activeOriginLeaseCount: 1,
    });
    expect(debug!.peakActiveHostPortCount).toBeGreaterThanOrEqual(1);

    await mount.destroy();
    expect(mount.getState()).toMatchObject({ phase: 'destroyed' });
    expect(debug).toMatchObject({
      activeHostPortCount: 0,
      activeInstanceCount: 0,
      activeOriginLeaseCount: 0,
    });
  });

  it('routes out-of-order READY messages and destruction to independent mounts', async () => {
    const firstContainer = document.createElement('div');
    const secondContainer = document.createElement('div');
    document.body.append(firstContainer, secondContainer);
    const options = (fileName: string) => ({
      hostUrl: window.location.origin + '/office-host.html',
      file: new File([fileName], fileName),
      fileName,
      destroyTimeoutMs: 1,
    });

    const first = mountOfficeEditor(firstContainer, options('first.docx'));
    const second = mountOfficeEditor(secondContainer, options('second.xlsx'));
    const firstFrame = firstContainer.querySelector<HTMLIFrameElement>('iframe')!;
    const secondFrame = secondContainer.querySelector<HTMLIFrameElement>('iframe')!;

    expect(first.id).not.toBe(second.id);
    expect(new URL(firstFrame.src).origin).not.toBe(new URL(secondFrame.src).origin);

    const firstReady = first.activate();
    const secondReady = second.activate();
    await connectHost(secondContainer);
    await expect(secondReady).resolves.toMatchObject({ id: second.id });
    expect(first.getState().phase).not.toBe('ready');
    await connectHost(firstContainer);
    await expect(firstReady).resolves.toMatchObject({ id: first.id });

    await first.destroy();
    expect(firstContainer.querySelector('iframe')).toBeNull();
    expect(secondContainer.querySelector('iframe')).toBe(secondFrame);
    expect(second.getState().phase).toBe('ready');
    await second.destroy();
  });

  it('rejects two active mounts that resolve to the same host origin', async () => {
    const firstContainer = document.createElement('div');
    const secondContainer = document.createElement('div');
    document.body.append(firstContainer, secondContainer);
    const options = {
      hostUrl: 'https://shared.office-host.example.com/office-host.html',
      file: new File(['a'], 'shared.docx'),
      fileName: 'shared.docx',
      destroyTimeoutMs: 1,
    };

    const first = mountOfficeEditor(firstContainer, options);
    expect(() => mountOfficeEditor(secondContainer, options)).toThrowError(
      expect.objectContaining({
        name: 'OfficeHostIsolationError',
        origin: 'https://shared.office-host.example.com',
        existingSessionId: first.id,
      }),
    );
    expect(firstContainer.querySelector('iframe')).not.toBeNull();
    expect(secondContainer.querySelector('iframe')).toBeNull();
    await first.destroy();
  });

  it('leases the twelve fixed constellation origins and reuses a released slot', async () => {
    const mounts = OFFICE_EDITOR_ORIGIN_SLOTS.map((expectedSlot, index) => {
      const container = document.createElement('div');
      document.body.appendChild(container);
      const mount = mountOfficeEditor(container, {
        hostUrl: ({ hostSlot }) => `https://${hostSlot}.getpi.work/office-host.html`,
        file: new File([String(index)], `document-${index}.docx`),
        destroyTimeoutMs: 1,
      });
      expect(new URL(container.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe(
        `${expectedSlot}.getpi.work`,
      );
      return mount;
    });
    const overflowContainer = document.createElement('div');
    document.body.appendChild(overflowContainer);

    expect(() =>
      mountOfficeEditor(overflowContainer, {
        hostUrl: ({ hostSlot }) => `https://${hostSlot}.getpi.work/office-host.html`,
        file: new File(['overflow'], 'overflow.docx'),
        destroyTimeoutMs: 1,
      }),
    ).toThrowError(
      expect.objectContaining({
        name: 'OfficeHostPoolExhaustedError',
        capacity: 12,
      }),
    );
    expect(overflowContainer.querySelector('iframe')).toBeNull();

    await mounts[0].destroy();
    const reusedContainer = document.createElement('div');
    document.body.appendChild(reusedContainer);
    const reused = mountOfficeEditor(reusedContainer, {
      hostUrl: ({ hostSlot }) => `https://${hostSlot}.getpi.work/office-host.html`,
      file: new File(['reused'], 'reused.docx'),
      destroyTimeoutMs: 1,
    });
    expect(new URL(reusedContainer.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe('aries.getpi.work');

    await Promise.all([...mounts.slice(1).map((mount) => mount.destroy()), reused.destroy()]);
  });

  it('prefers a persisted constellation slot and falls back when another document owns it', async () => {
    const firstContainer = document.createElement('div');
    const secondContainer = document.createElement('div');
    const thirdContainer = document.createElement('div');
    document.body.append(firstContainer, secondContainer, thirdContainer);
    const options = (fileName: string, preferredHostSlot?: (typeof OFFICE_EDITOR_ORIGIN_SLOTS)[number]) => ({
      hostUrl: ({ hostSlot }: { hostSlot: string }) => `https://${hostSlot}.getpi.work/office-host.html`,
      file: new File([fileName], fileName),
      fileName,
      preferredHostSlot,
      destroyTimeoutMs: 1,
    });

    const first = mountOfficeEditor(firstContainer, options('first.docx', 'aries'));
    const second = mountOfficeEditor(secondContainer, options('second.docx', 'aries'));
    expect(new URL(firstContainer.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe('aries.getpi.work');
    expect(new URL(secondContainer.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe('taurus.getpi.work');

    await first.destroy();
    const third = mountOfficeEditor(thirdContainer, options('third.docx', 'aries'));
    expect(new URL(thirdContainer.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe('aries.getpi.work');

    await Promise.all([second.destroy(), third.destroy()]);
  });

  it('keeps a closing origin retired until host teardown completes, then reuses it safely', async () => {
    const firstContainer = document.createElement('div');
    const secondContainer = document.createElement('div');
    document.body.append(firstContainer, secondContainer);
    const options = (fileName: string) => ({
      hostUrl: ({ hostSlot }: { hostSlot: string }) => `https://${hostSlot}.getpi.work/office-host.html`,
      file: new File([fileName], fileName),
      fileName,
      destroyTimeoutMs: 5_000,
    });

    const first = mountOfficeEditor(firstContainer, options('first.docx'));
    const second = mountOfficeEditor(secondContainer, options('second.docx'));
    const { iframe: firstFrame } = await connectHost(firstContainer);
    const firstOrigin = new URL(firstFrame.src).origin;
    const firstCompletion = first.destroy();

    const whileRetiringContainer = document.createElement('div');
    document.body.appendChild(whileRetiringContainer);
    const whileRetiring = mountOfficeEditor(whileRetiringContainer, options('while-retiring.docx'));
    expect(new URL(whileRetiringContainer.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe(
      'gemini.getpi.work',
    );

    const parentWindow = firstContainer.ownerDocument.defaultView || window;
    parentWindow.dispatchEvent(
      new parentWindow.MessageEvent('message', {
        origin: firstOrigin,
        source: firstFrame.contentWindow,
        data: {
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'HOST_RESET_DONE',
          sessionId: first.id,
        },
      }),
    );
    await flush();
    firstFrame.dispatchEvent(new Event('load'));
    await firstCompletion;

    const reusedContainer = document.createElement('div');
    document.body.appendChild(reusedContainer);
    const reused = mountOfficeEditor(reusedContainer, options('reused.docx'));
    expect(new URL(reusedContainer.querySelector<HTMLIFrameElement>('iframe')!.src).hostname).toBe('aries.getpi.work');

    await Promise.all([second.destroy(), whileRetiring.destroy(), reused.destroy()]);
  });

  it('does not leak an origin when retiring teardown reaches its bounded timeout', async () => {
    const firstContainer = document.createElement('div');
    const blockedContainer = document.createElement('div');
    document.body.append(firstContainer, blockedContainer);
    const options = {
      hostUrl: 'https://shared.office-host.example.com/office-host.html',
      file: new File(['timeout'], 'timeout.docx'),
      fileName: 'timeout.docx',
      destroyTimeoutMs: 1,
    };

    const first = mountOfficeEditor(firstContainer, options);
    await connectHost(firstContainer);
    vi.useFakeTimers();
    const completion = first.destroy();

    expect(() => mountOfficeEditor(blockedContainer, options)).toThrowError(
      expect.objectContaining({
        name: 'OfficeHostIsolationError',
        origin: 'https://shared.office-host.example.com',
        existingSessionId: first.id,
      }),
    );

    await vi.runAllTimersAsync();
    await completion;
    vi.useRealTimers();

    const reused = mountOfficeEditor(blockedContainer, options);
    expect(blockedContainer.querySelector('iframe')).not.toBeNull();
    await reused.destroy();
  });

  it('force-detaches and releases an origin when cooperative teardown messaging fails', async () => {
    const firstContainer = document.createElement('div');
    const blockedContainer = document.createElement('div');
    document.body.append(firstContainer, blockedContainer);
    const options = {
      hostUrl: 'https://failed-teardown.office-host.example.com/office-host.html',
      file: new File(['failure'], 'failure.docx'),
      fileName: 'failure.docx',
      destroyTimeoutMs: 1,
    };

    const first = mountOfficeEditor(firstContainer, options);
    await connectHost(firstContainer);
    vi.useFakeTimers();
    vi.spyOn(MessagePort.prototype, 'postMessage').mockImplementationOnce(() => {
      throw new Error('closed teardown port');
    });
    const completion = first.destroy();

    expect(() => mountOfficeEditor(blockedContainer, options)).toThrowError(
      expect.objectContaining({
        name: 'OfficeHostIsolationError',
        existingSessionId: first.id,
      }),
    );

    await vi.runAllTimersAsync();
    await expect(completion).resolves.toBeUndefined();
    expect(firstContainer.querySelector('iframe')).toBeNull();
    vi.useRealTimers();

    const reused = mountOfficeEditor(blockedContainer, options);
    expect(blockedContainer.querySelector('iframe')).not.toBeNull();
    await reused.destroy();
  });

  it('accepts an HTMLElement container from another window', async () => {
    const popupFrame = document.createElement('iframe');
    document.body.appendChild(popupFrame);
    const container = popupFrame.contentWindow!.document.createElement('div');
    popupFrame.contentWindow!.document.body.appendChild(container);

    expect(container).not.toBeInstanceOf(HTMLElement);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'popup.docx'),
      fileName: 'popup.docx',
      destroyTimeoutMs: 1,
    });
    await connectHost(container);
    const instance = await promise;

    expect(container.classList.contains('office-editor-host')).toBe(true);
    expect(container.querySelector('iframe')?.ownerDocument).toBe(popupFrame.contentWindow!.document);

    await instance.destroy();
  });

  it('round-trips save-to-new-format confirmation through the isolated host', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'legacy.doc'),
      fileName: 'legacy.doc',
      destroyTimeoutMs: 1,
    });
    const { messages } = await connectHost(container, (message, childPort) => {
      if (message.type !== 'CONFIRM_SAVE_TO_NEW_FORMAT') return;
      childPort.postMessage({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'CONFIRM_SAVE_TO_NEW_FORMAT_RESULT',
        sessionId: message.sessionId,
        requestId: message.requestId,
        confirmed: true,
      });
    });
    const instance = await promise;

    await expect(instance.confirmSaveToNewFormat({ dontshow: true })).resolves.toBe(true);
    expect(messages).toContainEqual(
      expect.objectContaining({
        type: 'CONFIRM_SAVE_TO_NEW_FORMAT',
        options: { dontshow: true },
      }),
    );

    await instance.destroy();
  });

  it('defaults the isolated host iframe to fill its container', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      destroyTimeoutMs: 1,
    });
    const { iframe } = await connectHost(container);
    const instance = await promise;

    expect(container.classList.contains('office-editor-host')).toBe(true);
    expect(container.style.width).toBe('100%');
    expect(container.style.height).toBe('100%');
    expect(container.style.minWidth).toBe('0px');
    expect(container.style.minHeight).toBe('0px');
    expect(iframe.className).toBe('office-editor-host-frame');
    expect(iframe.style.display).toBe('block');
    expect(iframe.style.width).toBe('100%');
    expect(iframe.style.height).toBe('100%');
    expect(iframe.style.minWidth).toBe('0px');
    expect(iframe.style.minHeight).toBe('0px');

    await instance.destroy();
  });

  it('hides the isolated host iframe before teardown navigation to avoid white-frame flashes', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      destroyTimeoutMs: 1,
    });
    const { iframe } = await connectHost(container);
    const instance = await promise;
    const sourceDescriptor = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'src');
    expect(sourceDescriptor?.get).toEqual(expect.any(Function));
    expect(sourceDescriptor?.set).toEqual(expect.any(Function));
    const assignments: Array<{
      src: string;
      visibility: string;
      opacity: string;
      pointerEvents: string;
      ariaHidden: string | null;
    }> = [];
    Object.defineProperty(iframe, 'src', {
      configurable: true,
      get() {
        return sourceDescriptor!.get!.call(this);
      },
      set(value: string) {
        assignments.push({
          src: value,
          visibility: iframe.style.visibility,
          opacity: iframe.style.opacity,
          pointerEvents: iframe.style.pointerEvents,
          ariaHidden: iframe.getAttribute('aria-hidden'),
        });
        sourceDescriptor!.set!.call(this, value);
        iframe.dispatchEvent(new Event('load'));
      },
    });

    await instance.destroy();

    expect(assignments.length).toBeGreaterThan(0);
    expect(
      assignments.every(
        (assignment) =>
          assignment.visibility === 'hidden' &&
          assignment.opacity === '0' &&
          assignment.pointerEvents === 'none' &&
          assignment.ariaHidden === 'true',
      ),
    ).toBe(true);
  });

  it('isolates localhost hostUrl with a reusable constellation slot origin', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: window.location.origin + '/office-host.html',
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      destroyTimeoutMs: 1,
    });
    const { iframe } = await connectHost(container);
    const instance = await promise;
    const iframeUrl = new URL(iframe.src);

    expect(iframeUrl.origin).not.toBe(window.location.origin);
    expect(iframeUrl.hostname).toBe('host-aries.office.localhost');
    expect(iframeUrl.hostname.endsWith('.localhost')).toBe(true);
    await instance.destroy();
  });

  it('preserves an explicitly selected local constellation slot', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: 'http://host-taurus.office.localhost/office-host.html',
      file: new File(['a'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      destroyTimeoutMs: 1,
    });
    const { iframe } = await connectHost(container);
    const instance = await promise;

    expect(new URL(iframe.src).hostname).toBe('host-taurus.office.localhost');
    await instance.destroy();
  });

  it('supports a hostUrl resolver for production wildcard host origins', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const hostUrl = vi.fn(
      ({ sessionId, fileName, fileType, mode }) =>
        `https://${sessionId}.office-host.example.com/office-host.html?name=${encodeURIComponent(fileName)}&type=${fileType}&mode=${mode}`,
    );

    const promise = createOfficeEditor(container, {
      hostUrl,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      destroyTimeoutMs: 1,
    });
    const { iframe } = await connectHost(container);
    const instance = await promise;
    const iframeUrl = new URL(iframe.src);

    expect(hostUrl).toHaveBeenCalledWith(
      expect.objectContaining({
        sessionId: expect.stringMatching(/^office-editor-/),
        hostSlot: 'aries',
        fileName: 'alpha.docx',
        fileType: 'docx',
        mode: 'preview',
      }),
    );
    expect(iframeUrl.hostname).toMatch(/^office-editor-.*\.office-host\.example\.com$/);
    expect(iframeUrl.searchParams.get('sessionId')).toBeNull();
    expect(new URLSearchParams(iframeUrl.hash.slice(1)).get('sessionId')).toBe(instance.id);
    await instance.destroy();
  });

  it('creates an isolated-origin host iframe and transfers document bytes over the port', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onReady = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      interfaceTheme: 'dark',
      canReturnToPreview: true,
      downloadedFonts: ['fonts/005.ttc', 'fonts/024.ttc'],
      onReady,
    });
    const { iframe, messages } = await connectHost(container);
    const instance = await promise;

    expect(iframe.getAttribute('sandbox')).toBeNull();
    expect(iframe.getAttribute('allow')).toBe('clipboard-read; clipboard-write; fullscreen');
    expect(iframe.getAttribute('referrerpolicy')).toBe('no-referrer');
    expect(messages[0]).toMatchObject({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'INIT',
      options: {
        fileName: 'alpha.docx',
        interfaceTheme: 'dark',
        canReturnToPreview: true,
        downloadedFonts: ['fonts/005.ttc', 'fonts/024.ttc'],
        spellcheck: false,
        source: {
          kind: 'buffer',
          fileName: 'alpha.docx',
          sourceKind: 'local-file',
        },
      },
    });
    expect(instance.getState()).toMatchObject({ fileName: 'alpha.docx', fileType: 'docx', status: 'ready' });
    expect(onReady).toHaveBeenCalledTimes(1);

    instance.setInterfaceTheme('light');
    await waitForMessage();
    expect(messages).toContainEqual(
      expect.objectContaining({
        type: 'SET_INTERFACE_THEME',
        interfaceTheme: 'light',
      }),
    );
  });

  it('forwards plugin configuration and resolves plugin operations over the host port', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onPluginReady = vi.fn();
    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'sheet.xlsx'),
      fileName: 'sheet.xlsx',
      plugins: {
        configUrls: ['/onlyoffice-plugin/config.json'],
        autostart: ['asc.test-plugin'],
      },
      onPluginReady,
    });
    const { childPort, messages } = await connectHost(container, (message, port) => {
      if (message.type !== 'INVOKE_PLUGIN') return;
      port.postMessage({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'PLUGIN_RESULT',
        sessionId: message.sessionId,
        requestId: message.requestId,
        pluginGuid: message.pluginGuid,
        ok: true,
        result: { values: [[42]] },
      });
    });
    const instance = await promise;
    expect(messages[0]).toMatchObject({
      type: 'INIT',
      options: {
        plugins: {
          configUrls: ['/onlyoffice-plugin/config.json'],
          autostart: ['asc.test-plugin'],
        },
      },
    });

    childPort.postMessage({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'PLUGIN_READY',
      sessionId: instance.id,
      pluginGuid: 'asc.test-plugin',
      editorType: 'cell',
    });
    await waitForMessage();
    expect(onPluginReady).toHaveBeenCalledWith('asc.test-plugin', 'cell', instance);
    await expect(instance.invokePlugin('asc.test-plugin', { type: 'get_range_values' })).resolves.toEqual({
      values: [[42]],
    });
  });

  it('rejects a plugin operation that never receives a host response', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'sheet.xlsx'),
      fileName: 'sheet.xlsx',
    });
    await connectHost(container);
    const instance = await promise;

    vi.useFakeTimers();
    const operation = instance.invokePlugin('asc.test-plugin', { type: 'get_range_values' });
    const rejection = expect(operation).rejects.toThrow('Office plugin operation timed out: asc.test-plugin');
    await vi.advanceTimersByTimeAsync(45_000);

    await rejection;
  });

  it('removes sandbox if an integration mutates the host iframe after mount', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      destroyTimeoutMs: 1,
    });
    const { iframe } = await connectHost(container);
    const instance = await promise;

    iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-modals');
    await waitForMessage();

    expect(iframe.getAttribute('sandbox')).toBeNull();
    await instance.destroy();
  });

  it('ignores spoofed host-ready messages from the wrong origin', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
    });
    const iframe = await waitForIframe(container);
    const postMessageSpy = vi.spyOn(iframe.contentWindow! as any, 'postMessage');
    const sessionId = getSessionId(iframe);

    window.dispatchEvent(
      new MessageEvent('message', {
        origin: 'http://malicious.example',
        source: iframe.contentWindow,
        data: {
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'HOST_READY',
          sessionId,
          identity: HOST_IDENTITY,
        },
      }),
    );
    await flush();
    expect(postMessageSpy).not.toHaveBeenCalled();

    await connectHost(container);
    await expect(promise).resolves.toMatchObject({ id: sessionId });
  });

  it('exposes the verified host identity after the ready handshake', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      expectedHostIdentity: HOST_IDENTITY,
      destroyTimeoutMs: 1,
    });
    await connectHost(container);
    const instance = await promise;

    expect(instance.getHostIdentity()).toEqual(HOST_IDENTITY);
    expect(instance.getHostIdentity()).not.toBe(HOST_IDENTITY);
    await instance.destroy();
  });

  it('rejects an incompatible host before transferring document bytes', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const expectedHostIdentity = { ...HOST_IDENTITY, hostBuildId: 'newer-host-build' };

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['sensitive'], 'alpha.docx'),
      fileName: 'alpha.docx',
      expectedHostIdentity,
      destroyTimeoutMs: 1,
    });
    const iframe = await waitForIframe(container);
    const parentWindow = container.ownerDocument.defaultView || window;
    const postMessage = vi.spyOn(iframe.contentWindow! as any, 'postMessage');
    parentWindow.dispatchEvent(
      new parentWindow.MessageEvent('message', {
        origin: new URL(iframe.src).origin,
        source: iframe.contentWindow,
        data: {
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'HOST_READY',
          sessionId: getSessionId(iframe),
          identity: HOST_IDENTITY,
        },
      }),
    );

    await expect(promise).rejects.toMatchObject({
      name: 'OfficeHostIdentityMismatchError',
      expected: expectedHostIdentity,
      actual: HOST_IDENTITY,
    });
    expect(postMessage).not.toHaveBeenCalled();
  });

  it('resolves save requests with transferred bytes and calls onSave once', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onSave = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onSave,
    });
    const { messages } = await connectHost(container, (message, childPort) => {
      if (message.type !== 'SAVE') return;
      const buffer = new Uint8Array([9, 8, 7]).buffer;
      childPort.postMessage(
        {
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'SAVE_RESULT',
          sessionId: message.sessionId,
          requestId: message.requestId,
          buffer,
          fileName: 'alpha.docx',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
        [buffer],
      );
    });
    const instance = await promise;

    await expect(instance.save('DOCX')).resolves.toMatchObject({ name: 'alpha.docx', size: 3 });
    expect(onSave).toHaveBeenCalledTimes(1);
    await waitForMessage();
    expect(messages).toContainEqual(
      expect.objectContaining({
        type: 'SAVE_ACK',
        requestId: expect.any(String),
        ok: true,
      }),
    );
  });

  it('handles native save results without a pending programmatic save request', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onSave = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onSave,
    });
    const { childPort, iframe, messages } = await connectHost(container);
    const instance = await promise;
    const buffer = new Uint8Array([4, 5, 6]).buffer;
    const requestId = `${getSessionId(iframe)}-native-save-1`;

    childPort.postMessage(
      {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'SAVE_RESULT',
        sessionId: getSessionId(iframe),
        requestId,
        buffer,
        fileName: 'alpha.docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      },
      [buffer],
    );
    await waitForMessage();

    expect(instance.getState().status).toBe('ready');
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(messages).toContainEqual(
      expect.objectContaining({
        type: 'SAVE_ACK',
        requestId,
        ok: true,
      }),
    );
  });

  it('downloads host-exported files from the parent page', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      destroyTimeoutMs: 1,
    });
    const { childPort, iframe } = await connectHost(container);
    const instance = await promise;
    const buffer = new Uint8Array([0x25, 0x50, 0x44, 0x46]).buffer;

    childPort.postMessage(
      {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'DOWNLOAD_RESULT',
        sessionId: getSessionId(iframe),
        buffer,
        fileName: 'alpha.pdf',
        mimeType: 'application/pdf',
      },
      [buffer],
    );
    await waitForMessage();

    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);
    expect(URL.createObjectURL).toHaveBeenCalledWith(expect.objectContaining({ name: 'alpha.pdf' }));

    await instance.destroy();
  });

  it('allows hosts to intercept Download As files without saving them', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onDownload = vi.fn();
    const onSave = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onDownload,
      onSave,
      destroyTimeoutMs: 1,
    });
    const { childPort, iframe } = await connectHost(container);
    const instance = await promise;
    const buffer = new Uint8Array([0x25, 0x50, 0x44, 0x46]).buffer;

    childPort.postMessage(
      {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'DOWNLOAD_RESULT',
        sessionId: getSessionId(iframe),
        buffer,
        fileName: 'alpha.pdf',
        mimeType: 'application/pdf',
      },
      [buffer],
    );
    await waitForMessage();

    expect(onDownload).toHaveBeenCalledTimes(1);
    expect(onDownload.mock.calls[0][0]).toMatchObject({ name: 'alpha.pdf' });
    expect(onSave).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('routes host Save Copy As files to onSaveAs instead of onDownload', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onDownload = vi.fn();
    const onSave = vi.fn();
    const onSaveAs = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onDownload,
      onSave,
      onSaveAs,
      destroyTimeoutMs: 1,
    });
    const { childPort, iframe } = await connectHost(container);
    const instance = await promise;
    const buffer = new Uint8Array([0x50, 0x4b, 0x03, 0x04]).buffer;

    childPort.postMessage(
      {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'SAVE_AS_RESULT',
        sessionId: getSessionId(iframe),
        buffer,
        fileName: 'alpha.odt',
        mimeType: 'application/vnd.oasis.opendocument.text',
      },
      [buffer],
    );
    await waitForMessage();

    expect(onSaveAs).toHaveBeenCalledTimes(1);
    expect(onSaveAs.mock.calls[0][0]).toMatchObject({ name: 'alpha.odt' });
    expect(onDownload).not.toHaveBeenCalled();
    expect(onSave).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('reports dirty state changes from the host', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onDirtyChange = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onDirtyChange,
    });
    const { childPort, iframe } = await connectHost(container);
    const instance = await promise;
    const sessionId = getSessionId(iframe);

    childPort.postMessage({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'STATE',
      sessionId,
      state: {
        id: sessionId,
        fileName: 'alpha.docx',
        fileType: 'docx',
        mode: 'edit',
        readonly: false,
        dirty: true,
        sourceKind: 'local-file',
        status: 'ready',
        destroyed: false,
      },
    });
    await waitForMessage();

    expect(instance.getState().dirty).toBe(true);
    expect(onDirtyChange).toHaveBeenCalledWith(true, instance);
  });

  it('rejects save requests when the outer onSave write fails and keeps the document dirty', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onSave = vi.fn(async () => {
      throw new Error('write failed');
    });

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onSave,
    });
    const { childPort, iframe, messages } = await connectHost(container, (message, port) => {
      if (message.type !== 'SAVE') return;
      const buffer = new Uint8Array([9, 8, 7]).buffer;
      port.postMessage(
        {
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'SAVE_RESULT',
          sessionId: message.sessionId,
          requestId: message.requestId,
          buffer,
          fileName: 'alpha.docx',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
        [buffer],
      );
    });
    const instance = await promise;
    const sessionId = getSessionId(iframe);
    childPort.postMessage({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'STATE',
      sessionId,
      state: {
        id: sessionId,
        fileName: 'alpha.docx',
        fileType: 'docx',
        mode: 'edit',
        readonly: false,
        dirty: true,
        sourceKind: 'local-file',
        status: 'ready',
        destroyed: false,
      },
    });
    await waitForMessage();

    await expect(instance.save('DOCX')).rejects.toThrow('write failed');
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(instance.getState().dirty).toBe(true);
    await waitForMessage();
    expect(messages).toContainEqual(
      expect.objectContaining({
        type: 'SAVE_ACK',
        requestId: expect.any(String),
        ok: false,
        message: 'write failed',
      }),
    );
  });

  it('destroys idempotently and force-removes the host iframe without an ack', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      destroyTimeoutMs: 1,
    });
    await connectHost(container);
    const instance = await promise;

    await Promise.all([instance.destroy(), instance.destroy()]);

    expect(container.querySelector('iframe')).toBeNull();
    expect(instance.getState()).toMatchObject({ destroyed: true, status: 'destroyed' });
  });
});
