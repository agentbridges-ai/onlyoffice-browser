import { beforeEach, describe, expect, it, vi } from 'vitest';

import { OFFICE_HOST_PROTOCOL, type OfficeHostParentMessage } from '../../src/lib/office-host-protocol';
import { createOfficeEditor, loadOfficeEditorApi } from '../../src/lib/office-editor';

const HOST_URL = 'http://127.0.0.1:5173/office-host.html';

function flush(): Promise<void> {
  return Promise.resolve().then(() => undefined);
}

function waitForMessage(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function getSessionId(iframe: HTMLIFrameElement): string {
  return new URL(iframe.src).searchParams.get('sessionId') || '';
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

  vi.spyOn(iframe.contentWindow! as any, 'postMessage').mockImplementation(
    (...args: unknown[]) => {
      const transfer = args[2] as Transferable[] | undefined;
    childPort = transfer?.[0] as MessagePort;
    childPort.onmessage = (event: MessageEvent<OfficeHostParentMessage>) => {
      messages.push(event.data);
      onChildMessage?.(event.data, childPort!);
      if (event.data.type === 'INIT') {
        const sourceKind = event.data.options.source.kind === 'empty' ? 'new-document' : event.data.options.source.sourceKind;
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
    },
  );

  parentWindow.dispatchEvent(
    new parentWindow.MessageEvent('message', {
      origin: hostOrigin,
      source: iframe.contentWindow,
      data: {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'HOST_READY',
        sessionId,
      },
    }),
  );

  await flush();
  return { childPort: childPort!, iframe, messages };
}

describe('office-editor parent proxy', () => {
  beforeEach(() => {
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
    expect(messages).toContainEqual(expect.objectContaining({
      type: 'CONFIRM_SAVE_TO_NEW_FORMAT',
      options: { dontshow: true },
    }));

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
    expect(assignments.every((assignment) => (
      assignment.visibility === 'hidden' &&
      assignment.opacity === '0' &&
      assignment.pointerEvents === 'none' &&
      assignment.ariaHidden === 'true'
    ))).toBe(true);
  });

  it('isolates localhost hostUrl to a per-editor origin', async () => {
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
    expect(iframeUrl.hostname).toMatch(/^host-office-editor-/);
    expect(iframeUrl.hostname.endsWith('.localhost')).toBe(true);
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
        fileName: 'alpha.docx',
        fileType: 'docx',
        mode: 'preview',
      }),
    );
    expect(iframeUrl.hostname).toMatch(/^office-editor-.*\.office-host\.example\.com$/);
    expect(iframeUrl.searchParams.get('sessionId')).toBe(instance.id);
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
    expect(messages).toContainEqual(expect.objectContaining({
      type: 'SET_INTERFACE_THEME',
      interfaceTheme: 'light',
    }));
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
        },
      }),
    );
    await flush();
    expect(postMessageSpy).not.toHaveBeenCalled();

    await connectHost(container);
    await expect(promise).resolves.toMatchObject({ id: sessionId });
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
    expect(messages).toContainEqual(expect.objectContaining({
      type: 'SAVE_ACK',
      requestId: expect.any(String),
      ok: true,
    }));
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
    expect(messages).toContainEqual(expect.objectContaining({
      type: 'SAVE_ACK',
      requestId,
      ok: true,
    }));
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
    expect(messages).toContainEqual(expect.objectContaining({
      type: 'SAVE_ACK',
      requestId: expect.any(String),
      ok: false,
      message: 'write failed',
    }));
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
