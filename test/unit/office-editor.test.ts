import { beforeEach, describe, expect, it, vi } from 'vitest';

import { OFFICE_HOST_PROTOCOL, type OfficeHostParentMessage } from '../../src/lib/office-host-protocol';
import { createOfficeEditor, loadOfficeEditorApi } from '../../src/lib/office-editor';

const HOST_URL = 'http://127.0.0.1:5173/office-host.html';

function flush(): Promise<void> {
  return Promise.resolve().then(() => undefined);
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
            status: 'ready',
            destroyed: false,
          },
        });
      }
    };
      childPort.start();
    },
  );

  window.dispatchEvent(
    new MessageEvent('message', {
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
  });

  it('keeps loadOfficeEditorApi as a parent-side no-op', async () => {
    await expect(loadOfficeEditorApi()).resolves.toBeUndefined();
    expect(document.querySelectorAll('script[data-office-editor-api="true"]')).toHaveLength(0);
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

  it('creates a sandboxed host iframe and transfers document bytes over the port', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onReady = vi.fn();

    const promise = createOfficeEditor(container, {
      hostUrl: HOST_URL,
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      onReady,
    });
    const { iframe, messages } = await connectHost(container);
    const instance = await promise;

    expect(iframe.getAttribute('sandbox')).toBe(
      'allow-scripts allow-same-origin allow-forms allow-modals allow-downloads allow-popups',
    );
    expect(iframe.getAttribute('allow')).toBe('clipboard-read; clipboard-write; fullscreen');
    expect(iframe.getAttribute('referrerpolicy')).toBe('no-referrer');
    expect(messages[0]).toMatchObject({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'INIT',
      options: {
        fileName: 'alpha.docx',
        source: {
          kind: 'buffer',
          fileName: 'alpha.docx',
        },
      },
    });
    expect(instance.getState()).toMatchObject({ fileName: 'alpha.docx', fileType: 'docx', status: 'ready' });
    expect(onReady).toHaveBeenCalledTimes(1);
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
    await connectHost(container, (message, childPort) => {
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
