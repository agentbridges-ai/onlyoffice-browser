import { createOfficeEditor as createRuntimeOfficeEditor, type OfficeEditorInstance } from './lib/office-editor-runtime';
import {
  isOfficeHostMessage,
  OFFICE_HOST_PROTOCOL,
  type OfficeHostChildMessage,
  type OfficeHostParentMessage,
  type OfficeHostState,
  type OfficeHostWindowMessage,
} from './lib/office-host-protocol';
import './styles/base.css';

type RuntimeOptions = Parameters<typeof createRuntimeOfficeEditor>[1];

const params = new URLSearchParams(window.location.search);
const sessionId = params.get('sessionId') || '';
const parentOrigin = params.get('parentOrigin') || '';
const root = document.querySelector<HTMLElement>('#office-host') ?? document.body;
const HOST_RESET_PATH = '/reset.html';

let port: MessagePort | null = null;
let editor: OfficeEditorInstance | null = null;
let destroyed = false;
let activeSaveRequestId: string | undefined;

function toError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}

function postWindowMessage(type: OfficeHostWindowMessage['type']): void {
  if (!parentOrigin || !sessionId) return;
  window.parent.postMessage(
    {
      protocol: OFFICE_HOST_PROTOCOL,
      type,
      sessionId,
    } satisfies OfficeHostWindowMessage,
    parentOrigin,
  );
}

function getHostResetUrl(): string {
  const url = new URL(HOST_RESET_PATH, window.location.href);
  url.searchParams.set('stay', '1');
  url.searchParams.set('officeHostReset', '1');
  url.searchParams.set('sessionId', sessionId);
  url.searchParams.set('parentOrigin', parentOrigin);
  return url.href;
}

function hardResetHostPage(): void {
  window.location.replace(getHostResetUrl());
}

function postPortMessage(message: OfficeHostChildMessage, transfer: Transferable[] = []): void {
  port?.postMessage(message, transfer);
}

function postError(
  phase: Extract<OfficeHostChildMessage, { type: 'ERROR' }>['phase'],
  error: unknown,
  requestId?: string,
): void {
  postPortMessage({
    protocol: OFFICE_HOST_PROTOCOL,
    type: 'ERROR',
    sessionId,
    requestId,
    phase,
    message: toError(error).message,
  });
}

async function postSavedFile(file: File, requestId?: string): Promise<void> {
  const buffer = await file.arrayBuffer();
  postPortMessage(
    {
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'SAVE_RESULT',
      sessionId,
      requestId,
      buffer,
      fileName: file.name,
      mimeType: file.type,
    },
    [buffer],
  );
}

function postState(type: 'READY' | 'STATE', state: OfficeHostState): void {
  postPortMessage({
    protocol: OFFICE_HOST_PROTOCOL,
    type,
    sessionId,
    state,
  });
}

function createResourceTracker() {
  const originalSetTimeout = window.setTimeout.bind(window);
  const originalClearTimeout = window.clearTimeout.bind(window);
  const originalSetInterval = window.setInterval.bind(window);
  const originalClearInterval = window.clearInterval.bind(window);
  const originalRequestAnimationFrame = window.requestAnimationFrame?.bind(window);
  const originalCancelAnimationFrame = window.cancelAnimationFrame?.bind(window);
  const originalRequestIdleCallback = window.requestIdleCallback?.bind(window);
  const originalCancelIdleCallback = window.cancelIdleCallback?.bind(window);
  const OriginalWorker = window.Worker;
  const OriginalWebSocket = window.WebSocket;
  const OriginalMessageChannel = window.MessageChannel;
  const OriginalAudioContext = window.AudioContext;
  const OriginalWebkitAudioContext = (window as Window & { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;

  const timeouts = new Set<number>();
  const intervals = new Set<number>();
  const animationFrames = new Set<number>();
  const idleCallbacks = new Set<number>();
  const workers = new Set<Worker>();
  const sockets = new Set<WebSocket>();
  const messagePorts = new Set<MessagePort>();
  const audioContexts = new Set<AudioContext>();

  window.setTimeout = ((handler: TimerHandler, timeout?: number, ...args: unknown[]) => {
    const id = originalSetTimeout(handler, timeout, ...args);
    timeouts.add(id);
    return id;
  }) as typeof window.setTimeout;

  window.clearTimeout = ((id?: number) => {
    if (typeof id === 'number') timeouts.delete(id);
    return originalClearTimeout(id);
  }) as typeof window.clearTimeout;

  window.setInterval = ((handler: TimerHandler, timeout?: number, ...args: unknown[]) => {
    const id = originalSetInterval(handler, timeout, ...args);
    intervals.add(id);
    return id;
  }) as typeof window.setInterval;

  window.clearInterval = ((id?: number) => {
    if (typeof id === 'number') intervals.delete(id);
    return originalClearInterval(id);
  }) as typeof window.clearInterval;

  if (originalRequestAnimationFrame && originalCancelAnimationFrame) {
    window.requestAnimationFrame = ((callback: FrameRequestCallback) => {
      const id = originalRequestAnimationFrame(callback);
      animationFrames.add(id);
      return id;
    }) as typeof window.requestAnimationFrame;

    window.cancelAnimationFrame = ((id: number) => {
      animationFrames.delete(id);
      return originalCancelAnimationFrame(id);
    }) as typeof window.cancelAnimationFrame;
  }

  if (originalRequestIdleCallback && originalCancelIdleCallback) {
    window.requestIdleCallback = ((callback: IdleRequestCallback, options?: IdleRequestOptions) => {
      const id = originalRequestIdleCallback(callback, options);
      idleCallbacks.add(id);
      return id;
    }) as typeof window.requestIdleCallback;

    window.cancelIdleCallback = ((id: number) => {
      idleCallbacks.delete(id);
      return originalCancelIdleCallback(id);
    }) as typeof window.cancelIdleCallback;
  }

  if (OriginalWorker) {
    const TrackedWorker = function (this: Worker, ...args: ConstructorParameters<typeof Worker>) {
      const worker = new OriginalWorker(...args);
      workers.add(worker);
      const terminate = worker.terminate.bind(worker);
      worker.terminate = () => {
        workers.delete(worker);
        terminate();
      };
      return worker;
    } as unknown as typeof Worker;
    TrackedWorker.prototype = OriginalWorker.prototype;
    window.Worker = TrackedWorker;
  }

  if (OriginalWebSocket) {
    const TrackedWebSocket = function (this: WebSocket, ...args: ConstructorParameters<typeof WebSocket>) {
      const socket = new OriginalWebSocket(...args);
      sockets.add(socket);
      socket.addEventListener('close', () => sockets.delete(socket), { once: true });
      return socket;
    } as unknown as typeof WebSocket;
    TrackedWebSocket.prototype = OriginalWebSocket.prototype;
    window.WebSocket = TrackedWebSocket;
  }

  if (OriginalMessageChannel) {
    const TrackedMessageChannel = function (this: MessageChannel) {
      const channel = new OriginalMessageChannel();
      messagePorts.add(channel.port1);
      messagePorts.add(channel.port2);
      return channel;
    } as unknown as typeof MessageChannel;
    TrackedMessageChannel.prototype = OriginalMessageChannel.prototype;
    window.MessageChannel = TrackedMessageChannel;
  }

  if (OriginalAudioContext) {
    const TrackedAudioContext = function (this: AudioContext, ...args: ConstructorParameters<typeof AudioContext>) {
      const context = new OriginalAudioContext(...args);
      audioContexts.add(context);
      return context;
    } as unknown as typeof AudioContext;
    TrackedAudioContext.prototype = OriginalAudioContext.prototype;
    window.AudioContext = TrackedAudioContext;
  }

  if (OriginalWebkitAudioContext) {
    const TrackedWebkitAudioContext = function (
      this: AudioContext,
      ...args: ConstructorParameters<typeof AudioContext>
    ) {
      const context = new OriginalWebkitAudioContext(...args);
      audioContexts.add(context);
      return context;
    } as unknown as typeof AudioContext;
    TrackedWebkitAudioContext.prototype = OriginalWebkitAudioContext.prototype;
    (window as Window & { webkitAudioContext?: typeof AudioContext }).webkitAudioContext = TrackedWebkitAudioContext;
  }

  return {
    cleanup() {
      for (const id of timeouts) originalClearTimeout(id);
      timeouts.clear();
      for (const id of intervals) originalClearInterval(id);
      intervals.clear();
      if (originalCancelAnimationFrame) {
        for (const id of animationFrames) originalCancelAnimationFrame(id);
      }
      animationFrames.clear();
      if (originalCancelIdleCallback) {
        for (const id of idleCallbacks) originalCancelIdleCallback(id);
      }
      idleCallbacks.clear();
      for (const worker of workers) worker.terminate();
      workers.clear();
      for (const socket of sockets) {
        try {
          socket.close();
        } catch {
          // Ignore best-effort shutdown failures.
        }
      }
      sockets.clear();
      for (const trackedPort of messagePorts) trackedPort.close();
      messagePorts.clear();
      for (const context of audioContexts) {
        void context.close().catch(() => {});
      }
      audioContexts.clear();
      for (const media of document.querySelectorAll<HTMLMediaElement>('audio, video')) {
        media.pause();
        media.removeAttribute('src');
        media.load();
      }
    },
  };
}

const resources = createResourceTracker();

async function blankAndRemoveEditorIframes(): Promise<void> {
  const frames = Array.from(document.querySelectorAll<HTMLIFrameElement>('iframe'));
  await Promise.all(
    frames.map(
      (frame) =>
        new Promise<void>((resolve) => {
          cleanupFrameWindow(frame);
          if (!frame.isConnected) {
            resolve();
            return;
          }
          const timeout = window.setTimeout(resolve, 250);
          frame.addEventListener(
            'load',
            () => {
              window.clearTimeout(timeout);
              resolve();
            },
            { once: true },
          );
          frame.src = 'about:blank';
        }).then(() => frame.remove()),
    ),
  );
}

function cleanupFrameWindow(frame: HTMLIFrameElement): void {
  const frameWindow = frame.contentWindow;
  if (!frameWindow) return;

  try {
    for (const childFrame of Array.from(frameWindow.document.querySelectorAll<HTMLIFrameElement>('iframe'))) {
      cleanupFrameWindow(childFrame);
    }
  } catch {
    // Ignore inaccessible frame windows.
  }

  try {
    for (let id = 0; id < 10_000; id += 1) {
      frameWindow.clearTimeout(id);
      frameWindow.clearInterval(id);
      frameWindow.cancelAnimationFrame?.(id);
      frameWindow.cancelIdleCallback?.(id);
    }
  } catch {
    // Ignore frames that are already unloading.
  }

  try {
    for (const media of Array.from(frameWindow.document.querySelectorAll<HTMLMediaElement>('audio, video'))) {
      media.pause();
      media.removeAttribute('src');
      media.load();
    }
    frameWindow.document.body?.replaceChildren();
  } catch {
    // Ignore DOMs that are already gone.
  }
}

async function destroyRuntime(): Promise<void> {
  if (destroyed) return;
  destroyed = true;

  try {
    await editor?.destroy();
  } catch (error) {
    console.warn('Failed to destroy runtime editor:', error);
  }
  editor = null;
  await blankAndRemoveEditorIframes();
  resources.cleanup();
  await clearHostWorkersAndCaches();
  root.replaceChildren();
  delete window.APP;
}

async function clearHostWorkersAndCaches(): Promise<void> {
  await Promise.all([
    navigator.serviceWorker
      ?.getRegistrations()
      .then((registrations) => Promise.all(registrations.map((registration) => registration.unregister())))
      .catch(() => undefined),
    window.caches
      ?.keys()
      .then((keys) => Promise.all(keys.map((key) => window.caches.delete(key))))
      .catch(() => undefined),
  ]);
}

async function handleInit(message: Extract<OfficeHostParentMessage, { type: 'INIT' }>): Promise<void> {
  if (destroyed) return;

  try {
    const { source } = message.options;
    const runtimeOptions: RuntimeOptions = {
      fileName: message.options.fileName,
      mode: message.options.mode,
      readonly: message.options.readonly,
      spellcheck: message.options.spellcheck ?? false,
      lang: message.options.lang,
      onReady: (instance) => {
        postState('STATE', instance.getState());
      },
      onError: (error) => {
        postError('runtime', error);
      },
      onSave: (file) => postSavedFile(file, activeSaveRequestId),
    };

    if (source.kind === 'empty') {
      runtimeOptions.emptyType = source.emptyType;
    } else {
      runtimeOptions.buffer = source.buffer;
      runtimeOptions.fileName = source.fileName;
    }

    editor = await createRuntimeOfficeEditor(root, runtimeOptions);
    postState('READY', editor.getState());
  } catch (error) {
    postError('init', error, message.requestId);
  }
}

async function handleSave(message: Extract<OfficeHostParentMessage, { type: 'SAVE' }>): Promise<void> {
  try {
    if (!editor) {
      throw new Error('Editor is not open');
    }
    activeSaveRequestId = message.requestId;
    await editor.save(message.targetExt);
  } catch (error) {
    postError('save', error, message.requestId);
  } finally {
    if (activeSaveRequestId === message.requestId) activeSaveRequestId = undefined;
  }
}

function handleSetReadonly(message: Extract<OfficeHostParentMessage, { type: 'SET_READONLY' }>): void {
  try {
    if (!editor) {
      throw new Error('Editor is not open');
    }
    editor.setReadonly(message.readonly);
    postState('STATE', editor.getState());
  } catch (error) {
    postError('setReadonly', error, message.requestId);
  }
}

async function handleDestroy(): Promise<void> {
  await destroyRuntime();
  const currentPort = port;
  port = null;
  if (currentPort) {
    currentPort.onmessage = null;
    currentPort.close();
  }
  window.setTimeout(hardResetHostPage, 0);
}

function handlePortMessage(event: MessageEvent<OfficeHostParentMessage>): void {
  if (!isOfficeHostMessage(event.data, sessionId)) return;

  switch (event.data.type) {
    case 'INIT':
      void handleInit(event.data);
      return;
    case 'SAVE':
      void handleSave(event.data);
      return;
    case 'SET_READONLY':
      handleSetReadonly(event.data);
      return;
    case 'DESTROY':
      void handleDestroy();
      return;
  }
}

function handleConnect(event: MessageEvent): void {
  if (!parentOrigin || !sessionId || event.origin !== parentOrigin || event.source !== window.parent) {
    return;
  }
  if (!isOfficeHostMessage(event.data, sessionId)) {
    return;
  }
  const message = event.data as OfficeHostWindowMessage;
  if (message.type !== 'CONNECT' || !event.ports[0]) {
    return;
  }

  window.removeEventListener('message', handleConnect);
  port = event.ports[0];
  port.onmessage = handlePortMessage;
  port.start();
}

window.addEventListener('message', handleConnect);
window.addEventListener('pagehide', () => {
  void destroyRuntime();
});
window.addEventListener('unload', () => undefined);
postWindowMessage('HOST_READY');
