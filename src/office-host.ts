import {
  createOfficeEditor as createRuntimeOfficeEditor,
  type OfficeEditorInstance,
} from './lib/office-editor-runtime';
import {
  isOfficeHostMessage,
  OFFICE_HOST_PROTOCOL,
  type OfficeHostChildMessage,
  type OfficeHostParentMessage,
  type OfficeHostStartupPhase,
  type OfficeHostState,
  type OfficeHostWindowMessage,
} from './lib/office-host-protocol';
import {
  isOfficePluginResultForRuntime,
  resolveOfficePluginReady,
  type OfficePluginRuntime,
} from './lib/office-plugin-runtime';
import './styles/base.css';

type RuntimeOptions = Parameters<typeof createRuntimeOfficeEditor>[1];

const params = new URLSearchParams(window.location.search);
const sessionId = params.get('sessionId') || '';
const parentOrigin = params.get('parentOrigin') || '';
const root = document.querySelector<HTMLElement>('#office-host') ?? document.body;
const HOST_RESET_PATH = '/reset.html';
const SAVE_ACK_TIMEOUT_MS = 60_000;
/** Bump whenever already-open host frames must be recreated. */
const OFFICE_BROWSER_PACKAGE_VERSION = '0.3.34';
const OFFICE_HOST_BUILD_ID = 'office-host-0.3.34-r1';
const OFFICE_RUNTIME_ASSET_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const STARTUP_HEARTBEAT_INTERVAL_MS = 5_000;
const PLUGIN_REQUEST_TIMEOUT_MS = 30_000;
const OFFICE_PLUGIN_PROTOCOL = 'onlyoffice-browser-plugin/v1';

let port: MessagePort | null = null;
let editor: OfficeEditorInstance | null = null;
let destroyed = false;
let initStarted = false;
let startupPhase: OfficeHostStartupPhase = 'connected';
let startupHeartbeatInterval: number | null = null;
let startupHeartbeatWorker: Worker | null = null;
let activeSaveRequestId: string | undefined;
let nativeSaveRequestSequence = 0;
const pluginRuntimes = new Map<string, OfficePluginRuntime<Window>>();
const configuredPluginGuids = new Set<string>();
const pendingPluginRequests = new Map<string, { runtime: OfficePluginRuntime<Window>; timeoutId: number }>();

type OfficePluginWindowMessage = {
  protocol: typeof OFFICE_PLUGIN_PROTOCOL;
  type: 'READY' | 'RESULT';
  pluginGuid: string;
  pluginInstanceId: string;
  editorType?: string;
  requestId?: string;
  ok?: boolean;
  result?: unknown;
  error?: string;
};

type PendingSaveAck = {
  resolve: () => void;
  reject: (error: Error) => void;
  timeoutId: number;
};

const pendingSaveAcks = new Map<string, PendingSaveAck>();

type PrintTitleHostWindow = typeof window & {
  __onlyOfficeBrowserSetPrintTitle?: (title: string, durationMs?: number) => void;
};

function toError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}

function postWindowMessage(message: OfficeHostWindowMessage): void {
  if (!parentOrigin || !sessionId) return;
  window.parent.postMessage(message, parentOrigin);
}

function bytesToHex(bytes: ArrayBuffer): string {
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function loadOfficeHostIdentity() {
  const response = await fetch(OFFICE_RUNTIME_ASSET_MANIFEST_PATH, {
    cache: 'no-store',
    credentials: 'omit',
  });
  if (!response.ok) {
    throw new Error(`Unable to load compact Office runtime manifest (${response.status})`);
  }
  const manifest = await response.arrayBuffer();
  return {
    packageVersion: OFFICE_BROWSER_PACKAGE_VERSION,
    hostBuildId: OFFICE_HOST_BUILD_ID,
    assetManifestDigest: bytesToHex(await crypto.subtle.digest('SHA-256', manifest)),
  };
}

async function announceHostReady(): Promise<void> {
  try {
    const identity = await loadOfficeHostIdentity();
    postWindowMessage({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'HOST_READY',
      sessionId,
      identity,
    });
  } catch (error) {
    // A host whose assets cannot be identified must fail closed. The parent
    // times out without ever transferring document bytes into this frame.
    console.error('[onlyoffice-browser] Failed to identify Office host runtime', error);
  }
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

function postStartupMessage(type: 'STARTUP_PHASE' | 'STARTUP_HEARTBEAT'): void {
  postPortMessage({
    protocol: OFFICE_HOST_PROTOCOL,
    type,
    sessionId,
    phase: startupPhase,
  });
}

function setStartupPhase(phase: OfficeHostStartupPhase): void {
  startupPhase = phase;
  postStartupMessage('STARTUP_PHASE');
  startupHeartbeatWorker?.postMessage({ type: 'PHASE', phase });
}

function startStartupHeartbeat(): void {
  if (startupHeartbeatWorker || startupHeartbeatInterval !== null) return;
  if (typeof Worker !== 'undefined' && typeof MessageChannel !== 'undefined') {
    try {
      const worker = new Worker(new URL('./startup-heartbeat-worker.ts', import.meta.url), { type: 'module' });
      const channel = new MessageChannel();
      startupHeartbeatWorker = worker;
      postPortMessage(
        {
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'STARTUP_HEARTBEAT_PORT',
          sessionId,
        },
        [channel.port2],
      );
      worker.postMessage(
        {
          type: 'START',
          sessionId,
          phase: startupPhase,
          intervalMs: STARTUP_HEARTBEAT_INTERVAL_MS,
          port: channel.port1,
        },
        [channel.port1],
      );
      return;
    } catch (error) {
      startupHeartbeatWorker?.terminate();
      startupHeartbeatWorker = null;
      console.warn('[onlyoffice-browser] Falling back to main-thread startup heartbeat', error);
    }
  }
  postStartupMessage('STARTUP_HEARTBEAT');
  startupHeartbeatInterval = window.setInterval(() => {
    postStartupMessage('STARTUP_HEARTBEAT');
  }, STARTUP_HEARTBEAT_INTERVAL_MS);
}

function stopStartupHeartbeat(): void {
  if (startupHeartbeatWorker) {
    startupHeartbeatWorker.postMessage({ type: 'STOP' });
    startupHeartbeatWorker.terminate();
    startupHeartbeatWorker = null;
  }
  if (startupHeartbeatInterval === null) return;
  window.clearInterval(startupHeartbeatInterval);
  startupHeartbeatInterval = null;
}

function nextNativeSaveRequestId(): string {
  nativeSaveRequestSequence += 1;
  return `${sessionId}-native-save-${Date.now()}-${nativeSaveRequestSequence}`;
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

function makeSaveAckError(message: string): Error {
  const error = new Error(message);
  error.name = 'OnlyOfficeHostSaveAckError';
  return error;
}

function waitForSaveAck(requestId: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      pendingSaveAcks.delete(requestId);
      reject(makeSaveAckError('Timed out waiting for parent save acknowledgement'));
    }, SAVE_ACK_TIMEOUT_MS);
    pendingSaveAcks.set(requestId, { resolve, reject, timeoutId });
  });
}

async function postSavedFile(file: File, requestId = nextNativeSaveRequestId()): Promise<void> {
  const buffer = await file.arrayBuffer();
  const ack = waitForSaveAck(requestId);
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
  await ack;
}

async function postDownloadedFile(file: File): Promise<void> {
  const buffer = await file.arrayBuffer();
  postPortMessage(
    {
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'DOWNLOAD_RESULT',
      sessionId,
      buffer,
      fileName: file.name,
      mimeType: file.type,
    },
    [buffer],
  );
}

async function postSaveAsFile(file: File): Promise<void> {
  const buffer = await file.arrayBuffer();
  postPortMessage(
    {
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'SAVE_AS_RESULT',
      sessionId,
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
  const OriginalWebkitAudioContext = (window as Window & { webkitAudioContext?: typeof AudioContext })
    .webkitAudioContext;

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

function hideIframeForTeardown(frame: HTMLIFrameElement): void {
  frame.setAttribute('aria-hidden', 'true');
  frame.style.visibility = 'hidden';
  frame.style.opacity = '0';
  frame.style.pointerEvents = 'none';
  frame.style.background = 'transparent';
}

async function blankAndRemoveEditorIframes(): Promise<void> {
  const frames = Array.from(document.querySelectorAll<HTMLIFrameElement>('iframe'));
  await Promise.all(
    frames.map((frame) =>
      new Promise<void>((resolve) => {
        hideIframeForTeardown(frame);
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
  stopStartupHeartbeat();

  for (const [requestId, pending] of pendingSaveAcks) {
    window.clearTimeout(pending.timeoutId);
    pending.reject(makeSaveAckError(`Editor was destroyed before save acknowledgement completed: ${requestId}`));
  }
  pendingSaveAcks.clear();
  rejectPendingPluginRequests(undefined, 'Editor was destroyed before plugin operation completed');
  pluginRuntimes.clear();

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
  delete (window as PrintTitleHostWindow).__onlyOfficeBrowserSetPrintTitle;
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
  if (initStarted || editor) {
    postError('init', new Error('Duplicate INIT is not allowed for an Office host session'), message.requestId);
    return;
  }
  initStarted = true;
  configuredPluginGuids.clear();
  for (const guid of message.options.plugins?.autostart ?? []) configuredPluginGuids.add(guid);
  setStartupPhase('reading-source');
  startStartupHeartbeat();

  try {
    (window as PrintTitleHostWindow).__onlyOfficeBrowserSetPrintTitle = (title, durationMs = 45_000) => {
      postPortMessage({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'PRINT_TITLE',
        sessionId,
        title,
        durationMs,
      });
    };

    const { source } = message.options;
    const runtimeOptions: RuntimeOptions = {
      fileName: message.options.fileName,
      mode: message.options.mode,
      readonly: message.options.readonly,
      canReturnToPreview: message.options.canReturnToPreview,
      spellcheck: message.options.spellcheck ?? false,
      interfaceTheme: message.options.interfaceTheme,
      lang: message.options.lang,
      plugins: message.options.plugins,
      saveBehavior: message.options.saveBehavior,
      onReady: (instance) => {
        postState('STATE', instance.getState());
      },
      onDirtyChange: (_dirty, instance) => {
        postState('STATE', instance.getState());
      },
      onStateChange: (state) => {
        postState('STATE', state);
      },
      onError: (error) => {
        postError('runtime', error);
      },
      onSave: (file) => postSavedFile(file, activeSaveRequestId),
      onSaveAs: (file) => postSaveAsFile(file),
      onDownload: (file) => postDownloadedFile(file),
    };

    if (source.kind === 'empty') {
      runtimeOptions.emptyType = source.emptyType;
      runtimeOptions.sourceKind = 'new-document';
    } else {
      runtimeOptions.buffer = source.buffer;
      runtimeOptions.fileName = source.fileName;
      runtimeOptions.sourceKind = source.sourceKind;
    }

    setStartupPhase('loading-runtime');
    setStartupPhase('creating-editor');
    editor = await createRuntimeOfficeEditor(root, runtimeOptions);
    stopStartupHeartbeat();
    postState('READY', editor.getState());
  } catch (error) {
    stopStartupHeartbeat();
    postError('init', error, message.requestId);
  }
}

function handlePluginWindowMessage(event: MessageEvent<OfficePluginWindowMessage>): void {
  const message = event.data;
  if (
    event.origin !== window.location.origin ||
    !message ||
    message.protocol !== OFFICE_PLUGIN_PROTOCOL ||
    !configuredPluginGuids.has(message.pluginGuid) ||
    typeof message.pluginInstanceId !== 'string' ||
    !message.pluginInstanceId ||
    !event.source
  ) {
    return;
  }

  if (message.type === 'READY') {
    const currentRuntime = pluginRuntimes.get(message.pluginGuid);
    const decision = resolveOfficePluginReady(currentRuntime, {
      pluginGuid: message.pluginGuid,
      pluginInstanceId: message.pluginInstanceId,
      source: event.source as Window,
    });
    if (decision.kind === 'ignored' || decision.kind === 'duplicate') return;
    if (decision.kind === 'replaced') {
      console.warn('[onlyoffice-browser] Office plugin runtime changed', {
        pluginGuid: message.pluginGuid,
        previousInstanceId: currentRuntime?.pluginInstanceId,
        nextInstanceId: decision.runtime.pluginInstanceId,
      });
      rejectPendingPluginRequests(
        message.pluginGuid,
        `Office plugin reloaded before operation completed: ${message.pluginGuid}`,
      );
    }
    pluginRuntimes.set(message.pluginGuid, decision.runtime);
    postPortMessage({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'PLUGIN_READY',
      sessionId,
      pluginGuid: message.pluginGuid,
      editorType: message.editorType || '',
    });
    return;
  }

  if (message.type !== 'RESULT' || !message.requestId) return;
  const pending = pendingPluginRequests.get(message.requestId);
  if (
    !pending ||
    !isOfficePluginResultForRuntime(
      pending.runtime,
      message.pluginGuid,
      message.pluginInstanceId,
      event.source as Window,
    )
  ) {
    return;
  }
  window.clearTimeout(pending.timeoutId);
  pendingPluginRequests.delete(message.requestId);
  postPortMessage({
    protocol: OFFICE_HOST_PROTOCOL,
    type: 'PLUGIN_RESULT',
    sessionId,
    requestId: message.requestId,
    pluginGuid: message.pluginGuid,
    ok: message.ok === true,
    result: message.result,
    error: message.error,
  });
}

function handleInvokePlugin(message: Extract<OfficeHostParentMessage, { type: 'INVOKE_PLUGIN' }>): void {
  const runtime = pluginRuntimes.get(message.pluginGuid);
  if (!runtime) {
    postError('plugin', new Error(`Office plugin is not ready: ${message.pluginGuid}`), message.requestId);
    return;
  }
  const timeoutId = window.setTimeout(() => {
    pendingPluginRequests.delete(message.requestId);
    postError('plugin', new Error(`Office plugin operation timed out: ${message.pluginGuid}`), message.requestId);
  }, PLUGIN_REQUEST_TIMEOUT_MS);
  pendingPluginRequests.set(message.requestId, { runtime, timeoutId });
  runtime.source.postMessage(
    {
      protocol: OFFICE_PLUGIN_PROTOCOL,
      type: 'INVOKE',
      pluginGuid: message.pluginGuid,
      pluginInstanceId: runtime.pluginInstanceId,
      requestId: message.requestId,
      payload: message.payload,
    },
    window.location.origin,
  );
}

function rejectPendingPluginRequests(pluginGuid: string | undefined, reason: string): void {
  for (const [requestId, pending] of pendingPluginRequests) {
    if (pluginGuid && pending.runtime.pluginGuid !== pluginGuid) continue;
    window.clearTimeout(pending.timeoutId);
    pendingPluginRequests.delete(requestId);
    postError('plugin', new Error(reason), requestId);
  }
}

function handleSaveAck(message: Extract<OfficeHostParentMessage, { type: 'SAVE_ACK' }>): void {
  const pending = pendingSaveAcks.get(message.requestId);
  if (!pending) return;

  pendingSaveAcks.delete(message.requestId);
  window.clearTimeout(pending.timeoutId);
  if (message.ok) {
    pending.resolve();
    return;
  }

  pending.reject(makeSaveAckError(message.message || 'Parent save acknowledgement failed'));
}

async function handleSave(message: Extract<OfficeHostParentMessage, { type: 'SAVE' }>): Promise<void> {
  try {
    if (!editor) {
      throw new Error('Editor is not open');
    }
    activeSaveRequestId = message.requestId;
    await editor.save(message.targetExt);
  } catch (error) {
    if (!(error instanceof Error && error.name === 'OnlyOfficeHostSaveAckError')) {
      postError('save', error, message.requestId);
    }
  } finally {
    if (activeSaveRequestId === message.requestId) activeSaveRequestId = undefined;
  }
}

async function handleConfirmSaveToNewFormat(
  message: Extract<OfficeHostParentMessage, { type: 'CONFIRM_SAVE_TO_NEW_FORMAT' }>,
): Promise<void> {
  try {
    if (!editor) {
      throw new Error('Editor is not open');
    }
    const confirmed = await editor.confirmSaveToNewFormat(message.options);
    postPortMessage({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'CONFIRM_SAVE_TO_NEW_FORMAT_RESULT',
      sessionId,
      requestId: message.requestId,
      confirmed,
    });
  } catch (error) {
    postError('confirm', error, message.requestId);
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

function handleSetInterfaceTheme(message: Extract<OfficeHostParentMessage, { type: 'SET_INTERFACE_THEME' }>): void {
  try {
    if (!editor) {
      throw new Error('Editor is not open');
    }
    editor.setInterfaceTheme(message.interfaceTheme);
  } catch (error) {
    postError('setInterfaceTheme', error, message.requestId);
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
    case 'SAVE_ACK':
      handleSaveAck(event.data);
      return;
    case 'CONFIRM_SAVE_TO_NEW_FORMAT':
      void handleConfirmSaveToNewFormat(event.data);
      return;
    case 'SET_READONLY':
      handleSetReadonly(event.data);
      return;
    case 'SET_INTERFACE_THEME':
      handleSetInterfaceTheme(event.data);
      return;
    case 'INVOKE_PLUGIN':
      handleInvokePlugin(event.data);
      return;
    case 'DESTROY':
      void handleDestroy();
      return;
  }
}

function handleConnect(event: MessageEvent): void {
  // Detached preview windows can have the opener deliver CONNECT while the
  // host iframe still reports HOST_READY to its popup parent. The trusted
  // boundary is the explicit parentOrigin plus the per-editor session id.
  if (!parentOrigin || !sessionId || event.origin !== parentOrigin) {
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
  setStartupPhase('connected');
}

window.addEventListener('message', handleConnect);
window.addEventListener('message', handlePluginWindowMessage);
window.addEventListener('pagehide', () => {
  void destroyRuntime();
});
window.addEventListener('unload', () => undefined);
void announceHostReady();
