import {
  isOfficeHostMessage,
  OFFICE_HOST_PROTOCOL,
  type OfficeHostChildMessage,
  type OfficeHostInitOptions,
  type OfficeHostParentMessage,
  type OfficeHostSource,
  type OfficeHostState,
  type OfficeHostWindowMessage,
} from './office-host-protocol';

const DESTROY_TIMEOUT_MS = 5_000;
const BLANK_NAVIGATION_TIMEOUT_MS = 250;
const RESET_NAVIGATION_TIMEOUT_MS = 750;
const HOST_READY_TIMEOUT_MS = 30_000;
const HOST_SELF_RESET_PATH = '/reset.html?stay=1&officeHostReset=1';
const SUPPORTED_EMPTY_TYPES = ['docx', 'xlsx', 'pptx', 'csv'] as const;
const OUTER_IFRAME_SANDBOX = 'allow-scripts allow-same-origin allow-forms allow-modals allow-downloads allow-popups';
const OUTER_IFRAME_ALLOW = 'clipboard-read; clipboard-write; fullscreen';

type OfficeEmptyType = (typeof SUPPORTED_EMPTY_TYPES)[number];
type OfficeEditorStatus = 'opening' | 'ready' | 'destroyed' | 'error';
export type OfficeEditorMode = 'edit' | 'readonly' | 'preview';
export type OfficeEditorInput = Blob | ArrayBuffer | Uint8Array;
export type OfficeHostUrlContext = {
  sessionId: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
};
export type OfficeHostUrlResolver = string | ((context: OfficeHostUrlContext) => string | URL);

export interface CreateOfficeEditorOptions {
  hostUrl: OfficeHostUrlResolver;
  file?: File | Blob;
  buffer?: OfficeEditorInput;
  url?: string;
  emptyType?: OfficeEmptyType;
  fileName?: string;
  mode?: OfficeEditorMode;
  readonly?: boolean;
  spellcheck?: boolean;
  lang?: string;
  fetchOptions?: RequestInit;
  hardResetOnLastDestroy?: boolean;
  destroyTimeoutMs?: number;
  onReady?: (instance: OfficeEditorInstance) => void;
  onSave?: (file: File, instance: OfficeEditorInstance) => void | Promise<void>;
  onError?: (error: Error, instance?: OfficeEditorInstance) => void;
}

export interface OfficeEditorState {
  id: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
  readonly: boolean;
  status: OfficeEditorStatus;
  destroyed: boolean;
}

export interface OfficeEditorInstance {
  readonly id: string;
  save(targetExt?: string): Promise<File>;
  setReadonly(readonly: boolean): void;
  destroy(): Promise<void>;
  getState(): OfficeEditorState;
}

type PendingRequest = {
  resolve: (file: File) => void;
  reject: (error: Error) => void;
};

type PreparedHostInit = {
  hostUrl: URL;
  options: OfficeHostInitOptions;
  transfer: Transferable[];
  initialState: OfficeEditorState;
};

let nextEditorId = 1;
const activeInstances = new Map<string, BrowserOfficeEditorProxy>();
const debugStats = {
  hostResetDoneCount: 0,
  hostResetTimeoutCount: 0,
};

(window as typeof window & { __officeHostDebug?: typeof debugStats }).__officeHostDebug = debugStats;

function toError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}

function normalizeExtension(value: string | undefined, fallback: string): string {
  return (value || fallback).replace(/^\./, '').toLowerCase();
}

function getFileExtension(fileName: string): string {
  return fileName.split('.').pop()?.toLowerCase() || '';
}

function getDefaultFileName(emptyType: OfficeEmptyType): string {
  return `New_Document.${emptyType}`;
}

function getSavedFileMimeType(fileName: string): string {
  const extension = getFileExtension(fileName);
  const mimeMap: Record<string, string> = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    csv: 'text/csv',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt: 'application/vnd.ms-powerpoint',
    pdf: 'application/pdf',
  };
  return mimeMap[extension] || 'application/octet-stream';
}

function readFileNameFromResponse(url: string, response: Response, fallback = 'document.docx'): string {
  const contentDisposition = response.headers.get('Content-Disposition');
  const match = contentDisposition?.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
  if (match?.[1]) {
    return match[1].replace(/['"]/g, '');
  }

  try {
    return new URL(url, window.location.href).pathname.split('/').pop() || fallback;
  } catch {
    return fallback;
  }
}

function resolveInitialMode(options: CreateOfficeEditorOptions): OfficeEditorMode {
  return options.mode || (options.readonly ? 'readonly' : 'edit');
}

function makeSessionId(): string {
  return `office-editor-${nextEditorId++}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function maybeHardResetPage(): void {
  const url = `${window.location.pathname}${window.location.search}`;
  window.setTimeout(() => {
    window.location.replace(`/reset.html?to=${encodeURIComponent(url)}`);
  }, 0);
}

function makeHostSessionLabel(sessionId: string): string {
  return sessionId.toLowerCase().replace(/[^a-z0-9-]/g, '-').replace(/-+/g, '-').slice(0, 48);
}

function isolateLocalhostHostUrl(resolved: URL, sessionId: string): void {
  if (resolved.hostname === 'localhost' || resolved.hostname.endsWith('.localhost')) {
    resolved.hostname = `host-${makeHostSessionLabel(sessionId)}.localhost`;
  }
}

function resolveHostUrl(options: CreateOfficeEditorOptions, sessionId: string, fileName: string, fileType: string): URL {
  const mode = resolveInitialMode(options);
  const hostUrl =
    typeof options.hostUrl === 'function'
      ? options.hostUrl({
          sessionId,
          fileName,
          fileType,
          mode,
        })
      : options.hostUrl;
  const resolved = new URL(hostUrl, window.location.href);
  isolateLocalhostHostUrl(resolved, sessionId);
  if (resolved.origin === window.location.origin) {
    throw new Error('createOfficeEditor requires hostUrl to be an independent origin');
  }
  resolved.searchParams.set('sessionId', sessionId);
  resolved.searchParams.set('parentOrigin', window.location.origin);
  return resolved;
}

function copyArrayBuffer(value: ArrayBuffer): ArrayBuffer {
  return value.slice(0);
}

function copyUint8Array(value: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(value.byteLength);
  copy.set(value);
  return copy.buffer;
}

async function toTransferableBuffer(input: OfficeEditorInput | Blob): Promise<ArrayBuffer> {
  if (input instanceof Blob) {
    return input.arrayBuffer();
  }
  if (input instanceof Uint8Array) {
    return copyUint8Array(input);
  }
  return copyArrayBuffer(input);
}

async function prepareHostInit(
  options: CreateOfficeEditorOptions,
  sessionId: string,
): Promise<PreparedHostInit> {
  const initialMode = resolveInitialMode(options);
  const initialReadonly = initialMode !== 'edit';

  let source: OfficeHostSource;
  let fileName = options.fileName || 'document.docx';
  let fileType = getFileExtension(fileName);
  const transfer: Transferable[] = [];

  if (options.emptyType) {
    const emptyType = normalizeExtension(options.emptyType, 'docx') as OfficeEmptyType;
    if (!SUPPORTED_EMPTY_TYPES.includes(emptyType)) {
      throw new Error(`Unsupported empty document type: ${options.emptyType}`);
    }
    fileName = options.fileName || getDefaultFileName(emptyType);
    fileType = emptyType;
    source = { kind: 'empty', emptyType };
  } else if (options.url) {
    const response = await fetch(options.url, options.fetchOptions);
    if (!response.ok) {
      throw new Error(`Failed to fetch document: ${response.status} ${response.statusText}`);
    }
    fileName = options.fileName || readFileNameFromResponse(options.url, response);
    fileType = getFileExtension(fileName);
    const blob = await response.blob();
    const buffer = await blob.arrayBuffer();
    transfer.push(buffer);
    source = {
      kind: 'buffer',
      buffer,
      fileName,
      mimeType: blob.type || getSavedFileMimeType(fileName),
    };
  } else {
    const input = options.file || options.buffer;
    if (!input) {
      throw new Error('createOfficeEditor requires file, buffer, url, or emptyType');
    }
    fileName = options.fileName || (input instanceof File ? input.name : 'document.docx');
    fileType = getFileExtension(fileName);
    const buffer = await toTransferableBuffer(input);
    transfer.push(buffer);
    source = {
      kind: 'buffer',
      buffer,
      fileName,
      mimeType: input instanceof Blob ? input.type || getSavedFileMimeType(fileName) : getSavedFileMimeType(fileName),
    };
  }

  const hostUrl = resolveHostUrl(options, sessionId, fileName, fileType);

  return {
    hostUrl,
    options: {
      fileName,
      mode: options.mode,
      readonly: options.readonly,
      spellcheck: options.spellcheck ?? false,
      lang: options.lang,
      source,
    },
    transfer,
    initialState: {
      id: sessionId,
      fileName,
      fileType,
      mode: initialMode,
      readonly: initialReadonly,
      status: 'opening',
      destroyed: false,
    },
  };
}

function toPublicState(state: OfficeHostState | OfficeEditorState): OfficeEditorState {
  return {
    id: state.id,
    fileName: state.fileName,
    fileType: state.fileType,
    mode: state.mode,
    readonly: state.readonly,
    status: state.status,
    destroyed: state.destroyed,
  };
}

function nextRequestId(sessionId: string): string {
  return `${sessionId}-request-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function applyFillContainerDefaults(element: HTMLElement): void {
  const { style } = element;
  style.width ||= '100%';
  style.height ||= '100%';
  style.minWidth ||= '0';
  style.minHeight ||= '0';
}

function applyHostFrameDefaults(iframe: HTMLIFrameElement): void {
  applyFillContainerDefaults(iframe);
  iframe.style.display ||= 'block';
  iframe.style.border ||= '0';
}

class BrowserOfficeEditorProxy implements OfficeEditorInstance {
  readonly id: string;
  private readonly container: HTMLElement;
  private readonly options: CreateOfficeEditorOptions;
  private readonly hostOrigin: string;
  private readonly prepared: PreparedHostInit;
  private readonly pendingRequests = new Map<string, PendingRequest>();
  private iframe: HTMLIFrameElement | null = null;
  private port: MessagePort | null = null;
  private hostWindow: Window | null = null;
  private windowMessageListener: ((event: MessageEvent) => void) | null = null;
  private hostReadyTimeout: number | null = null;
  private readyResolve: ((instance: BrowserOfficeEditorProxy) => void) | null = null;
  private readyReject: ((error: Error) => void) | null = null;
  private destroyAckResolve: (() => void) | null = null;
  private connected = false;
  private mounted = false;
  private readyNotified = false;
  private destroyed = false;
  private destroyPromise: Promise<void> | null = null;
  private state: OfficeEditorState;

  private constructor(container: HTMLElement, options: CreateOfficeEditorOptions, prepared: PreparedHostInit) {
    this.id = prepared.initialState.id;
    this.container = container;
    this.options = options;
    this.prepared = prepared;
    this.hostOrigin = prepared.hostUrl.origin;
    this.state = prepared.initialState;
  }

  static async create(container: HTMLElement, options: CreateOfficeEditorOptions): Promise<BrowserOfficeEditorProxy> {
    const sessionId = makeSessionId();
    const prepared = await prepareHostInit(options, sessionId);
    const instance = new BrowserOfficeEditorProxy(container, options, prepared);
    await instance.mount();
    return instance;
  }

  private mount(): Promise<BrowserOfficeEditorProxy> {
    this.container.replaceChildren();
    this.container.classList.add('office-editor-host');
    applyFillContainerDefaults(this.container);

    const iframe = document.createElement('iframe');
    iframe.className = 'office-editor-host-frame';
    iframe.title = this.prepared.options.fileName || 'Office editor';
    iframe.setAttribute('sandbox', OUTER_IFRAME_SANDBOX);
    iframe.setAttribute('allow', OUTER_IFRAME_ALLOW);
    iframe.setAttribute('referrerpolicy', 'no-referrer');
    applyHostFrameDefaults(iframe);

    this.iframe = iframe;
    this.hostWindow = iframe.contentWindow;

    const readyPromise = new Promise<BrowserOfficeEditorProxy>((resolve, reject) => {
      this.readyResolve = resolve;
      this.readyReject = reject;
    });

    this.windowMessageListener = (event) => this.handleWindowMessage(event);
    window.addEventListener('message', this.windowMessageListener);

    this.hostReadyTimeout = window.setTimeout(() => {
      this.failBeforeReady(new Error('Timed out waiting for the office host to become ready'));
    }, HOST_READY_TIMEOUT_MS);

    this.container.appendChild(iframe);
    this.hostWindow = iframe.contentWindow;
    iframe.src = this.prepared.hostUrl.href;
    activeInstances.set(this.id, this);

    return readyPromise.catch((error) => {
      this.destroy();
      throw error;
    });
  }

  private handleWindowMessage(event: MessageEvent): void {
    if (this.destroyed || this.connected || event.origin !== this.hostOrigin || event.source !== this.iframe?.contentWindow) {
      return;
    }
    if (!isOfficeHostMessage(event.data, this.id)) {
      return;
    }

    const message = event.data as OfficeHostWindowMessage;
    if (message.type !== 'HOST_READY') {
      return;
    }

    this.connected = true;
    this.clearHostReadyTimeout();
    this.removeWindowMessageListener();

    const channel = new MessageChannel();
    this.port = channel.port1;
    this.port.onmessage = (portEvent) => this.handlePortMessage(portEvent);
    this.port.start();

    this.iframe?.contentWindow?.postMessage(
      {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'CONNECT',
        sessionId: this.id,
      } satisfies OfficeHostWindowMessage,
      this.hostOrigin,
      [channel.port2],
    );

    this.postToHost(
      {
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'INIT',
        sessionId: this.id,
        options: this.prepared.options,
      },
      this.prepared.transfer,
    );
  }

  private handlePortMessage(event: MessageEvent<OfficeHostChildMessage>): void {
    if (this.destroyed && event.data?.type !== 'DESTROYED') return;
    if (!isOfficeHostMessage(event.data, this.id)) return;

    const message = event.data;
    switch (message.type) {
      case 'READY':
        this.state = toPublicState(message.state);
        this.mounted = true;
        this.readyResolve?.(this);
        this.readyResolve = null;
        this.readyReject = null;
        this.maybeNotifyReady();
        return;
      case 'STATE':
        this.state = toPublicState(message.state);
        this.maybeNotifyReady();
        return;
      case 'SAVE_RESULT':
        this.handleSaveResult(message);
        return;
      case 'ERROR':
        this.handleHostError(message);
        return;
      case 'DESTROYED':
        this.destroyAckResolve?.();
        this.destroyAckResolve = null;
        return;
    }
  }

  private handleSaveResult(message: Extract<OfficeHostChildMessage, { type: 'SAVE_RESULT' }>): void {
    const file = new File([message.buffer], message.fileName, { type: message.mimeType || getSavedFileMimeType(message.fileName) });
    if (message.requestId) {
      const request = this.pendingRequests.get(message.requestId);
      if (!request) return;
      this.pendingRequests.delete(message.requestId);
      request.resolve(file);
    }
    void Promise.resolve(this.options.onSave?.(file, this)).catch((error) => {
      this.options.onError?.(toError(error), this);
    });
  }

  private handleHostError(message: Extract<OfficeHostChildMessage, { type: 'ERROR' }>): void {
    const error = new Error(message.message);
    if (message.requestId) {
      const request = this.pendingRequests.get(message.requestId);
      if (request) {
        this.pendingRequests.delete(message.requestId);
        request.reject(error);
        this.options.onError?.(error, this);
        return;
      }
    }

    if (!this.mounted) {
      this.failBeforeReady(error);
      return;
    }

    this.state = { ...this.state, status: 'error' };
    this.options.onError?.(error, this);
  }

  private failBeforeReady(error: Error): void {
    this.state = { ...this.state, status: 'error' };
    this.clearHostReadyTimeout();
    this.readyReject?.(error);
    this.readyResolve = null;
    this.readyReject = null;
    this.options.onError?.(error, this);
  }

  private maybeNotifyReady(): void {
    if (this.readyNotified || this.state.status !== 'ready') return;
    this.readyNotified = true;
    this.options.onReady?.(this);
  }

  private postToHost(message: OfficeHostParentMessage, transfer: Transferable[] = []): void {
    this.port?.postMessage(message, transfer);
  }

  private clearHostReadyTimeout(): void {
    if (this.hostReadyTimeout !== null) {
      window.clearTimeout(this.hostReadyTimeout);
      this.hostReadyTimeout = null;
    }
  }

  private removeWindowMessageListener(): void {
    if (this.windowMessageListener) {
      window.removeEventListener('message', this.windowMessageListener);
      this.windowMessageListener = null;
    }
  }

  save(targetExt?: string): Promise<File> {
    if (this.destroyed || !this.port) {
      return Promise.reject(new Error('Editor is not open'));
    }
    if (this.state.readonly) {
      return Promise.reject(new Error('Current document is readonly'));
    }

    const requestId = nextRequestId(this.id);
    return new Promise<File>((resolve, reject) => {
      this.pendingRequests.set(requestId, { resolve, reject });
      this.postToHost({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'SAVE',
        sessionId: this.id,
        requestId,
        targetExt,
      });
    });
  }

  setReadonly(readonly: boolean): void {
    if (this.state.mode === 'preview' && !readonly) {
      throw new Error('Preview mode cannot be switched to editing; recreate the editor with mode: "edit".');
    }
    this.state = {
      ...this.state,
      readonly,
      mode: this.state.mode === 'preview' ? 'preview' : readonly ? 'readonly' : 'edit',
    };
    if (!this.destroyed && this.port) {
      this.postToHost({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'SET_READONLY',
        sessionId: this.id,
        readonly,
      });
    }
  }

  getState(): OfficeEditorState {
    return { ...this.state };
  }

  destroy(): Promise<void> {
    if (this.destroyPromise) return this.destroyPromise;
    this.destroyed = true;
    this.mounted = false;
    this.state = {
      ...this.state,
      status: 'destroyed',
      destroyed: true,
    };

    this.clearHostReadyTimeout();
    this.removeWindowMessageListener();
    this.readyReject?.(new Error('Editor was destroyed before it became ready'));
    this.readyResolve = null;
    this.readyReject = null;

    for (const request of this.pendingRequests.values()) {
      request.reject(new Error('Editor was destroyed before save completed'));
    }
    this.pendingRequests.clear();
    activeInstances.delete(this.id);

    this.destroyPromise = this.destroyHost();
    return this.destroyPromise;
  }

  private async destroyHost(): Promise<void> {
    const destroyTimeoutMs = this.options.destroyTimeoutMs ?? DESTROY_TIMEOUT_MS;
    let hostResetDone = false;
    let removeResetListener = () => {};

    if (this.port) {
      const portAck = new Promise<void>((resolve) => {
        this.destroyAckResolve = () => {
          resolve();
        };
      });
      const resetAck = new Promise<void>((resolve) => {
        const listener = (event: MessageEvent) => {
          if (event.origin !== this.hostOrigin || event.source !== this.iframe?.contentWindow) return;
          if (!isOfficeHostMessage(event.data, this.id)) return;
          const message = event.data as OfficeHostWindowMessage;
          if (message.type !== 'HOST_RESET_DONE') return;
          hostResetDone = true;
          debugStats.hostResetDoneCount += 1;
          resolve();
        };
        window.addEventListener('message', listener);
        removeResetListener = () => window.removeEventListener('message', listener);
      });
      this.postToHost({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'DESTROY',
        sessionId: this.id,
      });
      await Promise.race([resetAck, portAck, delay(destroyTimeoutMs)]);
    }
    if (!hostResetDone) {
      debugStats.hostResetTimeoutCount += 1;
    }
    removeResetListener();

    await this.forceRemoveIframe(hostResetDone);
    if (this.port) {
      this.port.onmessage = null;
      this.port.close();
    }
    this.port = null;
    this.hostWindow = null;
    this.destroyAckResolve = null;
    this.container.replaceChildren();

    if (this.options.hardResetOnLastDestroy && activeInstances.size === 0) {
      maybeHardResetPage();
    }
  }

  private async forceRemoveIframe(hostResetDone = false): Promise<void> {
    const iframe = this.iframe;
    this.iframe = null;
    if (!iframe) return;

    if (iframe.isConnected) {
      if (!hostResetDone) {
        const resetUrl = this.getHostResetUrl();
        await this.navigateIframeForTeardown(iframe, resetUrl, RESET_NAVIGATION_TIMEOUT_MS);
      }
      await this.navigateIframeForTeardown(iframe, 'about:blank', BLANK_NAVIGATION_TIMEOUT_MS);
      iframe.remove();
    }
  }

  private getHostResetUrl(): string {
    const resetUrl = new URL(HOST_SELF_RESET_PATH, this.prepared.hostUrl.href);
    resetUrl.searchParams.set('sessionId', this.id);
    resetUrl.searchParams.set('parentOrigin', window.location.origin);
    return resetUrl.href;
  }

  private async navigateIframeForTeardown(iframe: HTMLIFrameElement, url: string, timeoutMs: number): Promise<void> {
    const loaded = new Promise<void>((resolve) => {
      iframe.addEventListener('load', () => resolve(), { once: true });
    });
    iframe.src = url;
    await Promise.race([loaded, delay(timeoutMs)]);
  }
}

export async function createOfficeEditor(
  container: HTMLElement,
  options: CreateOfficeEditorOptions,
): Promise<OfficeEditorInstance> {
  if (!(container instanceof HTMLElement)) {
    throw new Error('createOfficeEditor requires an HTMLElement container');
  }

  try {
    return await BrowserOfficeEditorProxy.create(container, options);
  } catch (error) {
    const normalized = toError(error);
    options.onError?.(normalized);
    throw normalized;
  }
}

export function loadOfficeEditorApi(): Promise<void> {
  return Promise.resolve();
}

export function getActiveOfficeEditorCount(): number {
  return activeInstances.size;
}
