import {
  isOfficeHostMessage,
  OFFICE_HOST_PROTOCOL,
  type OfficeHostChildMessage,
  type OfficeHostInterfaceTheme,
  type OfficeHostIdentity,
  type OfficeHostInitOptions,
  type OfficeHostPluginOptions,
  type OfficeHostParentMessage,
  type OfficeHostSaveBehavior,
  type OfficeHostSource,
  type OfficeHostSourceKind,
  type OfficeSaveToNewFormatConfirmationOptions,
  type OfficeHostState,
  type OfficeHostWindowMessage,
} from './office-host-protocol';
import { officeHostIdentitiesEqual } from './office-host-identity';
import { readOfficeHostBootstrap, writeOfficeHostBootstrap } from './office-host-url';
import {
  OFFICE_EDITOR_ORIGIN_SLOTS,
  isReusableOfficeEditorHostname,
  type OfficeEditorOriginSlot,
} from './office-origin-pool';

const DESTROY_TIMEOUT_MS = 5_000;
const BLANK_NAVIGATION_TIMEOUT_MS = 250;
const RESET_NAVIGATION_TIMEOUT_MS = 750;
const HOST_READY_TIMEOUT_MS = 30_000;
const STARTUP_HEARTBEAT_TIMEOUT_MS = 30_000;
const STARTUP_TOTAL_TIMEOUT_MS = 5 * 60_000;
const PLUGIN_REQUEST_TIMEOUT_MS = 45_000;
const HOST_SELF_RESET_PATH = '/reset.html?stay=1&officeHostReset=1';
const SUPPORTED_EMPTY_TYPES = ['docx', 'xlsx', 'pptx', 'csv'] as const;
const OUTER_IFRAME_ALLOW = 'clipboard-read; clipboard-write; fullscreen';
const OFFICE_EDITOR_PROXY_RUNTIME_REVISION = 'font-packages-v1';
const PRINT_TITLE_RESTORE_MS = 45_000;

function isHTMLElementContainer(value: unknown): value is HTMLElement {
  if (value instanceof HTMLElement) return true;
  if (!value || typeof value !== 'object') return false;
  const element = value as Element;
  const ownerHTMLElement = element.ownerDocument?.defaultView?.HTMLElement;
  return typeof ownerHTMLElement === 'function' && value instanceof ownerHTMLElement;
}

function hideIframeForTeardown(iframe: HTMLIFrameElement): void {
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.visibility = 'hidden';
  iframe.style.opacity = '0';
  iframe.style.pointerEvents = 'none';
  iframe.style.background = 'transparent';
}

type OfficeEmptyType = (typeof SUPPORTED_EMPTY_TYPES)[number];
type OfficeEditorStatus = 'opening' | 'ready' | 'destroyed' | 'error';
export type OfficeEditorMode = 'edit' | 'readonly' | 'preview';
export type OfficeEditorInput = Blob | ArrayBuffer | Uint8Array;
export type OfficeEditorSourceKind = OfficeHostSourceKind;
export type OfficeSaveBehavior = OfficeHostSaveBehavior;
export type OfficeInterfaceTheme = OfficeHostInterfaceTheme;
export type OfficePluginOptions = OfficeHostPluginOptions;
export type { OfficeHostIdentity, OfficeSaveToNewFormatConfirmationOptions };
export type OfficeSaveCallbackResult = void | boolean;
export type OfficeSaveAsCallbackResult = void | boolean;
export type OfficeDownloadCallbackResult = void;
export type OfficeHostUrlContext = {
  sessionId: string;
  hostSlot: OfficeEditorOriginSlot;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
};
export type OfficeHostUrlResolver = string | ((context: OfficeHostUrlContext) => string | URL);

export interface CreateOfficeEditorOptions {
  hostUrl: OfficeHostUrlResolver;
  expectedHostIdentity?: OfficeHostIdentity;
  file?: File | Blob;
  buffer?: OfficeEditorInput;
  url?: string;
  emptyType?: OfficeEmptyType;
  fileName?: string;
  mode?: OfficeEditorMode;
  readonly?: boolean;
  /** Preserve the viewer origin when restoring a document directly in edit mode. */
  canReturnToPreview?: boolean;
  spellcheck?: boolean;
  interfaceTheme?: OfficeInterfaceTheme;
  lang?: string;
  plugins?: OfficePluginOptions;
  /** Font asset paths explicitly downloaded by the embedding application. */
  downloadedFonts?: string[];
  fetchOptions?: RequestInit;
  hardResetOnLastDestroy?: boolean;
  destroyTimeoutMs?: number;
  onReady?: (instance: OfficeEditorInstance) => void;
  onPluginReady?: (pluginGuid: string, editorType: string, instance: OfficeEditorInstance) => void;
  saveBehavior?: OfficeSaveBehavior;
  onSave?: (file: File, instance: OfficeEditorInstance) => OfficeSaveCallbackResult | Promise<OfficeSaveCallbackResult>;
  onSaveAs?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeSaveAsCallbackResult | Promise<OfficeSaveAsCallbackResult>;
  onDownload?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeDownloadCallbackResult | Promise<OfficeDownloadCallbackResult>;
  onDirtyChange?: (dirty: boolean, instance: OfficeEditorInstance) => void | Promise<void>;
  onStateChange?: (state: OfficeEditorState, instance: OfficeEditorInstance) => void | Promise<void>;
  onError?: (error: Error, instance?: OfficeEditorInstance) => void;
}

export interface OfficeEditorState {
  id: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
  readonly: boolean;
  dirty: boolean;
  sourceKind: OfficeEditorSourceKind;
  status: OfficeEditorStatus;
  destroyed: boolean;
}

export interface OfficeEditorInstance {
  readonly id: string;
  invokePlugin(pluginGuid: string, payload: unknown): Promise<unknown>;
  save(targetExt?: string): Promise<File>;
  confirmSaveToNewFormat(options?: OfficeSaveToNewFormatConfirmationOptions): Promise<boolean>;
  setInterfaceTheme(theme: OfficeInterfaceTheme): void;
  setReadonly(readonly: boolean): void;
  destroy(): Promise<void>;
  getState(): OfficeEditorState;
  getHostIdentity(): OfficeHostIdentity;
}

export type OfficeEditorMountPhase =
  | 'host-loading'
  | 'waiting-for-activation'
  | 'runtime-loading'
  | 'ready'
  | 'error'
  | 'destroyed';

export interface OfficeEditorMountState {
  id: string;
  origin: string;
  phase: OfficeEditorMountPhase;
  error?: Error;
}

export interface OfficeEditorMount {
  readonly id: string;
  activate(): Promise<OfficeEditorInstance>;
  destroy(): Promise<void>;
  getState(): OfficeEditorMountState;
}

export class OfficeHostIdentityMismatchError extends Error {
  readonly expected: OfficeHostIdentity;
  readonly actual: OfficeHostIdentity;

  constructor(expected: OfficeHostIdentity, actual: OfficeHostIdentity) {
    super(
      `Office host identity mismatch (expected ${expected.packageVersion}/${expected.hostBuildId}/${expected.assetManifestDigest}, received ${actual.packageVersion}/${actual.hostBuildId}/${actual.assetManifestDigest})`,
    );
    this.name = 'OfficeHostIdentityMismatchError';
    this.expected = expected;
    this.actual = actual;
  }
}

export class OfficeHostIsolationError extends Error {
  readonly origin: string;
  readonly existingSessionId: string;
  readonly requestedSessionId: string;

  constructor(origin: string, existingSessionId: string, requestedSessionId: string) {
    super(`Office host origin ${origin} is already owned by active session ${existingSessionId}`);
    this.name = 'OfficeHostIsolationError';
    this.origin = origin;
    this.existingSessionId = existingSessionId;
    this.requestedSessionId = requestedSessionId;
  }
}

export class OfficeHostPoolExhaustedError extends Error {
  readonly capacity = OFFICE_EDITOR_ORIGIN_SLOTS.length;

  constructor() {
    super(`All ${OFFICE_EDITOR_ORIGIN_SLOTS.length} Office editor origins are in use`);
    this.name = 'OfficeHostPoolExhaustedError';
  }
}

export class OfficeEditorStartupError extends Error {
  readonly phase: string;
  readonly retryable: boolean;

  constructor(phase: string, message: string, retryable = true) {
    super(message);
    this.name = 'OfficeEditorStartupError';
    this.phase = phase;
    this.retryable = retryable;
  }
}

type PendingRequest = {
  resolve: (file: File) => void;
  reject: (error: Error) => void;
};

type PendingConfirmationRequest = {
  resolve: (confirmed: boolean) => void;
  reject: (error: Error) => void;
};

type PendingPluginRequest = {
  resolve: (result: unknown) => void;
  reject: (error: Error) => void;
  timeoutId: number;
};

type PreparedHostInit = {
  options: OfficeHostInitOptions;
  transfer: Transferable[];
  initialState: OfficeEditorState;
};

type InitialHostDescriptor = {
  hostUrl: URL;
  initialState: OfficeEditorState;
};

let nextEditorId = 1;
const activeInstances = new Map<string, BrowserOfficeEditorProxy>();
const activeOrigins = new Map<string, BrowserOfficeEditorProxy>();
const debugStats = {
  hostResetDoneCount: 0,
  hostResetTimeoutCount: 0,
  startupHeartbeatPortCount: 0,
  startupHeartbeatCount: 0,
  activeHostPortCount: 0,
  peakActiveHostPortCount: 0,
  activeStartupHeartbeatPortCount: 0,
  peakActiveStartupHeartbeatPortCount: 0,
  get activeInstanceCount() {
    return activeInstances.size;
  },
  get activeOriginLeaseCount() {
    return activeOrigins.size;
  },
};
let temporaryDocumentTitleOriginal: string | null = null;
let temporaryDocumentTitleTimeout: number | null = null;

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

function downloadFile(file: File): void {
  const objectUrl = URL.createObjectURL(file);
  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = file.name || 'document';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function restoreTemporaryDocumentTitle(): void {
  if (temporaryDocumentTitleTimeout !== null) {
    window.clearTimeout(temporaryDocumentTitleTimeout);
    temporaryDocumentTitleTimeout = null;
  }
  if (temporaryDocumentTitleOriginal !== null) {
    document.title = temporaryDocumentTitleOriginal;
    temporaryDocumentTitleOriginal = null;
  }
  window.removeEventListener('afterprint', restoreTemporaryDocumentTitle);
}

function setTemporaryDocumentTitle(title: string, durationMs = PRINT_TITLE_RESTORE_MS): void {
  const normalizedTitle = title.trim();
  if (!normalizedTitle) return;

  if (temporaryDocumentTitleOriginal === null) {
    temporaryDocumentTitleOriginal = document.title;
  }
  document.title = normalizedTitle;

  if (temporaryDocumentTitleTimeout !== null) {
    window.clearTimeout(temporaryDocumentTitleTimeout);
  }
  window.removeEventListener('afterprint', restoreTemporaryDocumentTitle);
  window.addEventListener('afterprint', restoreTemporaryDocumentTitle, { once: true });
  temporaryDocumentTitleTimeout = window.setTimeout(restoreTemporaryDocumentTitle, Math.max(1_000, durationMs));
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
  const uuid = globalThis.crypto?.randomUUID?.();
  return uuid
    ? `office-editor-${uuid}`
    : `office-editor-${nextEditorId++}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function maybeHardResetPage(): void {
  const url = `${window.location.pathname}${window.location.search}`;
  window.setTimeout(() => {
    window.location.replace(`/reset.html?to=${encodeURIComponent(url)}`);
  }, 0);
}

function isolateLocalhostHostUrl(resolved: URL, hostSlot: OfficeEditorOriginSlot): void {
  if (
    (resolved.hostname === 'localhost' || resolved.hostname.endsWith('.localhost')) &&
    !isReusableOfficeEditorHostname(resolved.hostname)
  ) {
    resolved.hostname = `host-${hostSlot}.office.localhost`;
  }
}

function resolveHostUrl(
  options: CreateOfficeEditorOptions,
  sessionId: string,
  hostSlot: OfficeEditorOriginSlot,
  fileName: string,
  fileType: string,
): URL {
  const mode = resolveInitialMode(options);
  const hostUrl =
    typeof options.hostUrl === 'function'
      ? options.hostUrl({
          sessionId,
          hostSlot,
          fileName,
          fileType,
          mode,
        })
      : options.hostUrl;
  const resolved = new URL(hostUrl, window.location.href);
  isolateLocalhostHostUrl(resolved, hostSlot);
  if (resolved.origin === window.location.origin) {
    throw new Error('createOfficeEditor requires hostUrl to be an independent origin');
  }
  return writeOfficeHostBootstrap(resolved, {
    sessionId,
    parentOrigin: window.location.origin,
  });
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

function describeHostInit(
  options: CreateOfficeEditorOptions,
  sessionId: string,
  hostSlot: OfficeEditorOriginSlot,
): InitialHostDescriptor {
  const initialMode = resolveInitialMode(options);
  const initialReadonly = initialMode !== 'edit';
  const emptyType = options.emptyType ? (normalizeExtension(options.emptyType, 'docx') as OfficeEmptyType) : undefined;
  if (emptyType && !SUPPORTED_EMPTY_TYPES.includes(emptyType)) {
    throw new Error(`Unsupported empty document type: ${options.emptyType}`);
  }
  const input = options.file || options.buffer;
  let urlFileName: string | undefined;
  if (options.url) {
    try {
      urlFileName = new URL(options.url, window.location.href).pathname.split('/').pop() || 'document.docx';
    } catch {
      urlFileName = 'document.docx';
    }
  }
  const fileName =
    options.fileName ||
    (emptyType ? getDefaultFileName(emptyType) : undefined) ||
    (input instanceof File ? input.name : undefined) ||
    urlFileName ||
    'document.docx';
  const fileType = emptyType || getFileExtension(fileName);
  const sourceKind: OfficeEditorSourceKind = emptyType
    ? 'new-document'
    : options.url
      ? 'url'
      : options.file
        ? 'local-file'
        : 'buffer';

  if (!emptyType && !options.url && !input) {
    throw new Error('createOfficeEditor requires file, buffer, url, or emptyType');
  }

  return {
    hostUrl: resolveHostUrl(options, sessionId, hostSlot, fileName, fileType),
    initialState: {
      id: sessionId,
      fileName,
      fileType,
      mode: initialMode,
      readonly: initialReadonly,
      dirty: false,
      sourceKind,
      status: 'opening',
      destroyed: false,
    },
  };
}

async function prepareHostInit(
  options: CreateOfficeEditorOptions,
  descriptor: InitialHostDescriptor,
): Promise<PreparedHostInit> {
  const sessionId = descriptor.initialState.id;
  const initialMode = resolveInitialMode(options);
  const initialReadonly = initialMode !== 'edit';

  let source: OfficeHostSource;
  let fileName = descriptor.initialState.fileName;
  let fileType = descriptor.initialState.fileType;
  let sourceKind: OfficeEditorSourceKind = descriptor.initialState.sourceKind;
  const transfer: Transferable[] = [];

  if (options.emptyType) {
    const emptyType = normalizeExtension(options.emptyType, 'docx') as OfficeEmptyType;
    if (!SUPPORTED_EMPTY_TYPES.includes(emptyType)) {
      throw new Error(`Unsupported empty document type: ${options.emptyType}`);
    }
    fileName = options.fileName || getDefaultFileName(emptyType);
    fileType = emptyType;
    sourceKind = 'new-document';
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
    sourceKind = 'url';
    source = {
      kind: 'buffer',
      buffer,
      fileName,
      mimeType: blob.type || getSavedFileMimeType(fileName),
      sourceKind,
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
    sourceKind = options.file ? 'local-file' : 'buffer';
    source = {
      kind: 'buffer',
      buffer,
      fileName,
      mimeType: input instanceof Blob ? input.type || getSavedFileMimeType(fileName) : getSavedFileMimeType(fileName),
      sourceKind,
    };
  }

  return {
    options: {
      fileName,
      mode: options.mode,
      readonly: options.readonly,
      canReturnToPreview: options.canReturnToPreview,
      spellcheck: options.spellcheck ?? false,
      interfaceTheme: options.interfaceTheme,
      lang: options.lang,
      plugins: options.plugins,
      downloadedFonts: options.downloadedFonts,
      saveBehavior: options.saveBehavior,
      source,
    },
    transfer,
    initialState: {
      id: sessionId,
      fileName,
      fileType,
      mode: initialMode,
      readonly: initialReadonly,
      dirty: false,
      sourceKind,
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
    dirty: state.dirty,
    sourceKind: state.sourceKind || 'buffer',
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

function enforceUnsandboxedHostFrame(iframe: HTMLIFrameElement): () => void {
  const removeSandbox = () => {
    if (!iframe.hasAttribute('sandbox')) return;
    iframe.removeAttribute('sandbox');
    console.warn(
      'OnlyOffice browser removed sandbox from the editor host iframe. The isolated host origin is the security boundary; sandbox breaks native PDF printing in Chrome.',
    );
  };

  removeSandbox();
  if (typeof MutationObserver === 'undefined') return () => {};

  const observer = new MutationObserver((mutations) => {
    if (mutations.some((mutation) => mutation.attributeName === 'sandbox')) {
      removeSandbox();
    }
  });
  observer.observe(iframe, { attributes: true, attributeFilter: ['sandbox'] });
  return () => observer.disconnect();
}

class BrowserOfficeEditorProxy implements OfficeEditorInstance {
  readonly id: string;
  private readonly container: HTMLElement;
  private readonly options: CreateOfficeEditorOptions;
  private readonly hostOrigin: string;
  private readonly descriptor: InitialHostDescriptor;
  private prepared: PreparedHostInit | null = null;
  private readonly pendingRequests = new Map<string, PendingRequest>();
  private readonly pendingConfirmationRequests = new Map<string, PendingConfirmationRequest>();
  private readonly pendingPluginRequests = new Map<string, PendingPluginRequest>();
  private readonly parentWindow: Window;
  private readonly readyPromise: Promise<BrowserOfficeEditorProxy>;
  private iframe: HTMLIFrameElement | null = null;
  private port: MessagePort | null = null;
  private startupHeartbeatPort: MessagePort | null = null;
  private hostWindow: Window | null = null;
  private windowMessageListener: ((event: MessageEvent) => void) | null = null;
  private hostReadyTimeout: number | null = null;
  private startupHeartbeatTimeout: number | null = null;
  private startupTotalTimeout: number | null = null;
  private readyResolve: ((instance: BrowserOfficeEditorProxy) => void) | null = null;
  private readyReject: ((error: Error) => void) | null = null;
  private destroyAckResolve: (() => void) | null = null;
  private originLeaseState: 'active' | 'retiring' | 'released' = 'active';
  private connected = false;
  private activationStarted = false;
  private initSent = false;
  private mounted = false;
  private readyNotified = false;
  private destroyed = false;
  private destroyPromise: Promise<void> | null = null;
  private unsandboxedHostFrameCleanup: (() => void) | null = null;
  private state: OfficeEditorState;
  private mountPhase: OfficeEditorMountPhase = 'host-loading';
  private mountError: Error | undefined;
  private activationPromise: Promise<OfficeEditorInstance> | null = null;
  private hostIdentity: OfficeHostIdentity | null = null;
  private readonly returnsToPreview: boolean;

  private constructor(container: HTMLElement, options: CreateOfficeEditorOptions, descriptor: InitialHostDescriptor) {
    this.id = descriptor.initialState.id;
    this.container = container;
    this.options = options;
    this.descriptor = descriptor;
    this.hostOrigin = descriptor.hostUrl.origin;
    this.state = descriptor.initialState;
    this.returnsToPreview = descriptor.initialState.mode === 'preview' || options.canReturnToPreview === true;
    this.parentWindow = container.ownerDocument.defaultView || window;
    this.readyPromise = new Promise<BrowserOfficeEditorProxy>((resolve, reject) => {
      this.readyResolve = resolve;
      this.readyReject = reject;
    });
    void this.readyPromise.catch(() => undefined);
  }

  static mount(container: HTMLElement, options: CreateOfficeEditorOptions): BrowserOfficeEditorProxy {
    const sessionId = makeSessionId();
    let descriptor: InitialHostDescriptor | null = null;
    let firstConflict: { origin: string; sessionId: string } | null = null;
    const resolvedOrigins = new Set<string>();
    for (const hostSlot of OFFICE_EDITOR_ORIGIN_SLOTS) {
      const candidate = describeHostInit(options, sessionId, hostSlot);
      if (resolvedOrigins.has(candidate.hostUrl.origin)) continue;
      resolvedOrigins.add(candidate.hostUrl.origin);
      const existing = activeOrigins.get(candidate.hostUrl.origin);
      if (!existing?.holdsOriginLease()) {
        if (existing) activeOrigins.delete(candidate.hostUrl.origin);
        descriptor = candidate;
        break;
      }
      firstConflict ||= { origin: candidate.hostUrl.origin, sessionId: existing.id };
    }
    if (!descriptor && firstConflict && resolvedOrigins.size === 1) {
      throw new OfficeHostIsolationError(firstConflict.origin, firstConflict.sessionId, sessionId);
    }
    if (!descriptor) throw new OfficeHostPoolExhaustedError();

    const instance = new BrowserOfficeEditorProxy(container, options, descriptor);
    activeOrigins.set(descriptor.hostUrl.origin, instance);
    instance.mountHost();
    return instance;
  }

  private holdsOriginLease(): boolean {
    if (this.originLeaseState === 'retiring') return true;
    if (this.originLeaseState === 'released') return false;
    return Boolean(this.iframe?.isConnected);
  }

  private releaseOriginLease(): void {
    if (this.originLeaseState === 'released') return;
    this.originLeaseState = 'released';
    if (activeOrigins.get(this.hostOrigin) === this) {
      activeOrigins.delete(this.hostOrigin);
    }
  }

  private mountHost(): void {
    this.container.replaceChildren();
    this.container.classList.add('office-editor-host');
    applyFillContainerDefaults(this.container);

    const iframe = this.container.ownerDocument.createElement('iframe');
    iframe.className = 'office-editor-host-frame';
    iframe.dataset.officeEditorProxyRevision = OFFICE_EDITOR_PROXY_RUNTIME_REVISION;
    iframe.title = this.descriptor.initialState.fileName || 'Office editor';
    // Do not sandbox: native PDF printing needs same-origin script access inside
    // the independent editor host's nested print iframe.
    iframe.removeAttribute('sandbox');
    iframe.setAttribute('allow', OUTER_IFRAME_ALLOW);
    iframe.setAttribute('referrerpolicy', 'no-referrer');
    applyHostFrameDefaults(iframe);
    this.unsandboxedHostFrameCleanup = enforceUnsandboxedHostFrame(iframe);

    this.iframe = iframe;
    this.hostWindow = iframe.contentWindow;

    this.windowMessageListener = (event) => this.handleWindowMessage(event);
    this.parentWindow.addEventListener('message', this.windowMessageListener);

    this.hostReadyTimeout = this.parentWindow.setTimeout(() => {
      this.failBeforeReady(
        new OfficeEditorStartupError('host-connect', 'Timed out waiting for the office host to become ready'),
      );
    }, HOST_READY_TIMEOUT_MS);

    this.container.appendChild(iframe);
    this.hostWindow = iframe.contentWindow;
    iframe.src = this.descriptor.hostUrl.href;
    activeInstances.set(this.id, this);
  }

  activate(): Promise<OfficeEditorInstance> {
    if (this.activationPromise) return this.activationPromise;
    if (this.mountError) return Promise.reject(this.mountError);
    if (this.destroyed) return Promise.reject(new Error('Editor was destroyed before activation'));

    this.activationStarted = true;
    if (this.connected) this.mountPhase = 'runtime-loading';
    this.startStartupWatchdogs();
    this.activationPromise = (async () => {
      try {
        const prepared = await prepareHostInit(this.options, this.descriptor);
        if (this.destroyed) throw new Error('Editor was destroyed before activation completed');
        this.prepared = prepared;
        this.state = prepared.initialState;
        this.maybeStartRuntime();
        return await this.readyPromise;
      } catch (error) {
        const normalized = toError(error);
        if (!this.mountError) this.failBeforeReady(normalized);
        await this.destroy();
        this.mountPhase = 'error';
        this.mountError = normalized;
        throw normalized;
      }
    })();
    return this.activationPromise;
  }

  getMountState(): OfficeEditorMountState {
    return {
      id: this.id,
      origin: this.hostOrigin,
      phase: this.mountPhase,
      ...(this.mountError ? { error: this.mountError } : {}),
    };
  }

  private handleWindowMessage(event: MessageEvent): void {
    if (
      this.destroyed ||
      this.connected ||
      event.origin !== this.hostOrigin ||
      event.source !== this.iframe?.contentWindow
    ) {
      return;
    }
    if (!isOfficeHostMessage(event.data, this.id)) {
      return;
    }

    const message = event.data as OfficeHostWindowMessage;
    if (message.type !== 'HOST_READY') {
      return;
    }

    const expectedIdentity = this.options.expectedHostIdentity;
    if (expectedIdentity && !officeHostIdentitiesEqual(expectedIdentity, message.identity)) {
      this.failBeforeReady(new OfficeHostIdentityMismatchError(expectedIdentity, message.identity));
      return;
    }
    this.hostIdentity = message.identity;

    this.connected = true;
    this.clearHostReadyTimeout();
    this.removeWindowMessageListener();
    this.mountPhase = this.activationStarted ? 'runtime-loading' : 'waiting-for-activation';

    const channel = new MessageChannel();
    this.port = channel.port1;
    debugStats.activeHostPortCount += 1;
    debugStats.peakActiveHostPortCount = Math.max(debugStats.peakActiveHostPortCount, debugStats.activeHostPortCount);
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

    this.maybeStartRuntime();
  }

  private maybeStartRuntime(): void {
    if (!this.activationStarted || !this.connected || !this.prepared || this.initSent || this.destroyed) return;
    this.initSent = true;
    this.mountPhase = 'runtime-loading';
    this.resetStartupHeartbeatTimeout();
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
      case 'STARTUP_HEARTBEAT_PORT':
        this.installStartupHeartbeatPort(event.ports[0]);
        return;
      case 'STARTUP_PHASE':
      case 'STARTUP_HEARTBEAT':
        if (!this.mounted) this.resetStartupHeartbeatTimeout();
        return;
      case 'READY':
        this.applyHostState(message.state);
        this.mounted = true;
        this.mountPhase = 'ready';
        this.clearStartupWatchdogs();
        this.readyResolve?.(this);
        this.readyResolve = null;
        this.readyReject = null;
        this.maybeNotifyReady();
        return;
      case 'STATE':
        this.applyHostState(message.state);
        this.maybeNotifyReady();
        return;
      case 'SAVE_RESULT':
        void this.handleSaveResult(message);
        return;
      case 'DOWNLOAD_RESULT':
        void this.handleDownloadResult(message);
        return;
      case 'SAVE_AS_RESULT':
        void this.handleSaveAsResult(message);
        return;
      case 'PRINT_TITLE':
        setTemporaryDocumentTitle(message.title, message.durationMs);
        return;
      case 'CONFIRM_SAVE_TO_NEW_FORMAT_RESULT':
        this.handleConfirmSaveToNewFormatResult(message);
        return;
      case 'PLUGIN_READY':
        this.options.onPluginReady?.(message.pluginGuid, message.editorType, this);
        return;
      case 'PLUGIN_RESULT':
        this.handlePluginResult(message);
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

  private applyHostState(state: OfficeHostState | OfficeEditorState): void {
    const nextState = toPublicState(state);
    const dirtyChanged = nextState.dirty !== this.state.dirty;
    const stateChanged =
      nextState.mode !== this.state.mode ||
      nextState.readonly !== this.state.readonly ||
      nextState.status !== this.state.status ||
      nextState.destroyed !== this.state.destroyed ||
      nextState.dirty !== this.state.dirty;
    this.state = nextState;
    if (dirtyChanged) {
      void Promise.resolve(this.options.onDirtyChange?.(nextState.dirty, this)).catch((error) => {
        this.options.onError?.(toError(error), this);
      });
    }
    if (stateChanged) {
      void Promise.resolve(this.options.onStateChange?.(nextState, this)).catch((error) => {
        this.options.onError?.(toError(error), this);
      });
    }
  }

  private setDirty(dirty: boolean): void {
    if (this.state.dirty === dirty) return;
    this.state = { ...this.state, dirty };
    void Promise.resolve(this.options.onDirtyChange?.(dirty, this)).catch((error) => {
      this.options.onError?.(toError(error), this);
    });
  }

  private getSaveBehavior(): OfficeSaveBehavior {
    return this.options.saveBehavior || 'auto';
  }

  private async invokeSaveCallback(file: File): Promise<boolean> {
    const behavior = this.getSaveBehavior();
    const shouldCallCallback =
      behavior === 'callback' ||
      (behavior === 'auto' && (this.state.sourceKind !== 'new-document' || Boolean(this.options.onSave)));

    if (!shouldCallCallback) return false;
    if (!this.options.onSave) {
      throw new Error(
        'A save callback is required for this document source. Provide onSave or use saveBehavior: "download".',
      );
    }

    const handled = await this.options.onSave(file, this);
    return handled === true;
  }

  private async persistSavedFile(file: File): Promise<void> {
    const behavior = this.getSaveBehavior();
    const handledByCallback = await this.invokeSaveCallback(file);

    if (
      behavior === 'download' ||
      (behavior === 'auto' && this.state.sourceKind === 'new-document' && !handledByCallback)
    ) {
      downloadFile(file);
    }
  }

  private async persistDownloadedFile(file: File): Promise<void> {
    if (this.options.onDownload) {
      await this.options.onDownload(file, this);
      return;
    }
    downloadFile(file);
  }

  private async persistSaveAsFile(file: File): Promise<void> {
    if (this.options.onSaveAs) {
      await this.options.onSaveAs(file, this);
      return;
    }
    downloadFile(file);
  }

  private postSaveAck(requestId: string, ok: boolean, message?: string): void {
    this.postToHost({
      protocol: OFFICE_HOST_PROTOCOL,
      type: 'SAVE_ACK',
      sessionId: this.id,
      requestId,
      ok,
      message,
    });
  }

  private async handleSaveResult(message: Extract<OfficeHostChildMessage, { type: 'SAVE_RESULT' }>): Promise<void> {
    const file = new File([message.buffer], message.fileName, {
      type: message.mimeType || getSavedFileMimeType(message.fileName),
    });
    const requestId = message.requestId;
    const request = requestId ? this.pendingRequests.get(requestId) : undefined;
    if (requestId && request) {
      this.pendingRequests.delete(requestId);
    }

    try {
      await this.persistSavedFile(file);
      this.setDirty(false);
      request?.resolve(file);
      if (requestId) {
        this.postSaveAck(requestId, true);
      }
    } catch (error) {
      const normalized = toError(error);
      this.setDirty(true);
      request?.reject(normalized);
      this.options.onError?.(normalized, this);
      if (requestId) {
        this.postSaveAck(requestId, false, normalized.message);
      }
    }
  }

  private async handleDownloadResult(
    message: Extract<OfficeHostChildMessage, { type: 'DOWNLOAD_RESULT' }>,
  ): Promise<void> {
    const file = new File([message.buffer], message.fileName, {
      type: message.mimeType || getSavedFileMimeType(message.fileName),
    });

    try {
      await this.persistDownloadedFile(file);
    } catch (error) {
      const normalized = toError(error);
      this.options.onError?.(normalized, this);
    }
  }

  private async handleSaveAsResult(
    message: Extract<OfficeHostChildMessage, { type: 'SAVE_AS_RESULT' }>,
  ): Promise<void> {
    const file = new File([message.buffer], message.fileName, {
      type: message.mimeType || getSavedFileMimeType(message.fileName),
    });

    try {
      await this.persistSaveAsFile(file);
    } catch (error) {
      const normalized = toError(error);
      this.options.onError?.(normalized, this);
    }
  }

  private handleHostError(message: Extract<OfficeHostChildMessage, { type: 'ERROR' }>): void {
    const error = this.mounted
      ? new Error(message.message)
      : new OfficeEditorStartupError(message.phase, message.message);
    if (message.requestId) {
      const confirmRequest = this.pendingConfirmationRequests.get(message.requestId);
      if (confirmRequest) {
        this.pendingConfirmationRequests.delete(message.requestId);
        confirmRequest.reject(error);
        this.options.onError?.(error, this);
        return;
      }

      const request = this.pendingRequests.get(message.requestId);
      if (request) {
        this.pendingRequests.delete(message.requestId);
        request.reject(error);
        this.options.onError?.(error, this);
        return;
      }

      const pluginRequest = this.pendingPluginRequests.get(message.requestId);
      if (pluginRequest) {
        this.pendingPluginRequests.delete(message.requestId);
        window.clearTimeout(pluginRequest.timeoutId);
        pluginRequest.reject(error);
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

  private handleConfirmSaveToNewFormatResult(
    message: Extract<OfficeHostChildMessage, { type: 'CONFIRM_SAVE_TO_NEW_FORMAT_RESULT' }>,
  ): void {
    const request = this.pendingConfirmationRequests.get(message.requestId);
    if (!request) return;
    this.pendingConfirmationRequests.delete(message.requestId);
    request.resolve(message.confirmed);
  }

  private handlePluginResult(message: Extract<OfficeHostChildMessage, { type: 'PLUGIN_RESULT' }>): void {
    const request = this.pendingPluginRequests.get(message.requestId);
    if (!request) return;
    this.pendingPluginRequests.delete(message.requestId);
    window.clearTimeout(request.timeoutId);
    if (message.ok) request.resolve(message.result);
    else request.reject(new Error(message.error || 'Office plugin operation failed'));
  }

  private failBeforeReady(error: Error): void {
    if (this.mountError) return;
    this.state = { ...this.state, status: 'error' };
    this.mountPhase = 'error';
    this.mountError = error;
    this.clearHostReadyTimeout();
    this.clearStartupWatchdogs();
    this.readyReject?.(error);
    this.readyResolve = null;
    this.readyReject = null;
    this.options.onError?.(error, this);
    if (!this.activationStarted) {
      void this.destroy().finally(() => {
        this.mountPhase = 'error';
        this.mountError = error;
      });
    }
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
      this.parentWindow.clearTimeout(this.hostReadyTimeout);
      this.hostReadyTimeout = null;
    }
  }

  private startStartupWatchdogs(): void {
    if (this.startupTotalTimeout === null) {
      this.startupTotalTimeout = this.parentWindow.setTimeout(() => {
        this.failBeforeReady(
          new OfficeEditorStartupError('startup-total', 'Office editor startup exceeded five minutes'),
        );
      }, STARTUP_TOTAL_TIMEOUT_MS);
    }
    if (this.initSent) this.resetStartupHeartbeatTimeout();
  }

  private installStartupHeartbeatPort(port: MessagePort | undefined): void {
    if (!port || this.destroyed) {
      port?.close();
      return;
    }
    if (this.startupHeartbeatPort) {
      this.startupHeartbeatPort.close();
      debugStats.activeStartupHeartbeatPortCount = Math.max(0, debugStats.activeStartupHeartbeatPortCount - 1);
    }
    this.startupHeartbeatPort = port;
    debugStats.startupHeartbeatPortCount += 1;
    debugStats.activeStartupHeartbeatPortCount += 1;
    debugStats.peakActiveStartupHeartbeatPortCount = Math.max(
      debugStats.peakActiveStartupHeartbeatPortCount,
      debugStats.activeStartupHeartbeatPortCount,
    );
    port.onmessage = (event: MessageEvent<OfficeHostChildMessage>) => {
      if (!isOfficeHostMessage(event.data, this.id)) return;
      if (event.data.type === 'STARTUP_HEARTBEAT' && !this.mounted) {
        debugStats.startupHeartbeatCount += 1;
        this.resetStartupHeartbeatTimeout();
      }
    };
    port.start();
    this.resetStartupHeartbeatTimeout();
  }

  private resetStartupHeartbeatTimeout(): void {
    if (this.startupHeartbeatTimeout !== null) {
      this.parentWindow.clearTimeout(this.startupHeartbeatTimeout);
    }
    this.startupHeartbeatTimeout = this.parentWindow.setTimeout(() => {
      this.failBeforeReady(
        new OfficeEditorStartupError('startup-heartbeat', 'Office editor startup heartbeat timed out'),
      );
    }, STARTUP_HEARTBEAT_TIMEOUT_MS);
  }

  private clearStartupWatchdogs(): void {
    if (this.startupHeartbeatTimeout !== null) {
      this.parentWindow.clearTimeout(this.startupHeartbeatTimeout);
      this.startupHeartbeatTimeout = null;
    }
    if (this.startupTotalTimeout !== null) {
      this.parentWindow.clearTimeout(this.startupTotalTimeout);
      this.startupTotalTimeout = null;
    }
  }

  private removeWindowMessageListener(): void {
    if (this.windowMessageListener) {
      this.parentWindow.removeEventListener('message', this.windowMessageListener);
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

  invokePlugin(pluginGuid: string, payload: unknown): Promise<unknown> {
    if (this.destroyed || !this.port) {
      return Promise.reject(new Error('Editor is not open'));
    }
    if (!pluginGuid.trim()) {
      return Promise.reject(new Error('A plugin GUID is required'));
    }

    const requestId = nextRequestId(this.id);
    return new Promise<unknown>((resolve, reject) => {
      const timeoutId = window.setTimeout(() => {
        this.pendingPluginRequests.delete(requestId);
        reject(new Error(`Office plugin operation timed out: ${pluginGuid}`));
      }, PLUGIN_REQUEST_TIMEOUT_MS);
      this.pendingPluginRequests.set(requestId, { resolve, reject, timeoutId });
      this.postToHost({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'INVOKE_PLUGIN',
        sessionId: this.id,
        requestId,
        pluginGuid,
        payload,
      });
    });
  }

  confirmSaveToNewFormat(options?: OfficeSaveToNewFormatConfirmationOptions): Promise<boolean> {
    if (this.destroyed || !this.port) {
      return Promise.reject(new Error('Editor is not open'));
    }

    const requestId = nextRequestId(this.id);
    return new Promise<boolean>((resolve, reject) => {
      this.pendingConfirmationRequests.set(requestId, { resolve, reject });
      this.postToHost({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'CONFIRM_SAVE_TO_NEW_FORMAT',
        sessionId: this.id,
        requestId,
        options,
      });
    });
  }

  setReadonly(readonly: boolean): void {
    this.state = {
      ...this.state,
      readonly,
      mode: this.returnsToPreview && readonly ? 'preview' : readonly ? 'readonly' : 'edit',
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

  setInterfaceTheme(theme: OfficeInterfaceTheme): void {
    this.options.interfaceTheme = theme;
    if (this.prepared) this.prepared.options.interfaceTheme = theme;
    if (!this.destroyed && this.port) {
      this.postToHost({
        protocol: OFFICE_HOST_PROTOCOL,
        type: 'SET_INTERFACE_THEME',
        sessionId: this.id,
        interfaceTheme: theme,
      });
    }
  }

  getState(): OfficeEditorState {
    return { ...this.state };
  }

  getHostIdentity(): OfficeHostIdentity {
    if (!this.hostIdentity) {
      throw new Error('Office host identity is unavailable before the host is ready');
    }
    return { ...this.hostIdentity };
  }

  destroy(): Promise<void> {
    this.mountPhase = 'destroyed';
    if (this.destroyPromise) return this.destroyPromise;
    this.destroyed = true;
    this.originLeaseState = 'retiring';
    this.mounted = false;
    this.state = {
      ...this.state,
      status: 'destroyed',
      destroyed: true,
    };

    this.clearHostReadyTimeout();
    this.clearStartupWatchdogs();
    this.removeWindowMessageListener();
    this.readyReject?.(new Error('Editor was destroyed before it became ready'));
    this.readyResolve = null;
    this.readyReject = null;

    for (const request of this.pendingRequests.values()) {
      request.reject(new Error('Editor was destroyed before save completed'));
    }
    this.pendingRequests.clear();
    for (const request of this.pendingConfirmationRequests.values()) {
      request.reject(new Error('Editor was destroyed before confirmation completed'));
    }
    this.pendingConfirmationRequests.clear();
    for (const request of this.pendingPluginRequests.values()) {
      window.clearTimeout(request.timeoutId);
      request.reject(new Error('Editor was destroyed before plugin operation completed'));
    }
    this.pendingPluginRequests.clear();
    activeInstances.delete(this.id);

    // Keep the origin leased while the host destroys its runtime, closes its
    // broker ports, and resets the frame. Callers may ignore the returned
    // Promise, so releasing synchronously here would let a new session bind to
    // the same reusable origin while the old Service Worker session is still
    // retiring. Every teardown wait is bounded; finally prevents a failed
    // teardown from leaking the slot forever.
    this.destroyPromise = this.destroyHost().finally(() => this.releaseOriginLease());
    return this.destroyPromise;
  }

  private async destroyHost(): Promise<void> {
    const destroyTimeoutMs = this.options.destroyTimeoutMs ?? DESTROY_TIMEOUT_MS;
    let hostResetDone = false;
    let removeResetListener = () => {};

    try {
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
          this.parentWindow.addEventListener('message', listener);
          removeResetListener = () => this.parentWindow.removeEventListener('message', listener);
        });
        this.postToHost({
          protocol: OFFICE_HOST_PROTOCOL,
          type: 'DESTROY',
          sessionId: this.id,
        });
        await Promise.race([resetAck, portAck, delay(destroyTimeoutMs)]);
      }
    } catch {
      // Continue with the bounded frame reset below when a stale/closed port
      // rejects the cooperative teardown message.
    } finally {
      removeResetListener();
    }
    if (!hostResetDone) {
      debugStats.hostResetTimeoutCount += 1;
    }

    try {
      await this.forceRemoveIframe(hostResetDone);
    } catch {
      // forceRemoveIframe always detaches the browsing context in its finally
      // block. Teardown is best-effort and must not leak a retired slot.
    }
    if (this.port) {
      this.port.onmessage = null;
      try {
        this.port.close();
      } catch {
        // Ignore a port that the host already closed during teardown.
      }
      debugStats.activeHostPortCount = Math.max(0, debugStats.activeHostPortCount - 1);
    }
    this.port = null;
    if (this.startupHeartbeatPort) {
      this.startupHeartbeatPort.onmessage = null;
      try {
        this.startupHeartbeatPort.close();
      } catch {
        // Ignore a heartbeat port that the host already closed.
      }
      debugStats.activeStartupHeartbeatPortCount = Math.max(0, debugStats.activeStartupHeartbeatPortCount - 1);
    }
    this.startupHeartbeatPort = null;
    this.hostWindow = null;
    this.destroyAckResolve = null;

    if (this.options.hardResetOnLastDestroy && activeInstances.size === 0) {
      maybeHardResetPage();
    }
  }

  private async forceRemoveIframe(hostResetDone = false): Promise<void> {
    const iframe = this.iframe;
    this.iframe = null;
    this.unsandboxedHostFrameCleanup?.();
    this.unsandboxedHostFrameCleanup = null;
    if (!iframe) return;

    if (iframe.isConnected) {
      hideIframeForTeardown(iframe);
      try {
        if (!hostResetDone) {
          const resetUrl = this.getHostResetUrl();
          await this.navigateIframeForTeardown(iframe, resetUrl, RESET_NAVIGATION_TIMEOUT_MS);
        }
        await this.navigateIframeForTeardown(iframe, 'about:blank', BLANK_NAVIGATION_TIMEOUT_MS);
      } finally {
        iframe.remove();
      }
    }
  }

  private getHostResetUrl(): string {
    const resetUrl = new URL(HOST_SELF_RESET_PATH, this.descriptor.hostUrl.href);
    resetUrl.searchParams.set('sessionId', this.id);
    resetUrl.searchParams.set(
      'parentOrigin',
      readOfficeHostBootstrap(this.descriptor.hostUrl).parentOrigin || this.parentWindow.location.origin,
    );
    return resetUrl.href;
  }

  private async navigateIframeForTeardown(iframe: HTMLIFrameElement, url: string, timeoutMs: number): Promise<void> {
    hideIframeForTeardown(iframe);
    const loaded = new Promise<void>((resolve) => {
      iframe.addEventListener('load', () => resolve(), { once: true });
    });
    iframe.src = url;
    await Promise.race([loaded, delay(timeoutMs)]);
  }
}

export function mountOfficeEditor(container: HTMLElement, options: CreateOfficeEditorOptions): OfficeEditorMount {
  if (!isHTMLElementContainer(container)) {
    throw new Error('mountOfficeEditor requires an HTMLElement container');
  }

  try {
    const proxy = BrowserOfficeEditorProxy.mount(container, options);
    return {
      id: proxy.id,
      activate: () => proxy.activate(),
      destroy: () => proxy.destroy(),
      getState: () => proxy.getMountState(),
    };
  } catch (error) {
    const normalized = toError(error);
    options.onError?.(normalized);
    throw normalized;
  }
}

export async function createOfficeEditor(
  container: HTMLElement,
  options: CreateOfficeEditorOptions,
): Promise<OfficeEditorInstance> {
  const mount = mountOfficeEditor(container, options);
  try {
    return await mount.activate();
  } catch (error) {
    await mount.destroy();
    throw toError(error);
  }
}

export function loadOfficeEditorApi(): Promise<void> {
  return Promise.resolve();
}

export function getActiveOfficeEditorCount(): number {
  return activeInstances.size;
}
