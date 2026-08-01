import {
  RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES,
  RESOURCE_BROKER_MAX_READ_WINDOW_BYTES,
  OneTimeResourceBrokerCapabilityRegistry,
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerRequestId,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  parseResourceBrokerClientMessage,
  parseResourceBrokerServerMessage,
  type ResourceBrokerCapabilityConsumeResult,
  type ResourceBrokerCapabilityMetadata,
  type ResourceBrokerClientMessage,
  type ResourceBrokerProbeMessage,
  type ResourceBrokerReadMessage,
  type ResourceBrokerServerMessage,
} from './lib/resource-broker-protocol';
import {
  RESOURCE_BROKER_CANONICAL_ORIGIN,
  RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS,
  resolveResourceBrokerCanonicalOrigin,
  resolveResourceBrokerPhysicalEditorIdentity,
  resolveResourceBrokerTrustedParentIdentity,
  type ResourceBrokerFrameExpectedIdentity,
} from './lib/resource-broker-frame-client';
import {
  IndexedDbReleaseLeaseLedger,
  ReleaseLeaseHeartbeat,
  createReleaseLeaseBinding,
  type ReleaseLeaseStore,
} from './lib/release-lease-gc';

const RESOURCE_BROKER_SW_PATH = '/sw.js';
const RESOURCE_BROKER_SW_SCOPE = '/';
const RESOURCE_BROKER_CAPABILITY_TTL_MS = 30_000;
export const RESOURCE_BROKER_RELAY_MAX_ACTIVE_REQUESTS = 64;
export const RESOURCE_BROKER_RELAY_MAX_RESERVED_BYTES = 64 * 1024 * 1024;

export interface ResourceBrokerRelayMetrics {
  schemaVersion: 1;
  role: 'canonical-relay';
  activeRequests: number;
  peakActiveRequests: number;
  reservedBytes: number;
  peakReservedBytes: number;
  maxActiveRequests: number;
  maxReservedBytes: number;
  destroyed: boolean;
}

type BrokerRequest = ResourceBrokerProbeMessage | ResourceBrokerReadMessage;

export interface ResourceBrokerFrameIdentity extends ResourceBrokerFrameExpectedIdentity {
  physicalEditorOrigin: string;
  localMatrix: boolean;
}

export interface ResourceBrokerPortLike {
  onmessage: ((event: MessageEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  close(): void;
  postMessage(message: unknown, transfer?: Transferable[]): void;
  start(): void;
}

export function createResourceBrokerReleaseLease(
  identity: Pick<ResourceBrokerFrameIdentity, 'releaseId' | 'sessionId' | 'editorOrigin'>,
  options: {
    ledger?: ReleaseLeaseStore;
    indexedDb?: IDBFactory;
    randomUUID?: () => string;
    ttlMs?: number;
    heartbeatIntervalMs?: number;
    setTimeout?: typeof globalThis.setTimeout;
    clearTimeout?: typeof globalThis.clearTimeout;
    onError?: (error: unknown) => void;
  } = {},
): ReleaseLeaseHeartbeat {
  const ledger = options.ledger ?? new IndexedDbReleaseLeaseLedger(options.indexedDb);
  return new ReleaseLeaseHeartbeat({
    ledger,
    binding: createReleaseLeaseBinding(
      {
        releaseId: identity.releaseId,
        sessionId: identity.sessionId,
        editorOrigin: identity.editorOrigin,
      },
      options.randomUUID,
    ),
    ttlMs: options.ttlMs,
    heartbeatIntervalMs: options.heartbeatIntervalMs,
    setTimeout: options.setTimeout,
    clearTimeout: options.clearTimeout,
    onError: options.onError,
  });
}

export type ResourceBrokerWorkerPortOpener = (request: BrokerRequest) => Promise<ResourceBrokerPortLike>;

type ResourceBrokerRelayOptions = {
  releaseId: string;
  sessionId: string;
  openWorkerPort: ResourceBrokerWorkerPortOpener;
  postToClient: (message: ResourceBrokerServerMessage, transfer?: Transferable[]) => void;
  requestTimeoutMs?: number;
  maxActiveRequests?: number;
  maxReservedBytes?: number;
  setTimeout?: typeof globalThis.setTimeout;
  clearTimeout?: typeof globalThis.clearTimeout;
};

type RelayRequestState = {
  request: BrokerRequest;
  port: ResourceBrokerPortLike | null;
  pendingFlow: ResourceBrokerClientMessage[];
  timeout: ReturnType<typeof globalThis.setTimeout> | null;
  generation: number;
  headers: Extract<ResourceBrokerServerMessage, { type: 'HEADERS' }> | null;
  bytesReceived: number;
  reservedBytes: number;
  cancelRequested: boolean;
};

function exactSearchParams(search: string): URLSearchParams | null {
  const params = new URLSearchParams(search);
  const allowed = new Set(['releaseId', 'sessionId', 'parentOrigin', 'localMatrix']);
  for (const key of params.keys()) {
    if (!allowed.has(key) || params.getAll(key).length !== 1) return null;
  }
  if (!params.has('releaseId') || !params.has('sessionId') || !params.has('parentOrigin')) return null;
  const localMatrix = params.get('localMatrix');
  if (localMatrix !== null && localMatrix !== '1') return null;
  return params;
}

export function resolveResourceBrokerFrameIdentity(input: {
  locationOrigin: string;
  documentReferrer: string;
  search: string;
}): ResourceBrokerFrameIdentity | null {
  const params = exactSearchParams(input.search);
  if (!params) return null;
  const localMatrixRequested = params.get('localMatrix') === '1';
  const canonical = resolveResourceBrokerCanonicalOrigin(input.locationOrigin, localMatrixRequested);
  if (!canonical || canonical.localMatrix !== localMatrixRequested) return null;

  let referrerOrigin: string;
  try {
    referrerOrigin = new URL(input.documentReferrer).origin;
  } catch {
    return null;
  }
  const editor = resolveResourceBrokerPhysicalEditorIdentity(referrerOrigin, canonical.origin, localMatrixRequested);
  const rawParentOrigin = params.get('parentOrigin') ?? '';
  let parent = resolveResourceBrokerTrustedParentIdentity(rawParentOrigin, canonical.origin, localMatrixRequested);
  if (localMatrixRequested && !parent) {
    parent = resolveResourceBrokerTrustedParentIdentity(rawParentOrigin, RESOURCE_BROKER_CANONICAL_ORIGIN, false);
  }
  const releaseId = params.get('releaseId');
  const sessionId = params.get('sessionId');
  if (
    !editor ||
    !parent ||
    !releaseId ||
    !sessionId ||
    !isResourceBrokerReleaseId(releaseId) ||
    !isResourceBrokerSessionId(sessionId)
  ) {
    return null;
  }
  return {
    canonicalOrigin: canonical.origin,
    physicalEditorOrigin: editor.physicalOrigin,
    editorOrigin: editor.logicalOrigin,
    parentOrigin: parent.logicalOrigin,
    releaseId,
    sessionId,
    localMatrix: canonical.localMatrix,
  };
}

export class ResourceBrokerFrameCapabilityGate {
  readonly #identity: ResourceBrokerFrameIdentity;
  readonly #registry: OneTimeResourceBrokerCapabilityRegistry;
  #issued: ResourceBrokerCapabilityMetadata | null = null;
  #connected = false;

  constructor(identity: ResourceBrokerFrameIdentity, registry = new OneTimeResourceBrokerCapabilityRegistry()) {
    this.#identity = identity;
    this.#registry = registry;
  }

  issue(): ResourceBrokerCapabilityMetadata {
    if (this.#issued || this.#connected) throw new Error('Resource Broker capability was already issued');
    this.#issued = this.#registry.issue({
      parentOrigin: this.#identity.parentOrigin,
      editorOrigin: this.#identity.editorOrigin,
      releaseId: this.#identity.releaseId,
      sessionId: this.#identity.sessionId,
      ttlMs: RESOURCE_BROKER_CAPABILITY_TTL_MS,
    });
    return this.#issued;
  }

  consume(value: unknown, physicalEventOrigin: string, nowMs = Date.now()): ResourceBrokerCapabilityConsumeResult {
    if (this.#connected || physicalEventOrigin !== this.#identity.physicalEditorOrigin) {
      return { ok: false, code: 'mismatch' };
    }
    const parsed = parseResourceBrokerClientMessage(value);
    if (!parsed || parsed.type !== 'CONNECT') return { ok: false, code: 'invalid' };
    const result = this.#registry.consume(parsed.capability, this.#identity.editorOrigin, nowMs);
    if (result.ok) {
      this.#connected = true;
      this.#issued = null;
    }
    return result;
  }
}

export class ResourceBrokerRelaySession {
  readonly #releaseId: string;
  readonly #sessionId: string;
  readonly #openWorkerPort: ResourceBrokerWorkerPortOpener;
  readonly #postToClient: ResourceBrokerRelayOptions['postToClient'];
  readonly #requestTimeoutMs: number;
  readonly #maxActiveRequests: number;
  readonly #maxReservedBytes: number;
  readonly #setTimeout: typeof globalThis.setTimeout;
  readonly #clearTimeout: typeof globalThis.clearTimeout;
  readonly #requests = new Map<string, RelayRequestState>();
  #reservedBytes = 0;
  #peakActiveRequests = 0;
  #peakReservedBytes = 0;
  #generation = 0;
  #destroyed = false;

  constructor(options: ResourceBrokerRelayOptions) {
    if (!isResourceBrokerReleaseId(options.releaseId) || !isResourceBrokerSessionId(options.sessionId)) {
      throw new TypeError('Invalid Resource Broker relay identity');
    }
    const timeoutMs = options.requestTimeoutMs ?? RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS;
    if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS) {
      throw new TypeError('Invalid Resource Broker relay timeout');
    }
    this.#releaseId = options.releaseId;
    this.#sessionId = options.sessionId;
    this.#openWorkerPort = options.openWorkerPort;
    this.#postToClient = options.postToClient;
    this.#requestTimeoutMs = timeoutMs;
    this.#maxActiveRequests = options.maxActiveRequests ?? RESOURCE_BROKER_RELAY_MAX_ACTIVE_REQUESTS;
    this.#maxReservedBytes = options.maxReservedBytes ?? RESOURCE_BROKER_RELAY_MAX_RESERVED_BYTES;
    if (
      !Number.isSafeInteger(this.#maxActiveRequests) ||
      this.#maxActiveRequests <= 0 ||
      this.#maxActiveRequests > RESOURCE_BROKER_RELAY_MAX_ACTIVE_REQUESTS ||
      !Number.isSafeInteger(this.#maxReservedBytes) ||
      this.#maxReservedBytes < RESOURCE_BROKER_MAX_READ_WINDOW_BYTES ||
      this.#maxReservedBytes > RESOURCE_BROKER_RELAY_MAX_RESERVED_BYTES
    ) {
      throw new TypeError('Invalid Resource Broker relay limits');
    }
    this.#setTimeout = options.setTimeout ?? globalThis.setTimeout.bind(globalThis);
    this.#clearTimeout = options.clearTimeout ?? globalThis.clearTimeout.bind(globalThis);
  }

  get activeRequestCount(): number {
    return this.#requests.size;
  }

  get reservedBytes(): number {
    return this.#reservedBytes;
  }

  get metrics(): ResourceBrokerRelayMetrics {
    return {
      schemaVersion: 1,
      role: 'canonical-relay',
      activeRequests: this.#requests.size,
      peakActiveRequests: this.#peakActiveRequests,
      reservedBytes: this.#reservedBytes,
      peakReservedBytes: this.#peakReservedBytes,
      maxActiveRequests: this.#maxActiveRequests,
      maxReservedBytes: this.#maxReservedBytes,
      destroyed: this.#destroyed,
    };
  }

  handleClientMessage(value: unknown): boolean {
    if (this.#destroyed) return false;
    const message = parseResourceBrokerClientMessage(value);
    if (!message || message.type === 'CONNECT') return false;
    if (message.type === 'PROBE' || message.type === 'READ') {
      this.#startRequest(message);
      return true;
    }
    const state = this.#requests.get(message.id);
    if (!state) return false;
    if (message.type === 'CANCEL') {
      if (state.cancelRequested) return true;
      state.cancelRequested = true;
      this.#refreshTimeout(state);
    } else if (state.cancelRequested) {
      return false;
    }
    if (state.port) {
      try {
        state.port.postMessage(message);
      } catch {
        this.#fail(state, 'protocol');
      }
    } else {
      if (message.type === 'CANCEL') {
        state.pendingFlow.splice(0, state.pendingFlow.length, message);
      } else if (state.pendingFlow.some((pending) => pending.type === 'PULL' || pending.type === 'CANCEL')) {
        this.#fail(state, 'busy');
      } else {
        state.pendingFlow.push(message);
      }
    }
    return true;
  }

  notifyControllerChange(): void {
    if (this.#destroyed) return;
    this.#generation += 1;
    for (const state of this.#requests.values()) {
      this.#cancelWorker(state);
      this.#fail(state, 'protocol');
    }
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#generation += 1;
    for (const state of this.#requests.values()) {
      this.#cancelWorker(state);
      this.#fail(state, 'cancelled');
    }
  }

  #startRequest(request: BrokerRequest): void {
    if (request.releaseId !== this.#releaseId || request.sessionId !== this.#sessionId) {
      this.#postError(request.id, 'release');
      return;
    }
    if (this.#requests.has(request.id)) {
      this.#postError(request.id, 'busy');
      return;
    }
    const reservedBytes =
      request.type === 'READ' ? (request.windowBytes ?? RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES) : 0;
    if (
      this.#requests.size >= this.#maxActiveRequests ||
      this.#reservedBytes + reservedBytes > this.#maxReservedBytes
    ) {
      this.#postError(request.id, 'busy');
      return;
    }
    const state: RelayRequestState = {
      request,
      port: null,
      pendingFlow: [],
      timeout: null,
      generation: this.#generation,
      headers: null,
      bytesReceived: 0,
      reservedBytes,
      cancelRequested: false,
    };
    this.#requests.set(request.id, state);
    this.#reservedBytes += reservedBytes;
    this.#peakActiveRequests = Math.max(this.#peakActiveRequests, this.#requests.size);
    this.#peakReservedBytes = Math.max(this.#peakReservedBytes, this.#reservedBytes);
    this.#refreshTimeout(state);

    Promise.resolve()
      .then(() => this.#openWorkerPort(request))
      .then((port) => {
        if (this.#destroyed || state.generation !== this.#generation || this.#requests.get(request.id) !== state) {
          port.close();
          return;
        }
        state.port = port;
        port.onmessage = (event) => this.#handleWorkerMessage(state, event.data);
        port.onmessageerror = () => this.#fail(state, 'protocol');
        port.start();
        for (const pending of state.pendingFlow.splice(0)) {
          port.postMessage(pending);
        }
      })
      .catch(() => this.#fail(state, 'storage'));
  }

  #handleWorkerMessage(state: RelayRequestState, value: unknown): void {
    if (this.#requests.get(state.request.id) !== state) return;
    const message = parseResourceBrokerServerMessage(value);
    if (!message || message.id !== state.request.id) {
      this.#fail(state, 'protocol');
      return;
    }
    if (state.cancelRequested) {
      if (message.type === 'CHUNK' || message.type === 'HEADERS') {
        this.#refreshTimeout(state);
        return;
      }
      if (message.type === 'CANCELLED' || message.type === 'END' || message.type === 'ERROR') {
        this.#postToClient({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CANCELLED',
          id: state.request.id,
          bytesSent: state.bytesReceived,
        });
        this.#finish(state);
        return;
      }
      this.#fail(state, 'protocol');
      return;
    }
    if (state.request.type === 'PROBE') {
      if (message.type !== 'RESULT' && message.type !== 'ERROR') {
        this.#fail(state, 'protocol');
        return;
      }
      if (message.type === 'RESULT' && message.value.releaseId !== this.#releaseId) {
        this.#fail(state, 'release');
        return;
      }
    } else {
      if (message.type === 'RESULT') {
        this.#fail(state, 'protocol');
        return;
      }
      if (message.type === 'HEADERS') {
        if (state.headers) {
          this.#fail(state, 'protocol');
          return;
        }
        state.headers = message;
      } else if (message.type === 'CHUNK') {
        if (!state.headers || state.headers.status === 416) {
          this.#fail(state, 'protocol');
          return;
        }
        state.bytesReceived += message.bytes.byteLength;
        if (state.bytesReceived > state.headers.headers.contentLength) {
          this.#fail(state, 'integrity');
          return;
        }
      } else if (message.type === 'END') {
        if (
          !state.headers ||
          message.bytesSent !== state.bytesReceived ||
          state.bytesReceived !== state.headers.headers.contentLength
        ) {
          this.#fail(state, 'integrity');
          return;
        }
      }
    }

    this.#refreshTimeout(state);
    const transfer = message.type === 'CHUNK' ? [message.bytes] : undefined;
    this.#postToClient(message, transfer);
    if (
      message.type === 'RESULT' ||
      message.type === 'END' ||
      message.type === 'CANCELLED' ||
      message.type === 'ERROR'
    ) {
      this.#finish(state);
    }
  }

  #refreshTimeout(state: RelayRequestState): void {
    if (state.timeout !== null) this.#clearTimeout(state.timeout);
    state.timeout = this.#setTimeout(() => {
      this.#cancelWorker(state);
      this.#fail(state, 'timeout');
    }, this.#requestTimeoutMs);
  }

  #cancelWorker(state: RelayRequestState): void {
    if (!state.port) return;
    try {
      state.port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CANCEL',
        id: state.request.id,
      });
    } catch {
      // The worker channel is already gone; local cleanup still completes.
    }
  }

  #postError(id: string, code: Extract<ResourceBrokerServerMessage, { type: 'ERROR' }>['code']): void {
    if (!isResourceBrokerRequestId(id)) return;
    this.#postToClient({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'ERROR',
      id,
      code,
    });
  }

  #fail(state: RelayRequestState, code: Extract<ResourceBrokerServerMessage, { type: 'ERROR' }>['code']): void {
    if (this.#requests.get(state.request.id) !== state) return;
    this.#postError(state.request.id, code);
    this.#finish(state);
  }

  #finish(state: RelayRequestState): void {
    if (this.#requests.get(state.request.id) !== state) return;
    this.#requests.delete(state.request.id);
    this.#reservedBytes = Math.max(0, this.#reservedBytes - state.reservedBytes);
    if (state.timeout !== null) this.#clearTimeout(state.timeout);
    state.timeout = null;
    state.pendingFlow.length = 0;
    state.port?.close();
    state.port = null;
  }
}

async function waitForCanonicalWorker(serviceWorkers: ServiceWorkerContainer, online: boolean): Promise<ServiceWorker> {
  let registration: ServiceWorkerRegistration | undefined;
  if (online) {
    try {
      registration = await serviceWorkers.register(RESOURCE_BROKER_SW_PATH, {
        scope: RESOURCE_BROKER_SW_SCOPE,
        updateViaCache: 'none',
      });
    } catch {
      registration = await serviceWorkers.getRegistration(RESOURCE_BROKER_SW_SCOPE);
    }
  } else {
    registration = await serviceWorkers.getRegistration(RESOURCE_BROKER_SW_SCOPE);
  }
  if (!registration) throw new Error('Resource Broker Service Worker is unavailable');
  await serviceWorkers.ready;
  const worker = serviceWorkers.controller ?? registration.active;
  if (!worker) throw new Error('Resource Broker Service Worker is not active');
  const scriptUrl = new URL(worker.scriptURL);
  if (scriptUrl.origin !== location.origin || scriptUrl.pathname !== RESOURCE_BROKER_SW_PATH) {
    throw new Error('Resource Broker Service Worker identity mismatch');
  }
  return worker;
}

export function createCanonicalWorkerPortOpener(
  serviceWorkers: ServiceWorkerContainer,
  online: () => boolean,
): ResourceBrokerWorkerPortOpener {
  return async (request) => {
    const worker = await waitForCanonicalWorker(serviceWorkers, online());
    const channel = new MessageChannel();
    worker.postMessage(request, [channel.port2]);
    return channel.port1;
  };
}

export async function probeCanonicalResourceBroker(
  openWorkerPort: ResourceBrokerWorkerPortOpener,
  identity: Pick<ResourceBrokerFrameIdentity, 'releaseId' | 'sessionId'>,
  timeoutMs = RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS,
): Promise<void> {
  const id = `bootstrap-${crypto.randomUUID()}`;
  const request: ResourceBrokerProbeMessage = {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'PROBE',
    id,
    releaseId: identity.releaseId,
    sessionId: identity.sessionId,
  };
  const port = await openWorkerPort(request);
  await new Promise<void>((resolve, reject) => {
    const timeout = globalThis.setTimeout(() => {
      port.close();
      reject(new Error('Resource Broker probe timed out'));
    }, timeoutMs);
    const finish = (callback: () => void) => {
      globalThis.clearTimeout(timeout);
      port.close();
      callback();
    };
    port.onmessage = (event) => {
      const message = parseResourceBrokerServerMessage(event.data);
      if (
        !message ||
        message.id !== id ||
        (message.type !== 'RESULT' && message.type !== 'ERROR') ||
        (message.type === 'RESULT' && message.value.releaseId !== identity.releaseId)
      ) {
        finish(() => reject(new Error('Invalid Resource Broker probe response')));
        return;
      }
      if (message.type === 'ERROR') {
        finish(() => reject(new Error(`Resource Broker probe failed: ${message.code}`)));
        return;
      }
      finish(resolve);
    };
    port.onmessageerror = () => finish(() => reject(new Error('Resource Broker probe channel failed')));
    port.start();
  });
}

function connectionError(port: MessagePort, code: 'capability' | 'identity' | 'worker' | 'timeout'): void {
  port.postMessage({
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'CONNECT_ERROR',
    code,
  });
  port.close();
}

export async function startResourceBrokerFrame(
  browserWindow: Window = window,
  browserDocument: Document = document,
): Promise<() => void> {
  const identity = resolveResourceBrokerFrameIdentity({
    locationOrigin: browserWindow.location.origin,
    documentReferrer: browserDocument.referrer,
    search: browserWindow.location.search,
  });
  if (!identity || !('serviceWorker' in navigator)) {
    throw new Error('Resource Broker frame identity is not trusted');
  }

  const releaseLease = createResourceBrokerReleaseLease(identity);
  await releaseLease.start();
  const openWorkerPort = createCanonicalWorkerPortOpener(navigator.serviceWorker, () => navigator.onLine);

  const gate = new ResourceBrokerFrameCapabilityGate(identity);
  const capability = gate.issue();
  let session: ResourceBrokerRelaySession | null = null;
  let connectionPort: MessagePort | null = null;
  let destroyed = false;
  const metricsTarget = browserWindow as Window & {
    __ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__?: () => {
      schemaVersion: 1;
      role: 'canonical-relay-frame';
      connected: boolean;
      relay: ResourceBrokerRelayMetrics | null;
    };
  };
  metricsTarget.__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__ = () => ({
    schemaVersion: 1,
    role: 'canonical-relay-frame',
    connected: Boolean(connectionPort && session),
    relay: session?.metrics ?? null,
  });

  const onControllerChange = () => session?.notifyControllerChange();
  navigator.serviceWorker.addEventListener('controllerchange', onControllerChange);

  const onWindowMessage = (event: MessageEvent) => {
    const port = event.ports[0];
    if (
      destroyed ||
      !port ||
      event.ports.length !== 1 ||
      event.source !== browserWindow.parent ||
      event.origin !== identity.physicalEditorOrigin
    ) {
      return;
    }
    const connect = parseResourceBrokerClientMessage(event.data);
    if (!connect || connect.type !== 'CONNECT') {
      connectionError(port, 'identity');
      return;
    }
    const result = gate.consume(connect, event.origin);
    if (!result.ok || connectionPort) {
      connectionError(port, 'capability');
      return;
    }

    connectionPort = port;
    session = new ResourceBrokerRelaySession({
      releaseId: identity.releaseId,
      sessionId: identity.sessionId,
      openWorkerPort,
      postToClient(message, transfer) {
        connectionPort?.postMessage(message, transfer ?? []);
      },
    });
    port.onmessage = (portEvent) => session?.handleClientMessage(portEvent.data);
    port.onmessageerror = () => {
      session?.destroy();
      session = null;
      connectionPort?.close();
      connectionPort = null;
    };
    port.start();
    port.postMessage({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CONNECTED',
      releaseId: identity.releaseId,
      sessionId: identity.sessionId,
    });
  };
  browserWindow.addEventListener('message', onWindowMessage);

  browserWindow.parent.postMessage(
    {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHALLENGE',
      capability,
    },
    identity.physicalEditorOrigin,
  );

  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    browserWindow.removeEventListener('message', onWindowMessage);
    navigator.serviceWorker.removeEventListener('controllerchange', onControllerChange);
    session?.destroy();
    session = null;
    connectionPort?.close();
    connectionPort = null;
    delete metricsTarget.__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__;
    void releaseLease.stop().catch(() => undefined);
  };
  browserWindow.addEventListener('pagehide', destroy, { once: true });
  return destroy;
}

if (
  typeof window !== 'undefined' &&
  typeof document !== 'undefined' &&
  document.documentElement.dataset.onlyofficeResourceBroker === 'true'
) {
  void startResourceBrokerFrame().catch((error) => {
    console.error('[onlyoffice-browser] Failed to start the canonical Resource Broker frame', error);
  });
}
