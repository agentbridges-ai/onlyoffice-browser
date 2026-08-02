import {
  RESOURCE_BROKER_MAX_READ_WINDOW_BYTES,
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  normalizeResourceBrokerReadWindowBytes,
  normalizeResourceBrokerResourcePath,
  parseResourceBrokerRangeHeader,
  parseResourceBrokerServerMessage,
  resolveResourceBrokerRange,
  type ResourceBrokerByteRangeRequest,
  type ResourceBrokerErrorMessage,
  type ResourceBrokerEndMessage,
  type ResourceBrokerHeadersMessage,
  type ResourceBrokerServerMessage,
} from './resource-broker-protocol';

export const EDITOR_RESOURCE_BROKER_BIND_TYPE = 'ONLYOFFICE_BIND_BROKER' as const;
export const EDITOR_RESOURCE_BROKER_UNBIND_TYPE = 'ONLYOFFICE_UNBIND_BROKER' as const;
export const EDITOR_RESOURCE_BROKER_BOUND_TYPE = 'ONLYOFFICE_BROKER_BOUND' as const;
export const EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE = 'ONLYOFFICE_BROKER_BIND_ERROR' as const;
export const EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS = 30_000;
export const EDITOR_RESOURCE_BROKER_MAX_ACTIVE_REQUESTS = 64;
export const EDITOR_RESOURCE_BROKER_MAX_RESERVED_BYTES = 64 * 1024 * 1024;

export interface EditorResourceBrokerIdentity {
  releaseId: string;
  sessionId: string;
}

export interface EditorResourceBrokerBindMessage extends EditorResourceBrokerIdentity {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: typeof EDITOR_RESOURCE_BROKER_BIND_TYPE;
}

export interface EditorResourceBrokerUnbindMessage extends EditorResourceBrokerIdentity {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: typeof EDITOR_RESOURCE_BROKER_UNBIND_TYPE;
}

export interface EditorResourceBrokerBoundMessage extends EditorResourceBrokerIdentity {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: typeof EDITOR_RESOURCE_BROKER_BOUND_TYPE;
  ok: true;
}

export interface EditorResourceBrokerBindErrorMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: typeof EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE;
  ok: false;
  code: 'identity' | 'ports' | 'protocol';
}

export type EditorResourceBrokerConnectionState =
  | {
      status: 'disconnected' | 'connecting';
      identity: null;
      generation: number;
    }
  | {
      status: 'connected';
      identity: EditorResourceBrokerIdentity;
      generation: number;
    };

export type EditorResourceBrokerErrorCode =
  | ResourceBrokerErrorMessage['code']
  | 'connection'
  | 'identity'
  | 'request'
  | 'replaced';

export type EditorResourceBrokerErrorStage = 'bind' | 'connection' | 'request' | 'response';

export class EditorResourceBrokerError extends Error {
  readonly code: EditorResourceBrokerErrorCode;
  readonly stage: EditorResourceBrokerErrorStage;

  constructor(code: EditorResourceBrokerErrorCode, stage: EditorResourceBrokerErrorStage, message?: string) {
    super(message ?? `Editor Resource Broker failed (${stage}:${code})`);
    this.name = 'EditorResourceBrokerError';
    this.code = code;
    this.stage = stage;
  }
}

export interface EditorResourceBrokerPort {
  onmessage: ((event: MessageEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  close(): void;
  postMessage(message: unknown, transfer?: Transferable[]): void;
  start(): void;
  addEventListener?(type: string, listener: EventListener): void;
  removeEventListener?(type: string, listener: EventListener): void;
}

export interface EditorResourceBrokerFetchRequest {
  readonly method: string;
  readonly headers: Pick<Headers, 'get'>;
  readonly signal?: AbortSignal;
}

export interface EditorResourceBrokerFetchOptions {
  identity?: EditorResourceBrokerIdentity;
  windowBytes?: number;
  connectionTimeoutMs?: number;
}

export interface EditorResourceBrokerClientOptions {
  requestConnection?: (identity: EditorResourceBrokerIdentity | null) => void | Promise<void>;
  requestTimeoutMs?: number;
  maxActiveRequests?: number;
  maxReservedBytes?: number;
  setTimeout?: typeof globalThis.setTimeout;
  clearTimeout?: typeof globalThis.clearTimeout;
}

export interface EditorResourceBrokerMetrics {
  schemaVersion: 1;
  role: 'editor-service';
  connectionStatus: EditorResourceBrokerConnectionState['status'];
  connectionGeneration: number;
  activeRequests: number;
  peakActiveRequests: number;
  reservedBytes: number;
  peakReservedBytes: number;
  maxActiveRequests: number;
  maxReservedBytes: number;
}

type ConnectionWaiter = {
  identity: EditorResourceBrokerIdentity | null;
  resolve(identity: EditorResourceBrokerIdentity): void;
  reject(error: EditorResourceBrokerError): void;
  timeout: ReturnType<typeof globalThis.setTimeout>;
};

type PendingRead = {
  id: string;
  identity: EditorResourceBrokerIdentity;
  generation: number;
  windowBytes: number;
  range: ResourceBrokerByteRangeRequest | undefined;
  method: 'GET' | 'HEAD';
  resolveResponse(response: Response): void;
  rejectResponse(error: EditorResourceBrokerError): void;
  responseSettled: boolean;
  headers: ResourceBrokerHeadersMessage | null;
  controller: ReadableStreamDefaultController<Uint8Array> | null;
  pullOutstanding: boolean;
  cancelRequested: boolean;
  receivedBytes: number;
  finished: boolean;
  timeout: ReturnType<typeof globalThis.setTimeout> | null;
  signal: AbortSignal | undefined;
  abortListener: (() => void) | null;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function hasExactKeys(value: Record<string, unknown>, keys: readonly string[]): boolean {
  const actual = Object.keys(value);
  return actual.length === keys.length && keys.every((key) => Object.hasOwn(value, key));
}

function isPort(value: unknown): value is EditorResourceBrokerPort {
  if (!isRecord(value)) return false;
  return (
    typeof value.postMessage === 'function' &&
    typeof value.start === 'function' &&
    typeof value.close === 'function' &&
    'onmessage' in value &&
    'onmessageerror' in value
  );
}

function sameIdentity(left: EditorResourceBrokerIdentity, right: EditorResourceBrokerIdentity): boolean {
  return left.releaseId === right.releaseId && left.sessionId === right.sessionId;
}

function validateIdentity(value: EditorResourceBrokerIdentity): boolean {
  return isResourceBrokerReleaseId(value.releaseId) && isResourceBrokerSessionId(value.sessionId);
}

export function parseEditorResourceBrokerBindMessage(value: unknown): EditorResourceBrokerBindMessage | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['protocol', 'type', 'releaseId', 'sessionId']) ||
    value.protocol !== RESOURCE_BROKER_PROTOCOL ||
    value.type !== EDITOR_RESOURCE_BROKER_BIND_TYPE ||
    !isResourceBrokerReleaseId(value.releaseId) ||
    !isResourceBrokerSessionId(value.sessionId)
  ) {
    return null;
  }
  return {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: EDITOR_RESOURCE_BROKER_BIND_TYPE,
    releaseId: value.releaseId,
    sessionId: value.sessionId,
  };
}

export function parseEditorResourceBrokerUnbindMessage(value: unknown): EditorResourceBrokerUnbindMessage | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['protocol', 'type', 'releaseId', 'sessionId']) ||
    value.protocol !== RESOURCE_BROKER_PROTOCOL ||
    value.type !== EDITOR_RESOURCE_BROKER_UNBIND_TYPE ||
    !isResourceBrokerReleaseId(value.releaseId) ||
    !isResourceBrokerSessionId(value.sessionId)
  ) {
    return null;
  }
  return {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: EDITOR_RESOURCE_BROKER_UNBIND_TYPE,
    releaseId: value.releaseId,
    sessionId: value.sessionId,
  };
}

function responseHeaders(message: ResourceBrokerHeadersMessage): Headers {
  const headers = new Headers({
    'Accept-Ranges': message.headers.acceptRanges,
    'Content-Length': String(message.headers.contentLength),
    'Content-Type': message.headers.contentType,
  });
  if (message.headers.contentRange !== null) {
    headers.set('Content-Range', message.headers.contentRange);
  }
  return headers;
}

function validateHeadersForRequest(
  message: ResourceBrokerHeadersMessage,
  range: ResourceBrokerByteRangeRequest | undefined,
): boolean {
  const { status, headers } = message;
  if (status === 200) {
    return range === undefined && headers.contentRange === null;
  }
  if (status === 206) {
    if (range === undefined || headers.contentRange === null) return false;
    const match = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(headers.contentRange);
    if (!match) return false;
    const start = Number(match[1]);
    const end = Number(match[2]);
    const total = Number(match[3]);
    if (
      !Number.isSafeInteger(start) ||
      !Number.isSafeInteger(end) ||
      !Number.isSafeInteger(total) ||
      start < 0 ||
      end < start ||
      total <= end ||
      headers.contentLength !== end - start + 1
    ) {
      return false;
    }
    const expected = resolveResourceBrokerRange(range, total);
    return (
      expected.status === 206 &&
      expected.start === start &&
      expected.end === end &&
      expected.contentLength === headers.contentLength &&
      expected.contentRange === headers.contentRange
    );
  }
  if (range === undefined || headers.contentLength !== 0) return false;
  const match = headers.contentRange === null ? null : /^bytes \*\/(\d+)$/.exec(headers.contentRange);
  if (!match) return false;
  const total = Number(match[1]);
  return Number.isSafeInteger(total) && total >= 0 && resolveResourceBrokerRange(range, total).status === 416;
}

function toBrokerError(message: ResourceBrokerErrorMessage): EditorResourceBrokerError {
  return new EditorResourceBrokerError(message.code, 'response');
}

export class EditorResourceBrokerClient {
  readonly #requestConnectionCallback: EditorResourceBrokerClientOptions['requestConnection'];
  readonly #requestTimeoutMs: number;
  readonly #maxActiveRequests: number;
  readonly #maxReservedBytes: number;
  readonly #setTimeout: typeof globalThis.setTimeout;
  readonly #clearTimeout: typeof globalThis.clearTimeout;
  readonly #pending = new Map<string, PendingRead>();
  readonly #connectionWaiters = new Set<ConnectionWaiter>();
  #port: EditorResourceBrokerPort | null = null;
  #identity: EditorResourceBrokerIdentity | null = null;
  #status: EditorResourceBrokerConnectionState['status'] = 'disconnected';
  #generation = 0;
  #requestCounter = 0;
  #reservedBytes = 0;
  #peakActiveRequests = 0;
  #peakReservedBytes = 0;
  #connectionRequestInFlight = false;
  #portCloseListener: EventListener | null = null;

  constructor(options: EditorResourceBrokerClientOptions = {}) {
    this.#requestConnectionCallback = options.requestConnection;
    this.#requestTimeoutMs = options.requestTimeoutMs ?? EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS;
    this.#maxActiveRequests = options.maxActiveRequests ?? EDITOR_RESOURCE_BROKER_MAX_ACTIVE_REQUESTS;
    this.#maxReservedBytes = options.maxReservedBytes ?? EDITOR_RESOURCE_BROKER_MAX_RESERVED_BYTES;
    if (
      !Number.isSafeInteger(this.#requestTimeoutMs) ||
      this.#requestTimeoutMs <= 0 ||
      this.#requestTimeoutMs > EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS ||
      !Number.isSafeInteger(this.#maxActiveRequests) ||
      this.#maxActiveRequests <= 0 ||
      this.#maxActiveRequests > EDITOR_RESOURCE_BROKER_MAX_ACTIVE_REQUESTS ||
      !Number.isSafeInteger(this.#maxReservedBytes) ||
      this.#maxReservedBytes < RESOURCE_BROKER_MAX_READ_WINDOW_BYTES ||
      this.#maxReservedBytes > EDITOR_RESOURCE_BROKER_MAX_RESERVED_BYTES
    ) {
      throw new TypeError('Invalid Editor Resource Broker limits');
    }
    this.#setTimeout = options.setTimeout ?? globalThis.setTimeout.bind(globalThis);
    this.#clearTimeout = options.clearTimeout ?? globalThis.clearTimeout.bind(globalThis);
  }

  get connectionState(): EditorResourceBrokerConnectionState {
    if (this.#status === 'connected' && this.#identity) {
      return {
        status: 'connected',
        identity: { ...this.#identity },
        generation: this.#generation,
      };
    }
    return {
      status: this.#status === 'connecting' ? 'connecting' : 'disconnected',
      identity: null,
      generation: this.#generation,
    };
  }

  get activeRequestCount(): number {
    return this.#pending.size;
  }

  get reservedBytes(): number {
    return this.#reservedBytes;
  }

  get metrics(): EditorResourceBrokerMetrics {
    return {
      schemaVersion: 1,
      role: 'editor-service',
      connectionStatus: this.connectionState.status,
      connectionGeneration: this.#generation,
      activeRequests: this.#pending.size,
      peakActiveRequests: this.#peakActiveRequests,
      reservedBytes: this.#reservedBytes,
      peakReservedBytes: this.#peakReservedBytes,
      maxActiveRequests: this.#maxActiveRequests,
      maxReservedBytes: this.#maxReservedBytes,
    };
  }

  handleBindMessage(
    value: unknown,
    ports: readonly unknown[],
    beforeAcknowledge?: (identity: EditorResourceBrokerIdentity) => void,
  ): boolean {
    const brokerPort = ports.length === 2 ? ports[0] : null;
    const replyPort = ports.length === 2 ? ports[1] : null;
    if (!isPort(brokerPort) || !isPort(replyPort) || brokerPort === replyPort) {
      for (const port of ports) {
        if (isPort(port)) port.close();
      }
      return false;
    }
    return this.bind(value, brokerPort, replyPort, beforeAcknowledge);
  }

  bind(
    value: unknown,
    brokerPort: EditorResourceBrokerPort,
    replyPort: EditorResourceBrokerPort,
    beforeAcknowledge?: (identity: EditorResourceBrokerIdentity) => void,
  ): boolean {
    const message = parseEditorResourceBrokerBindMessage(value);
    if (!message || !isPort(brokerPort) || !isPort(replyPort) || brokerPort === replyPort) {
      const looksLikeBindMessage =
        isRecord(value) &&
        hasExactKeys(value, ['protocol', 'type', 'releaseId', 'sessionId']) &&
        value.protocol === RESOURCE_BROKER_PROTOCOL &&
        value.type === EDITOR_RESOURCE_BROKER_BIND_TYPE;
      const code = message ? 'ports' : looksLikeBindMessage ? 'identity' : 'protocol';
      this.#postBindError(replyPort, code);
      brokerPort.close();
      replyPort.close();
      return false;
    }

    const nextIdentity = {
      releaseId: message.releaseId,
      sessionId: message.sessionId,
    };
    const expectedIdentity = [...this.#connectionWaiters]
      .map((waiter) => waiter.identity)
      .find((identity): identity is EditorResourceBrokerIdentity => identity !== null);
    if (!validateIdentity(nextIdentity) || (expectedIdentity && !sameIdentity(expectedIdentity, nextIdentity))) {
      this.#postBindError(replyPort, 'identity');
      brokerPort.close();
      replyPort.close();
      return false;
    }

    this.#replaceConnection(
      new EditorResourceBrokerError('replaced', 'connection', 'Editor Resource Broker connection was replaced'),
    );
    this.#generation += 1;
    const generation = this.#generation;
    this.#port = brokerPort;
    this.#identity = nextIdentity;
    this.#status = 'connected';
    this.#connectionRequestInFlight = false;

    brokerPort.onmessage = (event) => {
      if (this.#generation === generation && this.#port === brokerPort) {
        this.#handleServerMessage(event.data);
      }
    };
    brokerPort.onmessageerror = () => {
      if (this.#generation === generation && this.#port === brokerPort) {
        this.#invalidateConnection(
          new EditorResourceBrokerError('connection', 'connection', 'Editor Resource Broker channel failed'),
        );
      }
    };
    const closeListener = () => {
      if (this.#generation === generation && this.#port === brokerPort) {
        this.#invalidateConnection(
          new EditorResourceBrokerError('connection', 'connection', 'Editor Resource Broker channel closed'),
        );
      }
    };
    this.#portCloseListener = closeListener;
    brokerPort.addEventListener?.('close', closeListener);
    brokerPort.start();

    try {
      beforeAcknowledge?.({ ...nextIdentity });
    } catch {
      this.#postBindError(replyPort, 'identity');
      replyPort.close();
      this.#invalidateConnection(
        new EditorResourceBrokerError('identity', 'bind', 'Editor Resource Broker identity commit failed'),
      );
      return false;
    }

    try {
      replyPort.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: EDITOR_RESOURCE_BROKER_BOUND_TYPE,
        ok: true,
        ...nextIdentity,
      } satisfies EditorResourceBrokerBoundMessage);
    } catch {
      replyPort.close();
      this.#invalidateConnection(
        new EditorResourceBrokerError('connection', 'bind', 'Editor Resource Broker binding acknowledgement failed'),
      );
      return false;
    }
    replyPort.close();
    this.#resolveConnectionWaiters(nextIdentity);
    return true;
  }

  requestConnection(identity: EditorResourceBrokerIdentity | null = null): void {
    if (identity && !validateIdentity(identity)) {
      throw new TypeError('Invalid Editor Resource Broker identity');
    }
    if (this.#status === 'connected' || this.#connectionRequestInFlight) return;
    this.#status = 'connecting';
    this.#connectionRequestInFlight = true;
    Promise.resolve()
      .then(() => this.#requestConnectionCallback?.(identity))
      .catch(() => {
        if (!this.#connectionRequestInFlight || this.#status === 'connected') return;
        this.#connectionRequestInFlight = false;
        this.#status = 'disconnected';
        this.#rejectConnectionWaiters(
          new EditorResourceBrokerError('connection', 'connection', 'Editor Resource Broker connection request failed'),
        );
      });
  }

  awaitConnection(
    timeoutMs = EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS,
    identity: EditorResourceBrokerIdentity | null = null,
  ): Promise<EditorResourceBrokerIdentity> {
    if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS) {
      return Promise.reject(new TypeError('Invalid Editor Resource Broker connection timeout'));
    }
    if (identity && !validateIdentity(identity)) {
      return Promise.reject(new EditorResourceBrokerError('identity', 'connection'));
    }
    if (this.#status === 'connected' && this.#identity) {
      return sameIdentity(this.#identity, identity ?? this.#identity)
        ? Promise.resolve({ ...this.#identity })
        : Promise.reject(new EditorResourceBrokerError('identity', 'connection'));
    }

    const promise = new Promise<EditorResourceBrokerIdentity>((resolve, reject) => {
      const waiter: ConnectionWaiter = {
        identity,
        resolve,
        reject,
        timeout: this.#setTimeout(() => {
          this.#connectionWaiters.delete(waiter);
          reject(new EditorResourceBrokerError('timeout', 'connection'));
          if (this.#connectionWaiters.size === 0 && this.#status !== 'connected') {
            this.#status = 'disconnected';
            this.#connectionRequestInFlight = false;
          }
        }, timeoutMs),
      };
      this.#connectionWaiters.add(waiter);
    });
    this.requestConnection(identity);
    return promise;
  }

  async fetchAsset(
    request: EditorResourceBrokerFetchRequest,
    pathValue: string,
    options: EditorResourceBrokerFetchOptions = {},
  ): Promise<Response> {
    const method = request.method.toUpperCase();
    if (method !== 'GET' && method !== 'HEAD') {
      throw new EditorResourceBrokerError('request', 'request', 'Only GET and HEAD are supported');
    }
    const path = normalizeResourceBrokerResourcePath(pathValue);
    if (!path || path !== pathValue) {
      throw new EditorResourceBrokerError('request', 'request', 'Invalid Office resource path');
    }
    if (request.signal?.aborted) {
      throw new EditorResourceBrokerError('cancelled', 'request');
    }
    const windowBytes = normalizeResourceBrokerReadWindowBytes(options.windowBytes);
    if (windowBytes === null) {
      throw new EditorResourceBrokerError('request', 'request', 'Invalid Resource Broker read window');
    }
    const rawRange = request.headers.get('range');
    const parsedRange = parseResourceBrokerRangeHeader(rawRange);
    const range = parsedRange === null ? ({ kind: 'suffix', bytes: 0 } as const) : parsedRange;
    const identity = await this.awaitConnection(
      options.connectionTimeoutMs ?? EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS,
      options.identity ?? null,
    );
    if (request.signal?.aborted) {
      throw new EditorResourceBrokerError('cancelled', 'request');
    }
    if (!this.#port || this.#status !== 'connected' || !this.#identity || !sameIdentity(identity, this.#identity)) {
      throw new EditorResourceBrokerError('connection', 'connection');
    }
    if (this.#pending.size >= this.#maxActiveRequests || this.#reservedBytes + windowBytes > this.#maxReservedBytes) {
      throw new EditorResourceBrokerError('busy', 'request');
    }

    const id = this.#nextRequestId();
    let resolveResponse!: (response: Response) => void;
    let rejectResponse!: (error: EditorResourceBrokerError) => void;
    const response = new Promise<Response>((resolve, reject) => {
      resolveResponse = resolve;
      rejectResponse = reject;
    });
    const state: PendingRead = {
      id,
      identity,
      generation: this.#generation,
      windowBytes,
      range,
      method,
      resolveResponse,
      rejectResponse,
      responseSettled: false,
      headers: null,
      controller: null,
      pullOutstanding: false,
      cancelRequested: false,
      receivedBytes: 0,
      finished: false,
      timeout: null,
      signal: request.signal,
      abortListener: null,
    };
    if (request.signal) {
      state.abortListener = () => {
        this.#cancelRead(state, new EditorResourceBrokerError('cancelled', 'request'));
      };
      request.signal.addEventListener('abort', state.abortListener, { once: true });
    }
    this.#pending.set(id, state);
    this.#reservedBytes += windowBytes;
    this.#peakActiveRequests = Math.max(this.#peakActiveRequests, this.#pending.size);
    this.#peakReservedBytes = Math.max(this.#peakReservedBytes, this.#reservedBytes);
    this.#refreshRequestTimeout(state);

    try {
      this.#port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id,
        releaseId: identity.releaseId,
        sessionId: identity.sessionId,
        path,
        ...(range ? { range } : {}),
        windowBytes,
      });
    } catch {
      this.#invalidateConnection(
        new EditorResourceBrokerError('connection', 'connection', 'Editor Resource Broker write failed'),
      );
    }
    return response;
  }

  disconnect(): void {
    this.#invalidateConnection(
      new EditorResourceBrokerError('connection', 'connection', 'Editor Resource Broker was disconnected'),
    );
  }

  #nextRequestId(): string {
    this.#requestCounter += 1;
    if (this.#requestCounter >= Number.MAX_SAFE_INTEGER) this.#requestCounter = 1;
    return `editor-${this.#generation}-${this.#requestCounter}`;
  }

  #handleServerMessage(value: unknown): void {
    const message = parseResourceBrokerServerMessage(value);
    if (!message) {
      this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
      return;
    }
    const state = this.#pending.get(message.id);
    if (!state || state.generation !== this.#generation) {
      this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
      return;
    }
    this.#refreshRequestTimeout(state);
    if (message.type === 'HEADERS') {
      this.#handleHeaders(state, message);
      return;
    }
    if (message.type === 'CHUNK') {
      this.#handleChunk(state, message);
      return;
    }
    if (message.type === 'END') {
      this.#handleEnd(state, message);
      return;
    }
    if (message.type === 'CANCELLED') {
      if (!state.cancelRequested || message.bytesSent !== state.receivedBytes) {
        this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
        return;
      }
      this.#finishRead(state);
      return;
    }
    if (message.type === 'ERROR') {
      this.#failRead(state, toBrokerError(message));
      return;
    }
    this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
  }

  #handleHeaders(state: PendingRead, message: ResourceBrokerHeadersMessage): void {
    if (state.headers || state.cancelRequested || !validateHeadersForRequest(message, state.range)) {
      this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
      return;
    }
    state.headers = message;
    let headers: Headers;
    try {
      headers = responseHeaders(message);
    } catch {
      this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
      return;
    }

    if (state.method === 'HEAD' || message.status === 416 || message.headers.contentLength === 0) {
      state.responseSettled = true;
      state.resolveResponse(new Response(null, { status: message.status, headers }));
      if (state.method === 'HEAD' && message.headers.contentLength > 0) {
        state.cancelRequested = true;
        this.#postFlowMessage(state, 'CANCEL');
      }
      return;
    }

    const stream = new ReadableStream<Uint8Array>(
      {
        start: (controller) => {
          state.controller = controller;
        },
        pull: () => {
          if (
            state.finished ||
            state.cancelRequested ||
            state.pullOutstanding ||
            state.receivedBytes >= message.headers.contentLength
          ) {
            return;
          }
          state.pullOutstanding = true;
          this.#postFlowMessage(state, 'PULL');
          this.#refreshRequestTimeout(state);
        },
        cancel: () => {
          this.#cancelRead(state, null);
        },
      },
      { highWaterMark: 0 },
    );
    state.responseSettled = true;
    state.resolveResponse(new Response(stream, { status: message.status, headers }));
  }

  #handleChunk(state: PendingRead, message: Extract<ResourceBrokerServerMessage, { type: 'CHUNK' }>): void {
    const headers = state.headers;
    if (
      !headers ||
      headers.status === 416 ||
      state.method === 'HEAD' ||
      state.cancelRequested ||
      !state.pullOutstanding ||
      message.bytes.byteLength > state.windowBytes ||
      state.receivedBytes + message.bytes.byteLength > headers.headers.contentLength ||
      !state.controller
    ) {
      this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
      return;
    }
    state.pullOutstanding = false;
    state.receivedBytes += message.bytes.byteLength;
    try {
      state.controller.enqueue(new Uint8Array(message.bytes));
    } catch {
      this.#cancelRead(state, null);
    }
  }

  #handleEnd(state: PendingRead, message: ResourceBrokerEndMessage): void {
    const headers = state.headers;
    if (
      !headers ||
      state.cancelRequested ||
      message.bytesSent !== state.receivedBytes ||
      state.receivedBytes !== headers.headers.contentLength
    ) {
      this.#invalidateConnection(new EditorResourceBrokerError('protocol', 'response'));
      return;
    }
    if (state.controller) {
      try {
        state.controller.close();
      } catch {
        // A consumer can cancel between the final chunk and END.
      }
    }
    this.#finishRead(state);
  }

  #postFlowMessage(state: PendingRead, type: 'PULL' | 'CANCEL'): void {
    if (state.finished || state.generation !== this.#generation || !this.#port || this.#status !== 'connected') {
      return;
    }
    try {
      this.#port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type,
        id: state.id,
      });
    } catch {
      this.#invalidateConnection(new EditorResourceBrokerError('connection', 'connection'));
    }
  }

  #cancelRead(state: PendingRead, error: EditorResourceBrokerError | null): void {
    if (state.finished || state.cancelRequested) return;
    state.cancelRequested = true;
    this.#postFlowMessage(state, 'CANCEL');
    if (error && !state.responseSettled) {
      state.responseSettled = true;
      state.rejectResponse(error);
    } else if (error && state.controller) {
      try {
        state.controller.error(error);
      } catch {
        // The stream was already cancelled.
      }
    }
    this.#refreshRequestTimeout(state);
  }

  #refreshRequestTimeout(state: PendingRead): void {
    if (state.timeout !== null) this.#clearTimeout(state.timeout);
    state.timeout = this.#setTimeout(() => {
      if (this.#pending.get(state.id) !== state) return;
      this.#invalidateConnection(new EditorResourceBrokerError('timeout', 'request'));
    }, this.#requestTimeoutMs);
  }

  #failRead(state: PendingRead, error: EditorResourceBrokerError): void {
    if (!state.responseSettled) {
      state.responseSettled = true;
      state.rejectResponse(error);
    } else if (state.controller) {
      try {
        state.controller.error(error);
      } catch {
        // The stream has already reached a terminal state.
      }
    }
    this.#finishRead(state);
  }

  #finishRead(state: PendingRead): void {
    if (state.finished || this.#pending.get(state.id) !== state) return;
    state.finished = true;
    this.#pending.delete(state.id);
    this.#reservedBytes -= state.windowBytes;
    if (state.timeout !== null) this.#clearTimeout(state.timeout);
    state.timeout = null;
    if (state.signal && state.abortListener) {
      state.signal.removeEventListener('abort', state.abortListener);
    }
    state.abortListener = null;
  }

  #replaceConnection(error: EditorResourceBrokerError): void {
    if (!this.#port) return;
    const previousPort = this.#port;
    for (const state of this.#pending.values()) {
      try {
        previousPort.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CANCEL',
          id: state.id,
        });
      } catch {
        break;
      }
    }
    this.#detachPort(previousPort);
    for (const state of this.#pending.values()) {
      this.#failRead(state, error);
    }
    this.#identity = null;
    this.#status = 'disconnected';
  }

  #invalidateConnection(error: EditorResourceBrokerError): void {
    if (this.#port) {
      const port = this.#port;
      for (const state of this.#pending.values()) {
        try {
          port.postMessage({
            protocol: RESOURCE_BROKER_PROTOCOL,
            type: 'CANCEL',
            id: state.id,
          });
        } catch {
          break;
        }
      }
      this.#detachPort(port);
    }
    for (const state of this.#pending.values()) {
      this.#failRead(state, error);
    }
    this.#identity = null;
    this.#status = 'disconnected';
    this.#connectionRequestInFlight = false;
    this.#rejectConnectionWaiters(error);
  }

  #detachPort(port: EditorResourceBrokerPort): void {
    port.onmessage = null;
    port.onmessageerror = null;
    if (this.#portCloseListener) {
      port.removeEventListener?.('close', this.#portCloseListener);
      this.#portCloseListener = null;
    }
    port.close();
    if (this.#port === port) this.#port = null;
  }

  #resolveConnectionWaiters(identity: EditorResourceBrokerIdentity): void {
    for (const waiter of this.#connectionWaiters) {
      this.#connectionWaiters.delete(waiter);
      this.#clearTimeout(waiter.timeout);
      if (!waiter.identity || sameIdentity(waiter.identity, identity)) {
        waiter.resolve({ ...identity });
      } else {
        waiter.reject(new EditorResourceBrokerError('identity', 'connection'));
      }
    }
  }

  #rejectConnectionWaiters(error: EditorResourceBrokerError): void {
    for (const waiter of this.#connectionWaiters) {
      this.#clearTimeout(waiter.timeout);
      waiter.reject(error);
    }
    this.#connectionWaiters.clear();
  }

  #postBindError(replyPort: EditorResourceBrokerPort, code: EditorResourceBrokerBindErrorMessage['code']): void {
    try {
      replyPort.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE,
        ok: false,
        code,
      } satisfies EditorResourceBrokerBindErrorMessage);
    } catch {
      // The binding peer is already gone.
    }
  }
}
