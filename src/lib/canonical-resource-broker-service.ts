import { AssetSpanReader } from './asset-span-reader';
import {
  CanonicalResourceStoreError,
  type CanonicalResourceJournal,
  type CanonicalResourceStore,
} from './canonical-resource-store';
import type { ContentObjectDescriptor } from './release-content-model';
import {
  RESOURCE_BROKER_MAX_READ_WINDOW_BYTES,
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  normalizeResourceBrokerReadWindowBytes,
  parseResourceBrokerClientMessage,
  resolveResourceBrokerRange,
  type ResourceBrokerProbeMessage,
  type ResourceBrokerReadMessage,
  type ResourceBrokerServerMessage,
} from './resource-broker-protocol';

export const CANONICAL_RESOURCE_BROKER_CLIENT_PATH = '/resource-broker.html';
export const CANONICAL_RESOURCE_BROKER_REQUEST_TIMEOUT_MS = 30_000;
export const CANONICAL_RESOURCE_BROKER_PROBE_REUSE_MS = 30_000;
export const CANONICAL_RESOURCE_BROKER_MAX_ACTIVE_READS = 64;
export const CANONICAL_RESOURCE_BROKER_MAX_RESERVED_BYTES = 64 * 1024 * 1024;
export const CANONICAL_RESOURCE_BROKER_READ_RESERVATION_FACTOR = 2;
export const CANONICAL_RESOURCE_BROKER_MAX_QUEUED_READS = 512;

const ALLOWED_PARENT_ORIGINS = new Set(['https://piwork.getpi.work', 'https://onlyoffice.getpi.work']);
const SERVABLE_RELEASE_STATES = new Set(['active', 'retained']);

export interface CanonicalResourceBrokerClientIdentity {
  releaseId: string;
  sessionId: string;
  parentOrigin: string;
  localMatrix: boolean;
}

export interface CanonicalResourceBrokerPort {
  onmessage: ((event: MessageEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  close(): void;
  postMessage(message: unknown, transfer?: Transferable[]): void;
  start(): void;
}

export interface CanonicalResourceBrokerServiceOptions {
  journal: Pick<CanonicalResourceJournal, 'getAssetReadView'>;
  store: Pick<CanonicalResourceStore, 'matchObject' | 'probeRelease'>;
  requestTimeoutMs?: number;
  maxActiveReads?: number;
  maxReservedBytes?: number;
  setTimeout?: typeof globalThis.setTimeout;
  clearTimeout?: typeof globalThis.clearTimeout;
  now?: () => number;
  probeReuseMs?: number;
}

export interface CanonicalResourceBrokerMetrics {
  schemaVersion: 1;
  role: 'canonical-service';
  activeReads: number;
  peakActiveReads: number;
  reservedBytes: number;
  peakReservedBytes: number;
  maxActiveReads: number;
  maxReservedBytes: number;
}

type BrokerErrorCode = Extract<ResourceBrokerServerMessage, { type: 'ERROR' }>['code'];
type QueuedRead = {
  request: ResourceBrokerReadMessage;
  port: CanonicalResourceBrokerPort;
  reservedBytes: number;
  enqueuedAt: number;
  timeout: ReturnType<typeof globalThis.setTimeout>;
  resolve(remainingTimeoutMs: number | null): void;
};

function hasExactSearchParams(params: URLSearchParams, allowed: ReadonlySet<string>): boolean {
  for (const key of params.keys()) {
    if (!allowed.has(key) || params.getAll(key).length !== 1) return false;
  }
  return true;
}

export function parseCanonicalResourceBrokerClientUrl(
  value: string,
  expectedOrigin: string,
  allowLocalMatrix = false,
): CanonicalResourceBrokerClientIdentity | null {
  let url: URL;
  let canonicalOrigin: URL;
  try {
    url = new URL(value);
    canonicalOrigin = new URL(expectedOrigin);
  } catch {
    return null;
  }
  if (
    url.origin !== canonicalOrigin.origin ||
    expectedOrigin !== canonicalOrigin.origin ||
    url.pathname !== CANONICAL_RESOURCE_BROKER_CLIENT_PATH ||
    url.hash
  ) {
    return null;
  }
  const allowed = new Set(['releaseId', 'sessionId', 'parentOrigin', 'localMatrix']);
  if (!hasExactSearchParams(url.searchParams, allowed)) return null;
  const releaseId = url.searchParams.get('releaseId');
  const sessionId = url.searchParams.get('sessionId');
  const parentOrigin = url.searchParams.get('parentOrigin');
  const localMatrix = url.searchParams.get('localMatrix');
  if (
    !releaseId ||
    !sessionId ||
    !parentOrigin ||
    !isResourceBrokerReleaseId(releaseId) ||
    !isResourceBrokerSessionId(sessionId) ||
    !ALLOWED_PARENT_ORIGINS.has(parentOrigin) ||
    (localMatrix !== null && localMatrix !== '1') ||
    (localMatrix === '1' && !allowLocalMatrix)
  ) {
    return null;
  }
  return {
    releaseId,
    sessionId,
    parentOrigin,
    localMatrix: localMatrix === '1',
  };
}

function emptyStream(): ReadableStream<Uint8Array> {
  return new ReadableStream<Uint8Array>({
    type: 'bytes',
    start(controller) {
      controller.close();
    },
  });
}

function brokerErrorCode(error: unknown): BrokerErrorCode {
  if (error instanceof CanonicalResourceStoreError) {
    if (error.code === 'integrity' || error.code === 'invalid-content') return 'integrity';
    if (error.code === 'incomplete') return 'missing';
    return 'storage';
  }
  if (error instanceof DOMException && error.name === 'AbortError') return 'cancelled';
  return 'storage';
}

export class CanonicalResourceBrokerService {
  readonly #journal: CanonicalResourceBrokerServiceOptions['journal'];
  readonly #store: CanonicalResourceBrokerServiceOptions['store'];
  readonly #requestTimeoutMs: number;
  readonly #maxActiveReads: number;
  readonly #maxReservedBytes: number;
  readonly #probeInFlight = new Map<string, Promise<Awaited<ReturnType<CanonicalResourceStore['probeRelease']>>>>();
  readonly #recentProbes = new Map<
    string,
    { expiresAt: number; value: Awaited<ReturnType<CanonicalResourceStore['probeRelease']>> }
  >();
  readonly #now: () => number;
  readonly #probeReuseMs: number;
  readonly #setTimeout: typeof globalThis.setTimeout;
  readonly #clearTimeout: typeof globalThis.clearTimeout;
  readonly #queuedReads: QueuedRead[] = [];
  #activeReads = 0;
  #reservedBytes = 0;
  #peakActiveReads = 0;
  #peakReservedBytes = 0;

  constructor(options: CanonicalResourceBrokerServiceOptions) {
    this.#journal = options.journal;
    this.#store = options.store;
    this.#requestTimeoutMs = options.requestTimeoutMs ?? CANONICAL_RESOURCE_BROKER_REQUEST_TIMEOUT_MS;
    this.#maxActiveReads = options.maxActiveReads ?? CANONICAL_RESOURCE_BROKER_MAX_ACTIVE_READS;
    this.#maxReservedBytes = options.maxReservedBytes ?? CANONICAL_RESOURCE_BROKER_MAX_RESERVED_BYTES;
    this.#now = options.now ?? Date.now;
    this.#probeReuseMs = options.probeReuseMs ?? CANONICAL_RESOURCE_BROKER_PROBE_REUSE_MS;
    if (
      !Number.isSafeInteger(this.#requestTimeoutMs) ||
      this.#requestTimeoutMs <= 0 ||
      this.#requestTimeoutMs > CANONICAL_RESOURCE_BROKER_REQUEST_TIMEOUT_MS ||
      !Number.isSafeInteger(this.#maxActiveReads) ||
      this.#maxActiveReads <= 0 ||
      !Number.isSafeInteger(this.#maxReservedBytes) ||
      this.#maxReservedBytes < RESOURCE_BROKER_MAX_READ_WINDOW_BYTES ||
      !Number.isSafeInteger(this.#probeReuseMs) ||
      this.#probeReuseMs < 0 ||
      this.#probeReuseMs > CANONICAL_RESOURCE_BROKER_PROBE_REUSE_MS
    ) {
      throw new TypeError('Invalid canonical Resource Broker limits');
    }
    this.#setTimeout = options.setTimeout ?? globalThis.setTimeout.bind(globalThis);
    this.#clearTimeout = options.clearTimeout ?? globalThis.clearTimeout.bind(globalThis);
  }

  get activeReadCount(): number {
    return this.#activeReads;
  }

  get reservedReadBytes(): number {
    return this.#reservedBytes;
  }

  get metrics(): CanonicalResourceBrokerMetrics {
    return {
      schemaVersion: 1,
      role: 'canonical-service',
      activeReads: this.#activeReads,
      peakActiveReads: this.#peakActiveReads,
      reservedBytes: this.#reservedBytes,
      peakReservedBytes: this.#peakReservedBytes,
      maxActiveReads: this.#maxActiveReads,
      maxReservedBytes: this.#maxReservedBytes,
    };
  }

  async handle(requestValue: unknown, port: CanonicalResourceBrokerPort): Promise<void> {
    const request = parseResourceBrokerClientMessage(requestValue);
    if (!request || (request.type !== 'PROBE' && request.type !== 'READ')) {
      this.#postError(
        port,
        typeof requestValue === 'object' && requestValue ? String((requestValue as { id?: unknown }).id ?? '') : '',
        'protocol',
      );
      port.close();
      return;
    }
    port.start();
    if (request.type === 'PROBE') {
      await this.#handleProbe(request, port);
      return;
    }
    await this.#handleRead(request, port);
  }

  async #handleProbe(request: ResourceBrokerProbeMessage, port: CanonicalResourceBrokerPort): Promise<void> {
    try {
      const now = this.#now();
      const recent = this.#recentProbes.get(request.releaseId);
      if (recent && recent.expiresAt <= now) this.#recentProbes.delete(request.releaseId);
      let probePromise =
        recent && recent.expiresAt > now ? Promise.resolve(recent.value) : this.#probeInFlight.get(request.releaseId);
      if (!probePromise) {
        probePromise = this.#store.probeRelease(request.releaseId);
        this.#probeInFlight.set(request.releaseId, probePromise);
        const clearProbe = () => {
          if (this.#probeInFlight.get(request.releaseId) === probePromise) {
            this.#probeInFlight.delete(request.releaseId);
          }
        };
        void probePromise.then(clearProbe, clearProbe);
      }
      const probe = await probePromise;
      if (!SERVABLE_RELEASE_STATES.has(probe.state)) {
        this.#postError(port, request.id, 'release');
        return;
      }
      if (!probe.ready || !probe.probeSucceeded || probe.releaseId !== request.releaseId || !probe.probePath) {
        this.#postError(port, request.id, 'missing');
        return;
      }
      this.#recentProbes.set(request.releaseId, {
        expiresAt: this.#now() + this.#probeReuseMs,
        value: probe,
      });
      port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'RESULT',
        id: request.id,
        value: {
          releaseId: request.releaseId,
          ready: true,
          probePath: probe.probePath,
          probeBytes: probe.probeAssetBytes!,
          probeSha256: probe.probeAssetSha256!,
        },
      } satisfies ResourceBrokerServerMessage);
    } catch (error) {
      this.#postError(port, request.id, brokerErrorCode(error));
    } finally {
      port.close();
    }
  }

  async #handleRead(request: ResourceBrokerReadMessage, port: CanonicalResourceBrokerPort): Promise<void> {
    const windowBytes = normalizeResourceBrokerReadWindowBytes(request.windowBytes);
    // A read can briefly own both its BYOB source buffer and the exact-sized
    // transferable chunk. Charging 2× the negotiated window makes the 64 MiB
    // budget a bound on actual in-flight canonical Broker buffers, not merely
    // a count of protocol payloads.
    const reservedBytes = windowBytes === null ? 0 : windowBytes * CANONICAL_RESOURCE_BROKER_READ_RESERVATION_FACTOR;
    if (windowBytes === null) {
      this.#postError(port, request.id, 'busy');
      port.close();
      return;
    }
    const admission = this.#acquireReadBudget(request, port, reservedBytes);
    const initialTimeoutMs = typeof admission === 'number' || admission === null ? admission : await admission;
    if (initialTimeoutMs === null) return;

    let reader: AssetSpanReader | null = null;
    let finished = false;
    let cancelled = false;
    let bytesSent = 0;
    let timeout: ReturnType<typeof globalThis.setTimeout> | null = null;
    let flowQueue = Promise.resolve();
    let initialized = false;
    let pullOutstanding = false;
    let contentLength = 0;
    const pendingPulls: Array<{ type: 'PULL'; id: string }> = [];
    let resolveLifetime!: () => void;
    const lifetime = new Promise<void>((resolve) => {
      resolveLifetime = resolve;
    });

    const releaseBudget = () => this.#releaseReadBudget(reservedBytes);
    const refreshTimeout = (delayMs = this.#requestTimeoutMs) => {
      if (timeout !== null) this.#clearTimeout(timeout);
      timeout = this.#setTimeout(() => {
        void fail('timeout');
      }, delayMs);
    };
    const finish = async () => {
      if (finished) return;
      finished = true;
      if (timeout !== null) this.#clearTimeout(timeout);
      timeout = null;
      port.onmessage = null;
      port.onmessageerror = null;
      await reader?.cancel().catch(() => undefined);
      port.close();
      releaseBudget();
      resolveLifetime();
    };
    const fail = async (code: BrokerErrorCode) => {
      if (finished) return;
      this.#postError(port, request.id, code);
      await finish();
    };
    const cancel = async () => {
      if (finished) return;
      cancelled = true;
      await reader?.cancel(new DOMException('The request was cancelled', 'AbortError')).catch(() => undefined);
      try {
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CANCELLED',
          id: request.id,
          bytesSent,
        } satisfies ResourceBrokerServerMessage);
      } finally {
        await finish();
      }
    };
    const pull = async () => {
      if (finished || cancelled || !reader) return;
      refreshTimeout();
      const chunk = await reader.read(windowBytes);
      if (!chunk) {
        if (bytesSent !== contentLength) {
          await fail('integrity');
          return;
        }
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'END',
          id: request.id,
          bytesSent,
        } satisfies ResourceBrokerServerMessage);
        await finish();
        return;
      }
      bytesSent += chunk.byteLength;
      if (bytesSent > contentLength) {
        await fail('integrity');
        return;
      }
      const transferable =
        chunk.byteOffset === 0 && chunk.buffer instanceof ArrayBuffer && chunk.buffer.byteLength === chunk.byteLength
          ? chunk.buffer
          : chunk.slice().buffer;
      port.postMessage(
        {
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CHUNK',
          id: request.id,
          bytes: transferable,
        } satisfies ResourceBrokerServerMessage,
        [transferable],
      );
      if (bytesSent === contentLength) {
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'END',
          id: request.id,
          bytesSent,
        } satisfies ResourceBrokerServerMessage);
        await finish();
      }
    };
    const queuePull = () => {
      if (pullOutstanding) {
        void fail('protocol');
        return;
      }
      pullOutstanding = true;
      flowQueue = flowQueue
        .then(pull)
        .catch((error) => fail(brokerErrorCode(error)))
        .finally(() => {
          pullOutstanding = false;
        });
    };

    port.onmessage = (event) => {
      const message = parseResourceBrokerClientMessage(event.data);
      if (!message || (message.type !== 'PULL' && message.type !== 'CANCEL') || message.id !== request.id) {
        void fail('protocol');
        return;
      }
      if (message.type === 'CANCEL') {
        void cancel();
        return;
      }
      if (!initialized) {
        if (pendingPulls.length > 0) {
          void fail('protocol');
          return;
        }
        pendingPulls.push(message);
        return;
      }
      queuePull();
    };
    port.onmessageerror = () => {
      void fail('protocol');
    };

    refreshTimeout(initialTimeoutMs);
    try {
      const view = await this.#journal.getAssetReadView(request.releaseId, request.path);
      const release = view.release;
      if (finished) return;
      if (!release || !SERVABLE_RELEASE_STATES.has(release.state)) {
        await fail('release');
        return;
      }
      const mapping = view.mapping;
      if (finished) return;
      if (!mapping || mapping.releaseId !== request.releaseId) {
        await fail('missing');
        return;
      }
      const resolvedRange = resolveResourceBrokerRange(request.range, mapping.assetBytes);
      contentLength = resolvedRange.contentLength;
      const contentType = mapping.mime ?? 'application/octet-stream';
      port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'HEADERS',
        id: request.id,
        status: resolvedRange.status,
        headers: {
          acceptRanges: 'bytes',
          contentLength: resolvedRange.contentLength,
          contentRange: resolvedRange.contentRange,
          contentType,
        },
      } satisfies ResourceBrokerServerMessage);
      if (resolvedRange.status === 416 || resolvedRange.contentLength === 0) {
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'END',
          id: request.id,
          bytesSent: 0,
        } satisfies ResourceBrokerServerMessage);
        await finish();
        return;
      }

      const records = new Map(view.objects.map((record) => [record.sha256, record]));
      reader = new AssetSpanReader(mapping, { start: resolvedRange.start, end: resolvedRange.end }, async (digest) => {
        const descriptor: ContentObjectDescriptor | undefined = records.get(digest);
        if (!descriptor)
          throw new CanonicalResourceStoreError('invalid-content', 'mapping references an unknown object');
        const response = await this.#store.matchObject(descriptor, records.get(digest));
        if (!response) throw new CanonicalResourceStoreError('incomplete', `content object ${digest} is missing`);
        return {
          bytes: descriptor.bytes,
          stream: response.body ?? emptyStream(),
        };
      });

      initialized = true;
      refreshTimeout();
      if (pendingPulls.length > 0) queuePull();
      await lifetime;
    } catch (error) {
      await fail(brokerErrorCode(error));
    }
  }

  #hasReadBudget(reservedBytes: number): boolean {
    return this.#activeReads < this.#maxActiveReads && this.#reservedBytes + reservedBytes <= this.#maxReservedBytes;
  }

  #reserveReadBudget(reservedBytes: number): void {
    this.#activeReads += 1;
    this.#reservedBytes += reservedBytes;
    this.#peakActiveReads = Math.max(this.#peakActiveReads, this.#activeReads);
    this.#peakReservedBytes = Math.max(this.#peakReservedBytes, this.#reservedBytes);
  }

  #acquireReadBudget(
    request: ResourceBrokerReadMessage,
    port: CanonicalResourceBrokerPort,
    reservedBytes: number,
  ): number | null | Promise<number | null> {
    if (this.#hasReadBudget(reservedBytes)) {
      this.#reserveReadBudget(reservedBytes);
      return this.#requestTimeoutMs;
    }
    if (this.#queuedReads.length >= CANONICAL_RESOURCE_BROKER_MAX_QUEUED_READS) {
      this.#postError(port, request.id, 'busy');
      port.close();
      return null;
    }

    return new Promise((resolve) => {
      const queued: QueuedRead = {
        request,
        port,
        reservedBytes,
        enqueuedAt: this.#now(),
        timeout: this.#setTimeout(() => {
          if (!this.#removeQueuedRead(queued)) return;
          this.#postError(port, request.id, 'timeout');
          port.close();
          resolve(null);
        }, this.#requestTimeoutMs),
        resolve,
      };
      port.onmessage = (event) => {
        const message = parseResourceBrokerClientMessage(event.data);
        if (!message || message.type !== 'CANCEL' || message.id !== request.id) {
          if (!this.#removeQueuedRead(queued)) return;
          this.#clearTimeout(queued.timeout);
          this.#postError(port, request.id, 'protocol');
          port.close();
          resolve(null);
          return;
        }
        if (!this.#removeQueuedRead(queued)) return;
        this.#clearTimeout(queued.timeout);
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CANCELLED',
          id: request.id,
          bytesSent: 0,
        } satisfies ResourceBrokerServerMessage);
        port.close();
        resolve(null);
      };
      port.onmessageerror = () => {
        if (!this.#removeQueuedRead(queued)) return;
        this.#clearTimeout(queued.timeout);
        this.#postError(port, request.id, 'protocol');
        port.close();
        resolve(null);
      };
      this.#queuedReads.push(queued);
    });
  }

  #removeQueuedRead(queued: QueuedRead): boolean {
    const index = this.#queuedReads.indexOf(queued);
    if (index < 0) return false;
    this.#queuedReads.splice(index, 1);
    return true;
  }

  #releaseReadBudget(reservedBytes: number): void {
    this.#activeReads -= 1;
    this.#reservedBytes -= reservedBytes;
    this.#drainReadQueue();
  }

  #drainReadQueue(): void {
    while (this.#queuedReads.length > 0) {
      const queued = this.#queuedReads[0];
      if (!this.#hasReadBudget(queued.reservedBytes)) return;
      this.#queuedReads.shift();
      this.#clearTimeout(queued.timeout);
      queued.port.onmessage = null;
      queued.port.onmessageerror = null;
      this.#reserveReadBudget(queued.reservedBytes);
      const elapsed = Math.max(0, this.#now() - queued.enqueuedAt);
      queued.resolve(Math.max(1, this.#requestTimeoutMs - elapsed));
    }
  }

  #postError(port: CanonicalResourceBrokerPort, id: string, code: BrokerErrorCode): void {
    if (!id) return;
    try {
      port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ERROR',
        id,
        code,
      } satisfies ResourceBrokerServerMessage);
    } catch {
      // The peer is already gone; local cleanup still completes.
    }
  }
}
