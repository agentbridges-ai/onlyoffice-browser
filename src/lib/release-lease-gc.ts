import {
  isResourceBrokerEditorOrigin,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  type ResourceBrokerEditorOrigin,
} from './resource-broker-protocol';
import type { ReleaseTransactionState } from './release-content-model';

export const RELEASE_LEASE_DEFAULT_TTL_MS = 45_000;
export const RELEASE_LEASE_MAX_TTL_MS = 120_000;
export const RELEASE_LEASE_HEARTBEAT_INTERVAL_MS = 15_000;
export const RELEASE_GC_MINIMUM_RETENTION_MS = 90 * 24 * 60 * 60 * 1_000;
export const RELEASE_GC_MINIMUM_RECENT_RELEASES = 3;
export const RELEASE_GC_EXECUTION_POLICY = 'no-delete' as const;

const DIGEST_PATTERN = /^[a-f0-9]{64}$/;
const LEASE_ID_PATTERN = /^[a-zA-Z0-9._-]{16,128}$/;
const PROTECTED_TRANSACTION_STATES = new Set<ReleaseTransactionState>(['installing', 'prepared', 'active', 'retained']);
const RELEASE_LEASE_DATABASE_NAME = 'onlyoffice-release-leases-v1';
const RELEASE_LEASE_DATABASE_VERSION = 1;
const RELEASE_LEASE_STORE_NAME = 'leases';

export type ReleaseLeaseErrorCode =
  | 'invalid-binding'
  | 'invalid-snapshot'
  | 'duplicate'
  | 'unknown'
  | 'expired'
  | 'mismatch'
  | 'clock';

export class ReleaseLeaseError extends Error {
  constructor(
    readonly code: ReleaseLeaseErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'ReleaseLeaseError';
  }
}

export interface ReleaseLeaseBinding {
  leaseId: string;
  releaseId: string;
  sessionId: string;
  editorOrigin: ResourceBrokerEditorOrigin;
}

export interface ReleaseLeaseRecord extends ReleaseLeaseBinding {
  acquiredAtMs: number;
  heartbeatAtMs: number;
  expiresAtMs: number;
}

export interface ReleaseLeaseLedgerSnapshot {
  version: 1;
  leases: ReleaseLeaseRecord[];
}

export interface ReleaseLeaseStore {
  acquireLease(binding: ReleaseLeaseBinding, ttlMs?: number): ReleaseLeaseRecord | Promise<ReleaseLeaseRecord>;
  acquireOrRenewLease(binding: ReleaseLeaseBinding, ttlMs?: number): ReleaseLeaseRecord | Promise<ReleaseLeaseRecord>;
  heartbeatLease(binding: ReleaseLeaseBinding, ttlMs?: number): ReleaseLeaseRecord | Promise<ReleaseLeaseRecord>;
  releaseLease(binding: ReleaseLeaseBinding): boolean | Promise<boolean>;
  reapExpired(nowMs?: number): string[] | Promise<string[]>;
  listActiveLeases(nowMs?: number): ReleaseLeaseRecord[] | Promise<ReleaseLeaseRecord[]>;
  snapshot(): ReleaseLeaseLedgerSnapshot | Promise<ReleaseLeaseLedgerSnapshot>;
}

export interface ReleaseGcReleaseRecord {
  releaseId: string;
  state: ReleaseTransactionState;
  publishedAtMs: number;
  objectSha256: string[];
}

export type ReleaseGcLedgerState = 'valid' | 'unknown' | 'corrupt';

export interface ReleaseGcInput {
  nowMs: number;
  ledgerState: ReleaseGcLedgerState;
  releases: ReleaseGcReleaseRecord[];
  knownObjectSha256: string[];
  leaseLedger: ReleaseLeaseLedgerSnapshot;
  piworkDescriptorReleaseIds: string[];
  stableReleaseIds: string[];
}

export type ReleaseGcRetentionReason =
  | 'transaction-installing'
  | 'transaction-prepared'
  | 'transaction-active'
  | 'transaction-retained'
  | 'live-editor-lease'
  | 'piwork-descriptor'
  | 'stable-pointer'
  | 'recent-release'
  | 'minimum-retention-window';

export type ReleaseGcBlockedReason = 'ledger-unknown' | 'ledger-corrupt' | 'execution-disabled';

export interface ReleaseGcRetentionMark {
  releaseId: string;
  reasons: ReleaseGcRetentionReason[];
}

export interface ReleaseGcPlan {
  blocked: boolean;
  blockedReason: ReleaseGcBlockedReason | null;
  retainedReleaseIds: string[];
  retentionMarks: ReleaseGcRetentionMark[];
  deleteReleaseIds: string[];
  deleteObjectSha256: string[];
}

function leaseFail(code: ReleaseLeaseErrorCode, message: string): never {
  throw new ReleaseLeaseError(code, message);
}

function isSafeTimestamp(value: unknown): value is number {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0;
}

function currentTime(now: () => number): number {
  const value = now();
  if (!isSafeTimestamp(value)) leaseFail('clock', 'lease clock returned an invalid timestamp');
  return value;
}

function validateTtl(ttlMs: number): void {
  if (!Number.isSafeInteger(ttlMs) || ttlMs <= 0 || ttlMs > RELEASE_LEASE_MAX_TTL_MS) {
    leaseFail('invalid-binding', `lease TTL must be between 1 and ${RELEASE_LEASE_MAX_TTL_MS} ms`);
  }
}

function validateBinding(binding: ReleaseLeaseBinding): void {
  if (!LEASE_ID_PATTERN.test(binding.leaseId)) leaseFail('invalid-binding', 'invalid lease id');
  if (!isResourceBrokerReleaseId(binding.releaseId)) leaseFail('invalid-binding', 'invalid release id');
  if (!isResourceBrokerSessionId(binding.sessionId)) leaseFail('invalid-binding', 'invalid session id');
  if (!isResourceBrokerEditorOrigin(binding.editorOrigin)) {
    leaseFail('invalid-binding', 'invalid editor origin');
  }
}

function cloneLease(lease: ReleaseLeaseRecord): ReleaseLeaseRecord {
  return { ...lease };
}

function validateLeaseRecord(record: ReleaseLeaseRecord): void {
  validateBinding(record);
  if (
    !isSafeTimestamp(record.acquiredAtMs) ||
    !isSafeTimestamp(record.heartbeatAtMs) ||
    !isSafeTimestamp(record.expiresAtMs) ||
    record.heartbeatAtMs < record.acquiredAtMs ||
    record.expiresAtMs <= record.heartbeatAtMs ||
    record.expiresAtMs - record.heartbeatAtMs > RELEASE_LEASE_MAX_TTL_MS
  ) {
    leaseFail('invalid-snapshot', `invalid timestamps for lease ${record.leaseId}`);
  }
}

function sameBinding(left: ReleaseLeaseBinding, right: ReleaseLeaseBinding): boolean {
  return (
    left.leaseId === right.leaseId &&
    left.releaseId === right.releaseId &&
    left.sessionId === right.sessionId &&
    left.editorOrigin === right.editorOrigin
  );
}

export class MemoryReleaseLeaseLedger {
  readonly #leases = new Map<string, ReleaseLeaseRecord>();
  readonly #now: () => number;
  readonly #defaultTtlMs: number;

  constructor(
    options: {
      snapshot?: ReleaseLeaseLedgerSnapshot;
      now?: () => number;
      defaultTtlMs?: number;
    } = {},
  ) {
    this.#now = options.now || Date.now;
    this.#defaultTtlMs = options.defaultTtlMs ?? RELEASE_LEASE_DEFAULT_TTL_MS;
    validateTtl(this.#defaultTtlMs);

    const snapshot = options.snapshot ?? { version: 1 as const, leases: [] };
    if (snapshot.version !== 1 || !Array.isArray(snapshot.leases)) {
      leaseFail('invalid-snapshot', 'unsupported release lease ledger snapshot');
    }
    for (const record of snapshot.leases) {
      validateLeaseRecord(record);
      if (this.#leases.has(record.leaseId)) {
        leaseFail('invalid-snapshot', `duplicate lease ${record.leaseId}`);
      }
      this.#leases.set(record.leaseId, cloneLease(record));
    }
  }

  acquireLease(binding: ReleaseLeaseBinding, ttlMs: number = this.#defaultTtlMs): ReleaseLeaseRecord {
    validateBinding(binding);
    validateTtl(ttlMs);
    if (this.#leases.has(binding.leaseId)) leaseFail('duplicate', `lease ${binding.leaseId} already exists`);
    const nowMs = currentTime(this.#now);
    if (!Number.isSafeInteger(nowMs + ttlMs)) leaseFail('clock', 'lease expiration timestamp overflow');
    const record: ReleaseLeaseRecord = {
      ...binding,
      acquiredAtMs: nowMs,
      heartbeatAtMs: nowMs,
      expiresAtMs: nowMs + ttlMs,
    };
    this.#leases.set(record.leaseId, record);
    return cloneLease(record);
  }

  acquireOrRenewLease(binding: ReleaseLeaseBinding, ttlMs: number = this.#defaultTtlMs): ReleaseLeaseRecord {
    validateBinding(binding);
    validateTtl(ttlMs);
    const record = this.#leases.get(binding.leaseId);
    if (!record) return this.acquireLease(binding, ttlMs);
    if (!sameBinding(record, binding)) leaseFail('mismatch', `lease ${binding.leaseId} binding mismatch`);
    const nowMs = currentTime(this.#now);
    if (nowMs < record.heartbeatAtMs) leaseFail('clock', 'lease clock moved backwards');
    if (!Number.isSafeInteger(nowMs + ttlMs)) leaseFail('clock', 'lease expiration timestamp overflow');
    if (nowMs >= record.expiresAtMs) record.acquiredAtMs = nowMs;
    record.heartbeatAtMs = nowMs;
    record.expiresAtMs = nowMs + ttlMs;
    return cloneLease(record);
  }

  heartbeatLease(binding: ReleaseLeaseBinding, ttlMs: number = this.#defaultTtlMs): ReleaseLeaseRecord {
    validateBinding(binding);
    validateTtl(ttlMs);
    const record = this.#leases.get(binding.leaseId);
    if (!record) leaseFail('unknown', `lease ${binding.leaseId} does not exist`);
    if (!sameBinding(record, binding)) leaseFail('mismatch', `lease ${binding.leaseId} binding mismatch`);
    const nowMs = currentTime(this.#now);
    if (nowMs < record.heartbeatAtMs) leaseFail('clock', 'lease clock moved backwards');
    if (nowMs >= record.expiresAtMs) leaseFail('expired', `lease ${binding.leaseId} expired`);
    if (!Number.isSafeInteger(nowMs + ttlMs)) leaseFail('clock', 'lease expiration timestamp overflow');
    record.heartbeatAtMs = nowMs;
    record.expiresAtMs = nowMs + ttlMs;
    return cloneLease(record);
  }

  releaseLease(binding: ReleaseLeaseBinding): boolean {
    validateBinding(binding);
    const record = this.#leases.get(binding.leaseId);
    if (!record) return false;
    if (!sameBinding(record, binding)) leaseFail('mismatch', `lease ${binding.leaseId} binding mismatch`);
    this.#leases.delete(binding.leaseId);
    return true;
  }

  reapExpired(nowMs: number = currentTime(this.#now)): string[] {
    if (!isSafeTimestamp(nowMs)) leaseFail('clock', 'invalid lease reap timestamp');
    const expired = [...this.#leases.values()]
      .filter((lease) => lease.expiresAtMs <= nowMs)
      .map((lease) => lease.leaseId)
      .sort();
    for (const leaseId of expired) this.#leases.delete(leaseId);
    return expired;
  }

  listActiveLeases(nowMs: number = currentTime(this.#now)): ReleaseLeaseRecord[] {
    if (!isSafeTimestamp(nowMs)) leaseFail('clock', 'invalid lease query timestamp');
    return [...this.#leases.values()]
      .filter((lease) => lease.expiresAtMs > nowMs)
      .map(cloneLease)
      .sort((left, right) => left.leaseId.localeCompare(right.leaseId));
  }

  snapshot(): ReleaseLeaseLedgerSnapshot {
    return {
      version: 1,
      leases: [...this.#leases.values()]
        .map(cloneLease)
        .sort((left, right) => left.leaseId.localeCompare(right.leaseId)),
    };
  }
}

function requestResult<T>(request: IDBRequest<T>, label: string): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error(`${label} failed`));
  });
}

function transactionCompletion(transaction: IDBTransaction, failure: () => unknown): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(failure() || transaction.error || new Error('release lease transaction failed'));
    transaction.onabort = () =>
      reject(failure() || transaction.error || new Error('release lease transaction aborted'));
  });
}

/**
 * A canonical-origin lease ledger. IndexedDB serializes readwrite transactions
 * for this object store across tabs, so acquire/renew/release cannot lose a
 * concurrent writer even when Web Locks are unavailable or a tab crashes.
 */
export class IndexedDbReleaseLeaseLedger implements ReleaseLeaseStore {
  readonly #databasePromise: Promise<IDBDatabase>;
  readonly #now: () => number;
  readonly #defaultTtlMs: number;

  constructor(
    indexedDb: IDBFactory = indexedDB,
    options: {
      databaseName?: string;
      now?: () => number;
      defaultTtlMs?: number;
    } = {},
  ) {
    this.#now = options.now || Date.now;
    this.#defaultTtlMs = options.defaultTtlMs ?? RELEASE_LEASE_DEFAULT_TTL_MS;
    validateTtl(this.#defaultTtlMs);
    this.#databasePromise = new Promise((resolve, reject) => {
      const request = indexedDb.open(
        options.databaseName || RELEASE_LEASE_DATABASE_NAME,
        RELEASE_LEASE_DATABASE_VERSION,
      );
      request.onupgradeneeded = () => {
        if (!request.result.objectStoreNames.contains(RELEASE_LEASE_STORE_NAME)) {
          request.result.createObjectStore(RELEASE_LEASE_STORE_NAME, { keyPath: 'leaseId' });
        }
      };
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onblocked = () => reject(new Error('release lease database open was blocked'));
      request.onerror = () => reject(request.error || new Error('release lease database open failed'));
    });
  }

  async acquireLease(binding: ReleaseLeaseBinding, ttlMs: number = this.#defaultTtlMs): Promise<ReleaseLeaseRecord> {
    return this.#upsertLease(binding, ttlMs, 'acquire');
  }

  async acquireOrRenewLease(
    binding: ReleaseLeaseBinding,
    ttlMs: number = this.#defaultTtlMs,
  ): Promise<ReleaseLeaseRecord> {
    return this.#upsertLease(binding, ttlMs, 'recover');
  }

  async heartbeatLease(binding: ReleaseLeaseBinding, ttlMs: number = this.#defaultTtlMs): Promise<ReleaseLeaseRecord> {
    return this.#upsertLease(binding, ttlMs, 'heartbeat');
  }

  async #upsertLease(
    binding: ReleaseLeaseBinding,
    ttlMs: number,
    mode: 'acquire' | 'recover' | 'heartbeat',
  ): Promise<ReleaseLeaseRecord> {
    validateBinding(binding);
    validateTtl(ttlMs);
    const nowMs = currentTime(this.#now);
    if (!Number.isSafeInteger(nowMs + ttlMs)) leaseFail('clock', 'lease expiration timestamp overflow');
    const database = await this.#databasePromise;
    const transaction = database.transaction(RELEASE_LEASE_STORE_NAME, 'readwrite');
    let result: ReleaseLeaseRecord | null = null;
    let operationFailure: unknown;
    const completion = transactionCompletion(transaction, () => operationFailure);
    const store = transaction.objectStore(RELEASE_LEASE_STORE_NAME);
    const request = store.get(binding.leaseId) as IDBRequest<ReleaseLeaseRecord | undefined>;
    request.onsuccess = () => {
      try {
        const existing = request.result;
        if (existing) {
          validateLeaseRecord(existing);
          if (!sameBinding(existing, binding)) {
            leaseFail('mismatch', `lease ${binding.leaseId} binding mismatch`);
          }
          if (nowMs < existing.heartbeatAtMs) leaseFail('clock', 'lease clock moved backwards');
          if (mode === 'acquire') leaseFail('duplicate', `lease ${binding.leaseId} already exists`);
          if (mode === 'heartbeat' && nowMs >= existing.expiresAtMs) {
            leaseFail('expired', `lease ${binding.leaseId} expired`);
          }
          result = {
            ...existing,
            acquiredAtMs: nowMs >= existing.expiresAtMs ? nowMs : existing.acquiredAtMs,
            heartbeatAtMs: nowMs,
            expiresAtMs: nowMs + ttlMs,
          };
        } else {
          if (mode === 'heartbeat') leaseFail('unknown', `lease ${binding.leaseId} does not exist`);
          result = {
            ...binding,
            acquiredAtMs: nowMs,
            heartbeatAtMs: nowMs,
            expiresAtMs: nowMs + ttlMs,
          };
        }
        store.put(result);
      } catch (error) {
        operationFailure = error;
        transaction.abort();
      }
    };
    request.onerror = () => {
      operationFailure = request.error || new Error('release lease read failed');
    };
    await completion;
    if (!result) throw new Error('release lease write completed without a result');
    return cloneLease(result);
  }

  async releaseLease(binding: ReleaseLeaseBinding): Promise<boolean> {
    validateBinding(binding);
    const database = await this.#databasePromise;
    const transaction = database.transaction(RELEASE_LEASE_STORE_NAME, 'readwrite');
    let released = false;
    let operationFailure: unknown;
    const completion = transactionCompletion(transaction, () => operationFailure);
    const store = transaction.objectStore(RELEASE_LEASE_STORE_NAME);
    const request = store.get(binding.leaseId) as IDBRequest<ReleaseLeaseRecord | undefined>;
    request.onsuccess = () => {
      try {
        const existing = request.result;
        if (!existing) return;
        validateLeaseRecord(existing);
        if (!sameBinding(existing, binding)) {
          leaseFail('mismatch', `lease ${binding.leaseId} binding mismatch`);
        }
        store.delete(binding.leaseId);
        released = true;
      } catch (error) {
        operationFailure = error;
        transaction.abort();
      }
    };
    request.onerror = () => {
      operationFailure = request.error || new Error('release lease read failed');
    };
    await completion;
    return released;
  }

  async reapExpired(nowMs: number = currentTime(this.#now)): Promise<string[]> {
    if (!isSafeTimestamp(nowMs)) leaseFail('clock', 'invalid lease reap timestamp');
    const database = await this.#databasePromise;
    const transaction = database.transaction(RELEASE_LEASE_STORE_NAME, 'readwrite');
    let expired: string[] = [];
    let operationFailure: unknown;
    const completion = transactionCompletion(transaction, () => operationFailure);
    const store = transaction.objectStore(RELEASE_LEASE_STORE_NAME);
    const request = store.getAll() as IDBRequest<ReleaseLeaseRecord[]>;
    request.onsuccess = () => {
      try {
        for (const record of request.result) validateLeaseRecord(record);
        expired = request.result
          .filter((record) => record.expiresAtMs <= nowMs)
          .map((record) => record.leaseId)
          .sort();
        for (const leaseId of expired) store.delete(leaseId);
      } catch (error) {
        operationFailure = error;
        transaction.abort();
      }
    };
    request.onerror = () => {
      operationFailure = request.error || new Error('release lease scan failed');
    };
    await completion;
    return expired;
  }

  async listActiveLeases(nowMs: number = currentTime(this.#now)): Promise<ReleaseLeaseRecord[]> {
    if (!isSafeTimestamp(nowMs)) leaseFail('clock', 'invalid lease query timestamp');
    const snapshot = await this.snapshot();
    return snapshot.leases
      .filter((lease) => lease.expiresAtMs > nowMs)
      .map(cloneLease)
      .sort((left, right) => left.leaseId.localeCompare(right.leaseId));
  }

  async snapshot(): Promise<ReleaseLeaseLedgerSnapshot> {
    const database = await this.#databasePromise;
    const transaction = database.transaction(RELEASE_LEASE_STORE_NAME, 'readonly');
    const completion = transactionCompletion(transaction, () => null);
    const records = await requestResult(
      transaction.objectStore(RELEASE_LEASE_STORE_NAME).getAll() as IDBRequest<ReleaseLeaseRecord[]>,
      'release lease snapshot read',
    );
    await completion;
    for (const record of records) validateLeaseRecord(record);
    return {
      version: 1,
      leases: records.map(cloneLease).sort((left, right) => left.leaseId.localeCompare(right.leaseId)),
    };
  }
}

export function createReleaseLeaseBinding(
  identity: Omit<ReleaseLeaseBinding, 'leaseId'>,
  randomUUID: () => string = () => crypto.randomUUID(),
): ReleaseLeaseBinding {
  const binding: ReleaseLeaseBinding = {
    ...identity,
    leaseId: `lease-${randomUUID()}`,
  };
  validateBinding(binding);
  return binding;
}

export class ReleaseLeaseHeartbeat {
  readonly #ledger: ReleaseLeaseStore;
  readonly #binding: ReleaseLeaseBinding;
  readonly #ttlMs: number;
  readonly #heartbeatIntervalMs: number;
  readonly #setTimeout: typeof globalThis.setTimeout;
  readonly #clearTimeout: typeof globalThis.clearTimeout;
  readonly #onError: ((error: unknown) => void) | null;
  #timer: ReturnType<typeof globalThis.setTimeout> | null = null;
  #inFlight: Promise<void> = Promise.resolve();
  #running = false;

  constructor(options: {
    ledger: ReleaseLeaseStore;
    binding: ReleaseLeaseBinding;
    ttlMs?: number;
    heartbeatIntervalMs?: number;
    setTimeout?: typeof globalThis.setTimeout;
    clearTimeout?: typeof globalThis.clearTimeout;
    onError?: (error: unknown) => void;
  }) {
    validateBinding(options.binding);
    this.#ledger = options.ledger;
    this.#binding = { ...options.binding };
    this.#ttlMs = options.ttlMs ?? RELEASE_LEASE_DEFAULT_TTL_MS;
    this.#heartbeatIntervalMs = options.heartbeatIntervalMs ?? RELEASE_LEASE_HEARTBEAT_INTERVAL_MS;
    validateTtl(this.#ttlMs);
    if (
      !Number.isSafeInteger(this.#heartbeatIntervalMs) ||
      this.#heartbeatIntervalMs <= 0 ||
      this.#heartbeatIntervalMs >= this.#ttlMs
    ) {
      leaseFail('invalid-binding', 'lease heartbeat interval must be shorter than its TTL');
    }
    // Window timer functions are Web IDL methods in Chromium and reject an
    // arbitrary receiver with "Illegal invocation". Private-field calls use
    // the heartbeat instance as `this`, so bind browser-native defaults to
    // their global object. Explicit test/injected schedulers keep their own
    // caller-defined semantics.
    this.#setTimeout = options.setTimeout ?? globalThis.setTimeout.bind(globalThis);
    this.#clearTimeout = options.clearTimeout ?? globalThis.clearTimeout.bind(globalThis);
    this.#onError = options.onError ?? null;
  }

  get binding(): ReleaseLeaseBinding {
    return { ...this.#binding };
  }

  get running(): boolean {
    return this.#running;
  }

  async start(): Promise<ReleaseLeaseRecord> {
    if (this.#running) {
      const active = await this.#ledger.listActiveLeases();
      const existing = active.find((lease) => sameBinding(lease, this.#binding));
      if (!existing) leaseFail('unknown', `lease ${this.#binding.leaseId} is not active`);
      return cloneLease(existing);
    }
    const record = await this.#ledger.acquireOrRenewLease(this.#binding, this.#ttlMs);
    this.#running = true;
    this.#schedule();
    return cloneLease(record);
  }

  async heartbeatNow(): Promise<ReleaseLeaseRecord> {
    if (!this.#running) leaseFail('unknown', `lease ${this.#binding.leaseId} is not running`);
    try {
      return await this.#ledger.heartbeatLease(this.#binding, this.#ttlMs);
    } catch (error) {
      if (error instanceof ReleaseLeaseError && (error.code === 'expired' || error.code === 'unknown')) {
        return this.#ledger.acquireOrRenewLease(this.#binding, this.#ttlMs);
      }
      throw error;
    }
  }

  async stop(options: { release?: boolean } = {}): Promise<void> {
    const release = options.release ?? true;
    this.#running = false;
    if (this.#timer !== null) this.#clearTimeout(this.#timer);
    this.#timer = null;
    await this.#inFlight.catch(() => undefined);
    if (release) await this.#ledger.releaseLease(this.#binding);
  }

  #schedule(): void {
    if (!this.#running || this.#timer !== null) return;
    this.#timer = this.#setTimeout(() => {
      this.#timer = null;
      if (!this.#running) return;
      this.#inFlight = this.heartbeatNow()
        .then(() => undefined)
        .catch((error) => {
          this.#onError?.(error);
        })
        .finally(() => this.#schedule());
    }, this.#heartbeatIntervalMs);
  }
}

const RETENTION_REASON_ORDER: readonly ReleaseGcRetentionReason[] = [
  'transaction-installing',
  'transaction-prepared',
  'transaction-active',
  'transaction-retained',
  'live-editor-lease',
  'piwork-descriptor',
  'stable-pointer',
  'recent-release',
  'minimum-retention-window',
];

function emptyBlockedPlan(blockedReason: ReleaseGcBlockedReason): ReleaseGcPlan {
  return {
    blocked: true,
    blockedReason,
    retainedReleaseIds: [],
    retentionMarks: [],
    deleteReleaseIds: [],
    deleteObjectSha256: [],
  };
}

/**
 * Physical deletion is intentionally disabled until stable-channel and Piwork
 * descriptor references can be read in the same trusted GC transaction. A
 * missing heartbeat alone is never sufficient evidence that a background-
 * frozen editor stopped using its pinned release.
 */
export function conservativeNoDeleteReleaseGcPlan(): ReleaseGcPlan {
  return emptyBlockedPlan('execution-disabled');
}

function isDigest(value: unknown): value is string {
  return typeof value === 'string' && DIGEST_PATTERN.test(value);
}

function validateStringSet(values: unknown, validator: (value: unknown) => value is string): string[] {
  if (!Array.isArray(values)) throw new Error('expected an array');
  const result = new Set<string>();
  for (const value of values) {
    if (!validator(value) || result.has(value)) throw new Error('invalid or duplicate identifier');
    result.add(value);
  }
  return [...result];
}

function validateGcInput(input: ReleaseGcInput): {
  nowMs: number;
  releases: ReleaseGcReleaseRecord[];
  knownObjects: Set<string>;
  liveLeases: ReleaseLeaseRecord[];
  piworkDescriptors: Set<string>;
  stableReleases: Set<string>;
} {
  if (!isSafeTimestamp(input.nowMs)) throw new Error('invalid GC timestamp');
  const knownObjects = new Set(validateStringSet(input.knownObjectSha256, isDigest));
  if (!Array.isArray(input.releases)) throw new Error('invalid releases');
  const releaseIds = new Set<string>();
  const releases = input.releases.map((release) => {
    if (
      !release ||
      !isResourceBrokerReleaseId(release.releaseId) ||
      releaseIds.has(release.releaseId) ||
      !['installing', 'prepared', 'active', 'retained', 'failed'].includes(release.state) ||
      !isSafeTimestamp(release.publishedAtMs)
    ) {
      throw new Error('invalid release record');
    }
    releaseIds.add(release.releaseId);
    const objectSha256 = validateStringSet(release.objectSha256, isDigest);
    if (objectSha256.some((sha256) => !knownObjects.has(sha256))) {
      throw new Error('release references an unknown object');
    }
    return { ...release, objectSha256: objectSha256.sort() };
  });
  if (releases.filter((release) => release.state === 'active').length > 1) {
    throw new Error('multiple active releases');
  }

  const leaseLedger = new MemoryReleaseLeaseLedger({ snapshot: input.leaseLedger });
  const liveLeases = leaseLedger.listActiveLeases(input.nowMs);
  if (liveLeases.some((lease) => !releaseIds.has(lease.releaseId))) {
    throw new Error('live lease references an unknown release');
  }
  const piworkDescriptors = new Set(validateStringSet(input.piworkDescriptorReleaseIds, isResourceBrokerReleaseId));
  const stableReleases = new Set(validateStringSet(input.stableReleaseIds, isResourceBrokerReleaseId));
  if ([...piworkDescriptors, ...stableReleases].some((releaseId) => !releaseIds.has(releaseId))) {
    throw new Error('external pointer references an unknown release');
  }

  return {
    nowMs: input.nowMs,
    releases,
    knownObjects,
    liveLeases,
    piworkDescriptors,
    stableReleases,
  };
}

/**
 * Computes a deterministic mark-and-sweep plan without mutating Cache Storage,
 * IndexedDB, or any caller-owned data. Unknown or malformed ledger state always
 * returns an empty deletion set.
 */
export function planReleaseGarbageCollection(input: ReleaseGcInput): ReleaseGcPlan {
  if (!input || input.ledgerState === 'unknown') return emptyBlockedPlan('ledger-unknown');
  if (input.ledgerState !== 'valid') return emptyBlockedPlan('ledger-corrupt');

  try {
    const validated = validateGcInput(input);
    const reasonsByRelease = new Map<string, Set<ReleaseGcRetentionReason>>(
      validated.releases.map((release) => [release.releaseId, new Set<ReleaseGcRetentionReason>()]),
    );
    const mark = (releaseId: string, reason: ReleaseGcRetentionReason): void => {
      reasonsByRelease.get(releaseId)!.add(reason);
    };

    for (const release of validated.releases) {
      if (PROTECTED_TRANSACTION_STATES.has(release.state)) {
        mark(release.releaseId, `transaction-${release.state}` as ReleaseGcRetentionReason);
      }
      if (validated.nowMs - release.publishedAtMs <= RELEASE_GC_MINIMUM_RETENTION_MS) {
        mark(release.releaseId, 'minimum-retention-window');
      }
    }
    for (const releaseId of validated.piworkDescriptors) mark(releaseId, 'piwork-descriptor');
    for (const releaseId of validated.stableReleases) mark(releaseId, 'stable-pointer');
    for (const lease of validated.liveLeases) mark(lease.releaseId, 'live-editor-lease');

    const recentReleaseIds = [...validated.releases]
      .sort((left, right) => right.publishedAtMs - left.publishedAtMs || left.releaseId.localeCompare(right.releaseId))
      .slice(0, RELEASE_GC_MINIMUM_RECENT_RELEASES)
      .map((release) => release.releaseId);
    for (const releaseId of recentReleaseIds) mark(releaseId, 'recent-release');

    const retainedReleaseIds = [...reasonsByRelease]
      .filter(([, reasons]) => reasons.size > 0)
      .map(([releaseId]) => releaseId)
      .sort();
    const retainedSet = new Set(retainedReleaseIds);
    const deleteReleaseIds = validated.releases
      .map((release) => release.releaseId)
      .filter((releaseId) => !retainedSet.has(releaseId))
      .sort();

    const referencedByRetainedRelease = new Set<string>();
    for (const release of validated.releases) {
      if (!retainedSet.has(release.releaseId)) continue;
      for (const sha256 of release.objectSha256) referencedByRetainedRelease.add(sha256);
    }
    const deleteObjectSha256 = [...validated.knownObjects]
      .filter((sha256) => !referencedByRetainedRelease.has(sha256))
      .sort();
    const retentionMarks = retainedReleaseIds.map((releaseId) => ({
      releaseId,
      reasons: RETENTION_REASON_ORDER.filter((reason) => reasonsByRelease.get(releaseId)!.has(reason)),
    }));

    return {
      blocked: false,
      blockedReason: null,
      retainedReleaseIds,
      retentionMarks,
      deleteReleaseIds,
      deleteObjectSha256,
    };
  } catch {
    return emptyBlockedPlan('ledger-corrupt');
  }
}
