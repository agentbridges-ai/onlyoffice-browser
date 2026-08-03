import { sha256 } from '@noble/hashes/sha2.js';
import {
  MemoryReleaseContentJournal,
  planReleaseContent,
  type AssetContentMapping,
  type ContentObjectDescriptor,
  type ContentObjectRecord,
  type ReleaseContentJournalSnapshot,
  type ReleaseContentModel,
  type ReleaseContentPlan,
  type ReleaseContentReadView,
  type ReleaseResumeState,
  type ReleaseTransactionRecord,
} from './release-content-model';
import { conservativeNoDeleteReleaseGcPlan, type ReleaseGcPlan } from './release-lease-gc';

export const CANONICAL_CONTENT_CACHE_NAME = 'onlyoffice-content-v1';
export const CANONICAL_CONTENT_PATH_PREFIX = '/__onlyoffice_content__/sha256/';
export const CANONICAL_CONTENT_DIGEST_HEADER = 'x-onlyoffice-content-sha256';
export const CANONICAL_CONTENT_BYTES_HEADER = 'x-onlyoffice-content-bytes';

const DEFAULT_CACHE_KEY_ORIGIN = 'https://onlyoffice.getpi.work';
const TRANSFER_RETRY_QUERY = '__onlyoffice_transfer_retry';
const JOURNAL_DATABASE_NAME = 'onlyoffice-content-journal-v2';
const JOURNAL_DATABASE_VERSION = 1;
const JOURNAL_LOCK_NAME = 'onlyoffice-content-journal-v2';
const OBJECT_LOCK_PREFIX = 'onlyoffice-content-object-v1:';
const RELEASE_LOCK_PREFIX = 'onlyoffice-content-release-v1:';
const ACTIVE_RELEASE_KEY = 'active-release-id';
const DIGEST_PATTERN = /^[a-f0-9]{64}$/;
const HEALTH_CHECK_CONCURRENCY = 32;
const DEEP_VERIFY_CONCURRENCY = 4;
const MAX_PACKAGE_INGEST_BYTES = 1_073_741_824;
const MAX_PACKAGE_INGEST_SEGMENT_BYTES = 64 * 1_024 * 1_024;
const MAX_PACKAGE_INGEST_ASSET_BYTES = 64 * 1_024 * 1_024;

export type CanonicalResourceStoreErrorCode =
  | 'invalid-content'
  | 'network'
  | 'http'
  | 'integrity'
  | 'storage'
  | 'incomplete'
  | 'invalid-state';

export class CanonicalResourceStoreError extends Error {
  constructor(
    readonly code: CanonicalResourceStoreErrorCode,
    message: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = 'CanonicalResourceStoreError';
  }
}

export interface CanonicalResourceJournal {
  beginRelease(
    plan: ReleaseContentPlan,
    identity: Pick<ReleaseContentModel, 'manifestSha256' | 'storageSetSha256'>,
  ): Promise<ReleaseTransactionRecord>;
  recordVerifiedObject(releaseId: string, object: ContentObjectDescriptor, verifiedAt?: number): Promise<void>;
  commitAssetMapping(releaseId: string, path: string): Promise<AssetContentMapping>;
  commitAssetMappings(releaseId: string, paths: readonly string[]): Promise<AssetContentMapping[]>;
  prepareRelease(releaseId: string): Promise<ReleaseTransactionRecord>;
  activateRelease(releaseId: string): Promise<ReleaseTransactionRecord>;
  reactivateRetainedRelease(releaseId: string): Promise<ReleaseTransactionRecord>;
  rollbackActivation(releaseId: string, failureCode: string): Promise<ReleaseTransactionRecord>;
  failRelease(releaseId: string, failureCode: string): Promise<ReleaseTransactionRecord>;
  resumeRelease(releaseId: string): Promise<ReleaseTransactionRecord>;
  getResumeState(releaseId: string): Promise<ReleaseResumeState>;
  getActiveRelease(): Promise<ReleaseTransactionRecord | null>;
  getTransaction(releaseId: string): Promise<ReleaseTransactionRecord | null>;
  listObjects(): Promise<ContentObjectRecord[]>;
  getObject(sha256: string): Promise<ContentObjectRecord | null>;
  getObjects(sha256s: readonly string[]): Promise<ContentObjectRecord[]>;
  listAssetMappings(releaseId: string): Promise<AssetContentMapping[]>;
  getReleaseReadView(releaseId?: string): Promise<ReleaseContentReadView>;
  getAssetReadView(releaseId: string, path: string): Promise<CanonicalAssetReadView>;
  getReleaseProbeView(releaseId: string): Promise<CanonicalReleaseProbeView>;
}

export interface CanonicalReleaseRuntimeRecord {
  releaseId: string;
  manifestSha256: string;
  storageSetSha256: string;
  state: ReleaseTransactionRecord['state'];
  updatedAt: number;
  probePath?: string;
  probeAssetBytes?: number;
  probeAssetSha256?: string;
}

export interface CanonicalAssetReadView {
  release: CanonicalReleaseRuntimeRecord | null;
  mapping: AssetContentMapping | null;
  objects: ContentObjectRecord[];
}

export interface CanonicalReleaseProbeView {
  release: CanonicalReleaseRuntimeRecord | null;
  mapping: AssetContentMapping | null;
  object: ContentObjectRecord | null;
}

export interface CanonicalResourceProbe {
  releaseId: string | null;
  state: ReleaseTransactionRecord['state'] | 'missing';
  ready: boolean;
  probeSucceeded: boolean;
  probePath: string | null;
  probeAssetBytes: number | null;
  probeAssetSha256: string | null;
}

export interface CanonicalObjectProgress {
  releaseId: string;
  object: ContentObjectDescriptor;
  chunkBytes: number;
  loadedBytes: number;
}

export interface CanonicalObjectVerified {
  releaseId: string;
  object: ContentObjectDescriptor;
}

export interface CanonicalResourceHealth {
  releaseId: string | null;
  state: ReleaseTransactionRecord['state'] | 'missing';
  ready: boolean;
  missingObjects: ContentObjectDescriptor[];
  probeSucceeded: boolean;
}

export type CanonicalResourceIntegrityFailureCode = 'missing' | 'integrity' | 'read';

export interface CanonicalResourceIntegrityFailure {
  object: ContentObjectDescriptor;
  code: CanonicalResourceIntegrityFailureCode;
  removed: boolean;
  actualBytes?: number;
  actualSha256?: string;
}

export interface CanonicalResourceIntegrityVerification {
  releaseId: string | null;
  state: ReleaseTransactionRecord['state'] | 'missing';
  status: 'complete' | 'aborted';
  ready: boolean;
  checkedObjects: number;
  checkedBytes: number;
  verifiedObjects: number;
  verifiedBytes: number;
  failures: CanonicalResourceIntegrityFailure[];
}

export interface CanonicalResourceIntegrityVerificationOptions {
  releaseId?: string;
  objectSha256s?: readonly string[];
  signal?: AbortSignal;
}

type CanonicalObjectIntegrityOutcome =
  | {
      status: 'verified';
      bytes: number;
    }
  | {
      status: 'failed';
      failure: CanonicalResourceIntegrityFailure;
    }
  | {
      status: 'aborted';
    };

export interface CanonicalResourceStoreOptions {
  cacheStorage: CacheStorage;
  journal: CanonicalResourceJournal;
  fetch: typeof fetch;
  objectUrl: (releaseId: string, object: ContentObjectDescriptor) => RequestInfo | URL;
  locks?: LockManager;
  cacheKeyOrigin?: string | URL;
  maxConcurrentDownloads?: number;
  now?: () => number;
  onObjectProgress?: (progress: CanonicalObjectProgress) => void | Promise<void>;
  onObjectVerified?: (verified: CanonicalObjectVerified) => void | Promise<void>;
}

export interface CanonicalPackageTransportSegment extends ContentObjectDescriptor {
  offset: number;
}

export interface CanonicalPackageTransportAsset {
  path: string;
  packageOffset: number;
  bytes: number;
  sha256: string;
  mime?: string;
}

/**
 * Release-bound package metadata used only as a cold-install transport.
 *
 * Package segments are never Cache Storage keys or runtime mappings. The
 * installer streams their bytes into independently verified whole-file CAS
 * responses, so Broker reads remain O(1) whole-object lookups.
 */
export interface CanonicalPackageTransport {
  bytes: number;
  headerBytes: number;
  segments: CanonicalPackageTransportSegment[];
  assets: CanonicalPackageTransportAsset[];
}

function cloneDescriptor(object: ContentObjectDescriptor): ContentObjectDescriptor {
  return { ...object };
}

function transferInputForAttempt(input: RequestInfo | URL, attempt: number): RequestInfo | URL {
  if (!Number.isSafeInteger(attempt) || attempt <= 0) return input;
  const href = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
  const url = new URL(href);
  url.searchParams.set(TRANSFER_RETRY_QUERY, String(attempt));
  return url;
}

function cloneRecord(object: ContentObjectRecord): ContentObjectRecord {
  return { ...object };
}

function cloneMapping(mapping: AssetContentMapping): AssetContentMapping {
  return {
    ...mapping,
    spans: mapping.spans.map((span) => ({ ...span })),
  };
}

function cloneTransaction(transaction: ReleaseTransactionRecord): ReleaseTransactionRecord {
  return {
    ...transaction,
    requiredObjects: transaction.requiredObjects.map(cloneDescriptor),
    plannedMappings: transaction.plannedMappings.map(cloneMapping),
    committedMappings: transaction.committedMappings.map(cloneMapping),
  };
}

function probeMappingPriority(path: string): number {
  if (path.startsWith('sdkjs/')) return 0;
  if (path.startsWith('web-apps/')) return 1;
  if (path.startsWith('wasm/')) return 2;
  if (path.startsWith('fonts/')) return 3;
  return 4;
}

function runtimeRecord(transaction: ReleaseTransactionRecord): CanonicalReleaseRuntimeRecord {
  let probeMapping: AssetContentMapping | undefined;
  for (const mapping of transaction.committedMappings) {
    if (
      mapping.assetBytes <= 0 ||
      mapping.path === 'office-host.html' ||
      mapping.path === 'editor-shell-prime.html' ||
      mapping.path.startsWith('assets/')
    ) {
      continue;
    }
    if (
      !probeMapping ||
      probeMappingPriority(mapping.path) < probeMappingPriority(probeMapping.path) ||
      (probeMappingPriority(mapping.path) === probeMappingPriority(probeMapping.path) &&
        mapping.path.localeCompare(probeMapping.path) < 0)
    ) {
      probeMapping = mapping;
    }
  }
  return {
    releaseId: transaction.releaseId,
    manifestSha256: transaction.manifestSha256,
    storageSetSha256: transaction.storageSetSha256,
    state: transaction.state,
    updatedAt: transaction.updatedAt,
    ...(probeMapping
      ? {
          probePath: probeMapping.path,
          probeAssetBytes: probeMapping.assetBytes,
          probeAssetSha256: probeMapping.assetSha256,
        }
      : {}),
  };
}

function assertDigest(digest: string): void {
  if (!DIGEST_PATTERN.test(digest)) {
    throw new CanonicalResourceStoreError('invalid-content', 'content object digest must be lowercase SHA-256');
  }
}

function bytesToHex(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function failureCode(error: unknown): string {
  if (error instanceof CanonicalResourceStoreError) return error.code;
  if (error instanceof DOMException && error.name === 'AbortError') return 'aborted';
  return 'storage';
}

function getDefaultLocks(): LockManager | undefined {
  return typeof navigator !== 'undefined' && 'locks' in navigator ? navigator.locks : undefined;
}

export function canonicalContentCacheKey(digest: string, origin: string | URL = DEFAULT_CACHE_KEY_ORIGIN): string {
  assertDigest(digest);
  return new URL(`${CANONICAL_CONTENT_PATH_PREFIX}${digest}`, origin).href;
}

class LocalExclusiveLocks {
  private readonly tails = new Map<string, Promise<void>>();

  async request<T>(name: string, callback: () => Promise<T>): Promise<T> {
    const previous = this.tails.get(name) || Promise.resolve();
    let release!: () => void;
    const current = new Promise<void>((resolve) => {
      release = resolve;
    });
    const tail = previous.catch(() => undefined).then(() => current);
    this.tails.set(name, tail);
    await previous.catch(() => undefined);
    try {
      return await callback();
    } finally {
      release();
      if (this.tails.get(name) === tail) this.tails.delete(name);
    }
  }
}

const localExclusiveLocks = new LocalExclusiveLocks();

async function withExclusiveLock<T>(
  name: string,
  locks: LockManager | undefined,
  callback: () => Promise<T>,
): Promise<T> {
  if (locks) {
    return locks.request(name, { mode: 'exclusive' }, callback);
  }
  return localExclusiveLocks.request(name, callback);
}

function packageSegmentsForAsset(
  transport: CanonicalPackageTransport,
  asset: CanonicalPackageTransportAsset,
): CanonicalPackageTransportSegment[] {
  const start = asset.packageOffset;
  const end = start + asset.bytes;
  if (asset.bytes === 0) return [];
  return transport.segments.filter((segment) => {
    const segmentEnd = segment.offset + segment.bytes;
    return Math.max(start, segment.offset) < Math.min(end, segmentEnd);
  });
}

function validatePackageTransport(
  model: ReleaseContentModel,
  transport: CanonicalPackageTransport,
): Map<string, CanonicalPackageTransportAsset[]> {
  if (
    !Number.isSafeInteger(transport.bytes) ||
    transport.bytes <= 0 ||
    transport.bytes > MAX_PACKAGE_INGEST_BYTES ||
    !Number.isSafeInteger(transport.headerBytes) ||
    transport.headerBytes <= 0 ||
    transport.headerBytes > transport.bytes
  ) {
    throw new CanonicalResourceStoreError('invalid-content', 'invalid package transport bounds');
  }
  const modelObjects = new Map(model.objects.map((object) => [object.sha256, object]));
  let expectedSegmentOffset = 0;
  const segmentDigests = new Set<string>();
  for (const segment of transport.segments) {
    assertDigest(segment.sha256);
    if (
      !Number.isSafeInteger(segment.offset) ||
      segment.offset !== expectedSegmentOffset ||
      !Number.isSafeInteger(segment.bytes) ||
      segment.bytes <= 0 ||
      segment.bytes > MAX_PACKAGE_INGEST_SEGMENT_BYTES ||
      segmentDigests.has(segment.sha256)
    ) {
      throw new CanonicalResourceStoreError('invalid-content', 'invalid package transport segment');
    }
    const modelObject = modelObjects.get(segment.sha256);
    if (!modelObject || modelObject.bytes !== segment.bytes) {
      throw new CanonicalResourceStoreError('invalid-content', 'package transport segment is outside the release');
    }
    segmentDigests.add(segment.sha256);
    expectedSegmentOffset += segment.bytes;
  }
  if (expectedSegmentOffset !== transport.bytes) {
    throw new CanonicalResourceStoreError('invalid-content', 'package transport does not cover the package');
  }

  const modelAssets = new Map(model.assets.map((asset) => [asset.path, asset]));
  const transportPaths = new Set<string>();
  const assetsByDigest = new Map<string, CanonicalPackageTransportAsset[]>();
  const sortedAssets = [...transport.assets].sort(
    (left, right) => left.packageOffset - right.packageOffset || left.path.localeCompare(right.path),
  );
  let expectedAssetOffset = transport.headerBytes;
  for (const asset of sortedAssets) {
    assertDigest(asset.sha256);
    if (
      transportPaths.has(asset.path) ||
      !Number.isSafeInteger(asset.packageOffset) ||
      asset.packageOffset !== expectedAssetOffset ||
      !Number.isSafeInteger(asset.bytes) ||
      asset.bytes < 0 ||
      asset.bytes > MAX_PACKAGE_INGEST_ASSET_BYTES ||
      asset.packageOffset + asset.bytes > transport.bytes
    ) {
      throw new CanonicalResourceStoreError('invalid-content', `invalid package asset ${asset.path}`);
    }
    const modelAsset = modelAssets.get(asset.path);
    const whole = modelAsset?.representations.find((representation) => representation.kind === 'whole');
    const packaged = modelAsset?.representations.find((representation) => representation.kind === 'package');
    if (
      !modelAsset ||
      modelAsset.bytes !== asset.bytes ||
      modelAsset.sha256 !== asset.sha256 ||
      modelAsset.mime !== asset.mime ||
      !whole ||
      whole.spans.length !== 1 ||
      whole.spans[0].objectSha256 !== asset.sha256 ||
      whole.spans[0].bytes !== asset.bytes ||
      !packaged
    ) {
      throw new CanonicalResourceStoreError('invalid-content', `package asset ${asset.path} conflicts with release`);
    }
    const expectedSpans = packageSegmentsForAsset(transport, asset).map((segment) => {
      const overlapStart = Math.max(asset.packageOffset, segment.offset);
      const overlapEnd = Math.min(asset.packageOffset + asset.bytes, segment.offset + segment.bytes);
      return {
        objectSha256: segment.sha256,
        objectOffset: overlapStart - segment.offset,
        assetOffset: overlapStart - asset.packageOffset,
        bytes: overlapEnd - overlapStart,
      };
    });
    if (asset.bytes === 0) {
      const container =
        transport.segments.find(
          (segment) => asset.packageOffset >= segment.offset && asset.packageOffset <= segment.offset + segment.bytes,
        ) || transport.segments.at(-1);
      if (!container) {
        throw new CanonicalResourceStoreError('invalid-content', `empty package asset ${asset.path} has no segment`);
      }
      expectedSpans.push({
        objectSha256: container.sha256,
        objectOffset: asset.packageOffset - container.offset,
        assetOffset: 0,
        bytes: 0,
      });
    }
    if (JSON.stringify(packaged.spans) !== JSON.stringify(expectedSpans)) {
      throw new CanonicalResourceStoreError('invalid-content', `package spans for ${asset.path} are inconsistent`);
    }
    transportPaths.add(asset.path);
    expectedAssetOffset += asset.bytes;
    const candidates = assetsByDigest.get(asset.sha256) || [];
    candidates.push({ ...asset });
    assetsByDigest.set(asset.sha256, candidates);
  }
  if (transportPaths.size !== modelAssets.size || expectedAssetOffset !== transport.bytes) {
    throw new CanonicalResourceStoreError('invalid-content', 'package assets do not cover the release payload');
  }
  return assetsByDigest;
}

type CanonicalPackageIngestUnit = ContentObjectDescriptor & {
  path: string;
  packageOffset: number;
  mime?: string;
};

function packageSegmentsForRange(
  transport: CanonicalPackageTransport,
  offset: number,
  bytes: number,
): CanonicalPackageTransportSegment[] {
  if (bytes === 0) return [];
  const end = offset + bytes;
  return transport.segments.filter((segment) => {
    const segmentEnd = segment.offset + segment.bytes;
    return Math.max(offset, segment.offset) < Math.min(end, segmentEnd);
  });
}

function createCanonicalPackageIngestPlan(
  model: ReleaseContentModel,
  transport: CanonicalPackageTransport,
  plan: Pick<ReleaseContentPlan, 'mappings' | 'missingObjects'>,
): { segments: CanonicalPackageTransportSegment[]; units: CanonicalPackageIngestUnit[] } {
  validatePackageTransport(model, transport);
  const transportAssets = new Map(transport.assets.map((asset) => [asset.path, asset]));
  const modelObjects = new Map(model.objects.map((object) => [object.sha256, object]));
  const missing = new Map(plan.missingObjects.map((object) => [object.sha256, object]));
  const candidates = new Map<string, CanonicalPackageIngestUnit[]>();
  for (const mapping of plan.mappings) {
    if (mapping.representationKind === 'package') {
      throw new CanonicalResourceStoreError('invalid-content', 'package transport cannot become an active mapping');
    }
    const asset = transportAssets.get(mapping.path);
    if (!asset) {
      throw new CanonicalResourceStoreError('invalid-content', `package transport is missing ${mapping.path}`);
    }
    for (const span of mapping.spans) {
      const object = modelObjects.get(span.objectSha256);
      if (
        !object ||
        !missing.has(object.sha256) ||
        span.objectOffset !== 0 ||
        span.bytes !== object.bytes ||
        span.assetOffset + span.bytes > asset.bytes
      ) {
        continue;
      }
      const units = candidates.get(object.sha256) || [];
      units.push({
        path: mapping.path,
        packageOffset: asset.packageOffset + span.assetOffset,
        bytes: object.bytes,
        sha256: object.sha256,
        mime: mapping.mime,
      });
      candidates.set(object.sha256, units);
    }
  }
  const selected = new Map<string, CanonicalPackageTransportSegment>();
  const selectedUnits: CanonicalPackageIngestUnit[] = [];
  for (const object of plan.missingObjects) {
    assertDigest(object.sha256);
    const objectCandidates = (candidates.get(object.sha256) || []).filter((unit) => unit.bytes === object.bytes);
    if (objectCandidates.length === 0) {
      throw new CanonicalResourceStoreError(
        'invalid-content',
        `planned object ${object.sha256} has no package byte range`,
      );
    }
    const unit = [...objectCandidates].sort((left, right) => {
      const leftSegments = packageSegmentsForRange(transport, left.packageOffset, left.bytes);
      const rightSegments = packageSegmentsForRange(transport, right.packageOffset, right.bytes);
      return (
        leftSegments.reduce((sum, segment) => sum + segment.bytes, 0) -
          rightSegments.reduce((sum, segment) => sum + segment.bytes, 0) ||
        leftSegments.length - rightSegments.length ||
        left.packageOffset - right.packageOffset ||
        left.path.localeCompare(right.path)
      );
    })[0];
    selectedUnits.push(unit);
    for (const segment of packageSegmentsForRange(transport, unit.packageOffset, unit.bytes)) {
      selected.set(segment.sha256, segment);
    }
  }
  return {
    segments: [...selected.values()].sort((left, right) => left.offset - right.offset),
    units: selectedUnits.sort(
      (left, right) => left.packageOffset - right.packageOffset || left.path.localeCompare(right.path),
    ),
  };
}

export function planCanonicalPackageTransfer(
  model: ReleaseContentModel,
  transport: CanonicalPackageTransport,
  plan: Pick<ReleaseContentPlan, 'mappings' | 'missingObjects'>,
): CanonicalPackageTransportSegment[] {
  return createCanonicalPackageIngestPlan(model, transport, plan).segments;
}

export class MemoryCanonicalResourceJournal implements CanonicalResourceJournal {
  private journal: MemoryReleaseContentJournal;
  private readonly now: () => number;
  private queue: Promise<void> = Promise.resolve();

  constructor(options: { snapshot?: ReleaseContentJournalSnapshot; now?: () => number } = {}) {
    this.now = options.now || Date.now;
    this.journal = new MemoryReleaseContentJournal({ snapshot: options.snapshot, now: this.now });
  }

  private async read<T>(read: (journal: MemoryReleaseContentJournal) => T): Promise<T> {
    await this.queue;
    return read(this.journal);
  }

  private async write<T>(write: (journal: MemoryReleaseContentJournal) => T): Promise<T> {
    const operation = this.queue.then(() => write(this.journal));
    this.queue = operation.then(
      () => undefined,
      () => undefined,
    );
    return operation;
  }

  async beginRelease(
    plan: ReleaseContentPlan,
    identity: Pick<ReleaseContentModel, 'manifestSha256' | 'storageSetSha256'>,
  ): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.beginRelease(plan, identity));
  }

  async recordVerifiedObject(
    releaseId: string,
    object: ContentObjectDescriptor,
    verifiedAt: number = this.now(),
  ): Promise<void> {
    await this.write((journal) => journal.recordVerifiedObject(releaseId, object, verifiedAt));
  }

  async recordVerifiedObjects(
    releaseId: string,
    objects: readonly ContentObjectDescriptor[],
    verifiedAt: number = this.now(),
  ): Promise<void> {
    await this.write((journal) => {
      for (const object of objects) journal.recordVerifiedObject(releaseId, object, verifiedAt);
    });
  }

  async commitAssetMapping(releaseId: string, path: string): Promise<AssetContentMapping> {
    return this.write((journal) => journal.commitAssetMapping(releaseId, path));
  }

  async commitAssetMappings(releaseId: string, paths: readonly string[]): Promise<AssetContentMapping[]> {
    return this.write((journal) => paths.map((path) => journal.commitAssetMapping(releaseId, path)));
  }

  async prepareRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.prepareRelease(releaseId));
  }

  async activateRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.activateRelease(releaseId));
  }

  async reactivateRetainedRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.reactivateRetainedRelease(releaseId));
  }

  async rollbackActivation(releaseId: string, failureCode: string): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.rollbackActivation(releaseId, failureCode));
  }

  async failRelease(releaseId: string, code: string): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.failRelease(releaseId, code));
  }

  async resumeRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.write((journal) => journal.resumeRelease(releaseId));
  }

  async getResumeState(releaseId: string): Promise<ReleaseResumeState> {
    return this.read((journal) => journal.getResumeState(releaseId));
  }

  async getActiveRelease(): Promise<ReleaseTransactionRecord | null> {
    return this.read((journal) => journal.getActiveRelease());
  }

  async getTransaction(releaseId: string): Promise<ReleaseTransactionRecord | null> {
    return this.read((journal) => journal.getTransaction(releaseId));
  }

  async listObjects(): Promise<ContentObjectRecord[]> {
    return this.read((journal) => journal.listObjects());
  }

  async getObject(sha256: string): Promise<ContentObjectRecord | null> {
    return this.read((journal) => journal.listObjects().find((object) => object.sha256 === sha256) || null);
  }

  async getObjects(sha256s: readonly string[]): Promise<ContentObjectRecord[]> {
    return this.read((journal) => {
      const requested = new Set(sha256s);
      return journal.listObjects().filter((object) => requested.has(object.sha256));
    });
  }

  async listAssetMappings(releaseId: string): Promise<AssetContentMapping[]> {
    return this.read((journal) => journal.listAssetMappings(releaseId));
  }

  async getReleaseReadView(releaseId?: string): Promise<ReleaseContentReadView> {
    return this.read((journal) => journal.getReleaseReadView(releaseId));
  }

  async getAssetReadView(releaseId: string, path: string): Promise<CanonicalAssetReadView> {
    return this.read((journal) => {
      const transaction = journal.getTransaction(releaseId);
      if (!transaction) return { release: null, mapping: null, objects: [] };
      const mapping = transaction.committedMappings.find((candidate) => candidate.path === path) || null;
      if (!mapping) return { release: runtimeRecord(transaction), mapping: null, objects: [] };
      const digests = new Set(mapping.spans.map((span) => span.objectSha256));
      return {
        release: runtimeRecord(transaction),
        mapping: cloneMapping(mapping),
        objects: journal.listObjects().filter((object) => digests.has(object.sha256)),
      };
    });
  }

  async getReleaseProbeView(releaseId: string): Promise<CanonicalReleaseProbeView> {
    return this.read((journal) => {
      const transaction = journal.getTransaction(releaseId);
      if (!transaction) return { release: null, mapping: null, object: null };
      const runtime = runtimeRecord(transaction);
      const mapping = runtime.probePath
        ? transaction.committedMappings.find((candidate) => candidate.path === runtime.probePath) || null
        : null;
      const firstDigest = mapping?.spans[0]?.objectSha256;
      const object = firstDigest
        ? journal.listObjects().find((candidate) => candidate.sha256 === firstDigest) || null
        : null;
      return {
        release: runtime,
        mapping: mapping ? cloneMapping(mapping) : null,
        object,
      };
    });
  }

  async snapshot(): Promise<ReleaseContentJournalSnapshot> {
    return this.read((journal) => journal.snapshot());
  }
}

function requestResult<T>(request: IDBRequest<T>, label: string): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error(`${label} failed`));
  });
}

function transactionCompletion(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error('IndexedDB transaction failed'));
    transaction.onabort = () => reject(transaction.error || new Error('IndexedDB transaction aborted'));
  });
}

export class IndexedDbCanonicalResourceJournal implements CanonicalResourceJournal {
  private readonly databasePromise: Promise<IDBDatabase>;
  private readonly locks: LockManager | undefined;
  private readonly now: () => number;

  constructor(
    indexedDb: IDBFactory = indexedDB,
    options: { databaseName?: string; locks?: LockManager; now?: () => number } = {},
  ) {
    this.locks = options.locks ?? getDefaultLocks();
    this.now = options.now || Date.now;
    this.databasePromise = new Promise((resolve, reject) => {
      const request = indexedDb.open(options.databaseName || JOURNAL_DATABASE_NAME, JOURNAL_DATABASE_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains('objects')) {
          database.createObjectStore('objects', { keyPath: 'sha256' });
        }
        if (!database.objectStoreNames.contains('assetMappings')) {
          const mappings = database.createObjectStore('assetMappings', {
            keyPath: ['releaseId', 'path'],
          });
          mappings.createIndex('releaseId', 'releaseId');
        }
        if (!database.objectStoreNames.contains('releaseTransactions')) {
          database.createObjectStore('releaseTransactions', { keyPath: 'releaseId' });
        }
        if (!database.objectStoreNames.contains('metadata')) {
          database.createObjectStore('metadata');
        }
        if (!database.objectStoreNames.contains('releaseRuntime')) {
          database.createObjectStore('releaseRuntime', { keyPath: 'releaseId' });
        }
        if (!database.objectStoreNames.contains('releaseObjects')) {
          const releaseObjects = database.createObjectStore('releaseObjects', {
            keyPath: ['releaseId', 'sha256'],
          });
          releaseObjects.createIndex('releaseId', 'releaseId');
        }
        if (!database.objectStoreNames.contains('plannedAssetMappings')) {
          const plannedMappings = database.createObjectStore('plannedAssetMappings', {
            keyPath: ['releaseId', 'path'],
          });
          plannedMappings.createIndex('releaseId', 'releaseId');
        }
      };
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onblocked = () => reject(new Error('IndexedDB open was blocked'));
      request.onerror = () => reject(request.error || new Error('IndexedDB open failed'));
    });
  }

  private async readSnapshot(): Promise<ReleaseContentJournalSnapshot> {
    const database = await this.databasePromise;
    const transaction = database.transaction(
      ['objects', 'assetMappings', 'releaseTransactions', 'metadata'],
      'readonly',
    );
    const completion = transactionCompletion(transaction);
    const [objects, mappings, transactions, activeReleaseId] = await Promise.all([
      requestResult(
        transaction.objectStore('objects').getAll() as IDBRequest<ContentObjectRecord[]>,
        'IndexedDB object read',
      ),
      requestResult(
        transaction.objectStore('assetMappings').getAll() as IDBRequest<AssetContentMapping[]>,
        'IndexedDB mapping read',
      ),
      requestResult(
        transaction.objectStore('releaseTransactions').getAll() as IDBRequest<ReleaseTransactionRecord[]>,
        'IndexedDB transaction read',
      ),
      requestResult(
        transaction.objectStore('metadata').get(ACTIVE_RELEASE_KEY) as IDBRequest<string | undefined>,
        'IndexedDB metadata read',
      ),
    ]);
    await completion;
    const mappingsByRelease = new Map<string, AssetContentMapping[]>();
    for (const mapping of mappings) {
      const releaseMappings = mappingsByRelease.get(mapping.releaseId) || [];
      releaseMappings.push(cloneMapping(mapping));
      mappingsByRelease.set(mapping.releaseId, releaseMappings);
    }
    return {
      version: 1,
      activeReleaseId: activeReleaseId || null,
      objects: objects.map(cloneRecord),
      transactions: transactions.map((record) => ({
        ...cloneTransaction(record),
        committedMappings: (mappingsByRelease.get(record.releaseId) || []).map(cloneMapping),
      })),
    };
  }

  private async readReleaseView(releaseId?: string): Promise<ReleaseContentReadView> {
    const database = await this.databasePromise;
    const transaction = database.transaction(['objects', 'releaseTransactions', 'metadata'], 'readonly');
    const completion = transactionCompletion(transaction);
    const objectStore = transaction.objectStore('objects');
    const transactionStore = transaction.objectStore('releaseTransactions');
    const metadataRequest = transaction.objectStore('metadata').get(ACTIVE_RELEASE_KEY) as IDBRequest<
      string | undefined
    >;

    const view = new Promise<ReleaseContentReadView>((resolve, reject) => {
      let activeReleaseId: string | null | undefined;
      let selected: ReleaseTransactionRecord | undefined;
      let selectionComplete = false;
      let objectsStarted = false;

      const failRead = (label: string, request: { error: DOMException | null }) => {
        reject(request.error || new Error(`${label} failed`));
      };
      const finishSelection = () => {
        if (objectsStarted || activeReleaseId === undefined || !selectionComplete) return;
        objectsStarted = true;
        if (!selected) {
          if (activeReleaseId && activeReleaseId === releaseId) {
            reject(new CanonicalResourceStoreError('storage', `active release ${activeReleaseId} has no journal`));
            return;
          }
          resolve({ transaction: null, objects: [] });
          return;
        }
        if (
          (selected.state === 'active' && activeReleaseId !== selected.releaseId) ||
          (activeReleaseId === selected.releaseId && selected.state !== 'active')
        ) {
          reject(
            new CanonicalResourceStoreError(
              'storage',
              `active release pointer and transaction ${selected.releaseId} disagree`,
            ),
          );
          return;
        }

        const records: ContentObjectRecord[] = [];
        let remaining = selected.requiredObjects.length;
        const finishObjects = () => {
          try {
            const scoped = new MemoryReleaseContentJournal({
              snapshot: {
                version: 1,
                activeReleaseId: selected!.state === 'active' ? selected!.releaseId : null,
                objects: records,
                transactions: [selected!],
              },
              now: this.now,
            });
            resolve(scoped.getReleaseReadView(selected!.releaseId));
          } catch (error) {
            reject(error);
          }
        };
        if (remaining === 0) {
          finishObjects();
          return;
        }
        for (const descriptor of selected.requiredObjects) {
          const request = objectStore.get(descriptor.sha256) as IDBRequest<ContentObjectRecord | undefined>;
          request.onerror = () => failRead('IndexedDB release object read', request);
          request.onsuccess = () => {
            if (request.result) records.push(cloneRecord(request.result));
            remaining -= 1;
            if (remaining === 0) finishObjects();
          };
        }
      };

      metadataRequest.onerror = () => failRead('IndexedDB metadata read', metadataRequest);
      metadataRequest.onsuccess = () => {
        activeReleaseId = metadataRequest.result || null;
        if (releaseId) {
          finishSelection();
          return;
        }
        if (!activeReleaseId) {
          selectionComplete = true;
          finishSelection();
          return;
        }
        const request = transactionStore.get(activeReleaseId) as IDBRequest<ReleaseTransactionRecord | undefined>;
        request.onerror = () => failRead('IndexedDB active release read', request);
        request.onsuccess = () => {
          selected = request.result ? cloneTransaction(request.result) : undefined;
          selectionComplete = true;
          finishSelection();
        };
      };

      if (releaseId) {
        const request = transactionStore.get(releaseId) as IDBRequest<ReleaseTransactionRecord | undefined>;
        request.onerror = () => failRead('IndexedDB release read', request);
        request.onsuccess = () => {
          selected = request.result ? cloneTransaction(request.result) : undefined;
          selectionComplete = true;
          finishSelection();
        };
      }
    });

    const [result] = await Promise.all([view, completion]);
    return result;
  }

  private async mutate<T>(
    mutation: (journal: MemoryReleaseContentJournal) => {
      result: T;
      changedReleases: string[];
      changedObjects?: ContentObjectRecord[];
      activeChanged?: boolean;
      rewriteReleaseContent?: string[];
    },
  ): Promise<T> {
    return withExclusiveLock(JOURNAL_LOCK_NAME, this.locks, async () => {
      const snapshot = await this.readSnapshot();
      const memory = MemoryReleaseContentJournal.fromSnapshot(snapshot, this.now);
      const change = mutation(memory);
      const next = memory.snapshot();
      const database = await this.databasePromise;
      const transaction = database.transaction(
        [
          'objects',
          'assetMappings',
          'plannedAssetMappings',
          'releaseObjects',
          'releaseRuntime',
          'releaseTransactions',
          'metadata',
        ],
        'readwrite',
      );
      const completion = transactionCompletion(transaction);
      const transactionStore = transaction.objectStore('releaseTransactions');
      const mappingStore = transaction.objectStore('assetMappings');
      const plannedMappingStore = transaction.objectStore('plannedAssetMappings');
      const releaseObjectStore = transaction.objectStore('releaseObjects');
      const runtimeStore = transaction.objectStore('releaseRuntime');
      const objectStore = transaction.objectStore('objects');
      const rewriteReleaseContent = new Set(change.rewriteReleaseContent || []);
      for (const object of change.changedObjects || []) objectStore.put(cloneRecord(object));
      for (const releaseId of new Set(change.changedReleases)) {
        const record = next.transactions.find((candidate) => candidate.releaseId === releaseId);
        if (!record) continue;
        transactionStore.put(cloneTransaction(record));
        runtimeStore.put(runtimeRecord(record));
        if (!rewriteReleaseContent.has(releaseId)) continue;
        const index = mappingStore.index('releaseId');
        const deleteRequest = index.openKeyCursor(IDBKeyRange.only(releaseId));
        deleteRequest.onsuccess = () => {
          const cursor = deleteRequest.result;
          if (!cursor) {
            for (const mapping of record.committedMappings) mappingStore.put(cloneMapping(mapping));
            return;
          }
          mappingStore.delete(cursor.primaryKey);
          cursor.continue();
        };
        const plannedIndex = plannedMappingStore.index('releaseId');
        const plannedDeleteRequest = plannedIndex.openKeyCursor(IDBKeyRange.only(releaseId));
        plannedDeleteRequest.onsuccess = () => {
          const cursor = plannedDeleteRequest.result;
          if (!cursor) {
            for (const mapping of record.plannedMappings) plannedMappingStore.put(cloneMapping(mapping));
            return;
          }
          plannedMappingStore.delete(cursor.primaryKey);
          cursor.continue();
        };
        const releaseObjectIndex = releaseObjectStore.index('releaseId');
        const releaseObjectDeleteRequest = releaseObjectIndex.openKeyCursor(IDBKeyRange.only(releaseId));
        releaseObjectDeleteRequest.onsuccess = () => {
          const cursor = releaseObjectDeleteRequest.result;
          if (!cursor) {
            for (const object of record.requiredObjects) {
              releaseObjectStore.put({ releaseId, ...cloneDescriptor(object) });
            }
            return;
          }
          releaseObjectStore.delete(cursor.primaryKey);
          cursor.continue();
        };
      }
      if (change.activeChanged) {
        if (next.activeReleaseId) {
          transaction.objectStore('metadata').put(next.activeReleaseId, ACTIVE_RELEASE_KEY);
        } else {
          transaction.objectStore('metadata').delete(ACTIVE_RELEASE_KEY);
        }
      }
      await completion;
      return change.result;
    });
  }

  async beginRelease(
    plan: ReleaseContentPlan,
    identity: Pick<ReleaseContentModel, 'manifestSha256' | 'storageSetSha256'>,
  ): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => {
      const result = journal.beginRelease(plan, identity);
      return {
        result,
        changedReleases: [plan.releaseId],
        rewriteReleaseContent: [plan.releaseId],
      };
    });
  }

  async recordVerifiedObject(
    releaseId: string,
    object: ContentObjectDescriptor,
    verifiedAt: number = this.now(),
  ): Promise<void> {
    await this.recordVerifiedObjects(releaseId, [object], verifiedAt);
  }

  async recordVerifiedObjects(
    releaseId: string,
    objects: readonly ContentObjectDescriptor[],
    verifiedAt: number = this.now(),
  ): Promise<void> {
    const unique = new Map<string, ContentObjectDescriptor>();
    for (const object of objects) {
      assertDigest(object.sha256);
      if (!Number.isSafeInteger(object.bytes) || object.bytes < 0) {
        throw new CanonicalResourceStoreError('invalid-content', `invalid verified object ${object.sha256}`);
      }
      const previous = unique.get(object.sha256);
      if (previous && previous.bytes !== object.bytes) {
        throw new CanonicalResourceStoreError('invalid-content', `conflicting verified object ${object.sha256}`);
      }
      unique.set(object.sha256, object);
    }
    if (!Number.isFinite(verifiedAt) || verifiedAt < 0) {
      throw new CanonicalResourceStoreError('invalid-content', 'invalid verification timestamp');
    }
    const descriptors = [...unique.values()];
    if (descriptors.length === 0) return;
    await withExclusiveLock(JOURNAL_LOCK_NAME, this.locks, async () => {
      for (let offset = 0; offset < descriptors.length; offset += 256) {
        await this.recordVerifiedObjectBatch(releaseId, descriptors.slice(offset, offset + 256), verifiedAt);
      }
    });
  }

  private async recordVerifiedObjectBatch(
    releaseId: string,
    objects: readonly ContentObjectDescriptor[],
    verifiedAt: number,
  ): Promise<void> {
    const database = await this.databasePromise;
    await new Promise<void>((resolve, reject) => {
      const transaction = database.transaction(['objects', 'releaseObjects', 'releaseRuntime'], 'readwrite');
      const objectStore = transaction.objectStore('objects');
      const runtimeRequest = transaction.objectStore('releaseRuntime').get(releaseId) as IDBRequest<
        CanonicalReleaseRuntimeRecord | undefined
      >;
      const expectedStore = transaction.objectStore('releaseObjects');
      const expectedRequests = objects.map(
        (object) =>
          expectedStore.get([releaseId, object.sha256]) as IDBRequest<
            (ContentObjectDescriptor & { releaseId: string }) | undefined
          >,
      );
      const existingRequests = objects.map(
        (object) => objectStore.get(object.sha256) as IDBRequest<ContentObjectRecord | undefined>,
      );
      let completedReads = 0;
      let validationError: Error | null = null;
      const finish = () => {
        completedReads += 1;
        if (completedReads !== 1 + objects.length * 2) return;
        const runtime = runtimeRequest.result;
        if (!runtime || runtime.state !== 'installing') {
          validationError = new CanonicalResourceStoreError(
            'invalid-state',
            `release ${releaseId} is not accepting object writes`,
          );
        }
        for (let index = 0; !validationError && index < objects.length; index += 1) {
          const object = objects[index];
          const expected = expectedRequests[index].result;
          const existing = existingRequests[index].result;
          if (!expected || expected.bytes !== object.bytes) {
            validationError = new CanonicalResourceStoreError(
              'invalid-content',
              `object ${object.sha256} is not planned for ${releaseId}`,
            );
          } else if (existing && existing.bytes !== object.bytes) {
            validationError = new CanonicalResourceStoreError(
              'invalid-content',
              `object ${object.sha256} conflicts with the ledger`,
            );
          }
        }
        if (validationError) {
          transaction.abort();
          return;
        }
        for (let index = 0; index < objects.length; index += 1) {
          objectStore.put({
            ...cloneDescriptor(objects[index]),
            verifiedAt: existingRequests[index].result?.verifiedAt || verifiedAt,
          });
        }
      };
      runtimeRequest.onsuccess = finish;
      for (const request of expectedRequests) request.onsuccess = finish;
      for (const request of existingRequests) request.onsuccess = finish;
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(validationError || transaction.error || new Error('IndexedDB write failed'));
      transaction.onabort = () => reject(validationError || transaction.error || new Error('IndexedDB write aborted'));
    });
  }

  async commitAssetMapping(releaseId: string, path: string): Promise<AssetContentMapping> {
    return (await this.commitAssetMappings(releaseId, [path]))[0];
  }

  async commitAssetMappings(releaseId: string, paths: readonly string[]): Promise<AssetContentMapping[]> {
    const uniquePaths = [...new Set(paths)];
    if (uniquePaths.length !== paths.length) {
      throw new CanonicalResourceStoreError('invalid-content', `duplicate asset mapping commit for ${releaseId}`);
    }
    if (uniquePaths.length === 0) return [];
    return withExclusiveLock(JOURNAL_LOCK_NAME, this.locks, async () => {
      const committed: AssetContentMapping[] = [];
      for (let offset = 0; offset < uniquePaths.length; offset += 256) {
        committed.push(...(await this.commitAssetMappingBatch(releaseId, uniquePaths.slice(offset, offset + 256))));
      }
      return committed;
    });
  }

  private async commitAssetMappingBatch(releaseId: string, paths: readonly string[]): Promise<AssetContentMapping[]> {
    const database = await this.databasePromise;
    return new Promise<AssetContentMapping[]>((resolve, reject) => {
      const transaction = database.transaction(
        ['assetMappings', 'objects', 'plannedAssetMappings', 'releaseObjects', 'releaseRuntime'],
        'readwrite',
      );
      const runtimeRequest = transaction.objectStore('releaseRuntime').get(releaseId) as IDBRequest<
        CanonicalReleaseRuntimeRecord | undefined
      >;
      const plannedStore = transaction.objectStore('plannedAssetMappings');
      const mappingRequests = paths.map(
        (path) => plannedStore.get([releaseId, path]) as IDBRequest<AssetContentMapping | undefined>,
      );
      let completedReads = 0;
      let committed: AssetContentMapping[] = [];
      let validationError: Error | null = null;
      const finishObjects = (
        mappings: AssetContentMapping[],
        digests: string[],
        expectedRequests: IDBRequest<(ContentObjectDescriptor & { releaseId: string }) | undefined>[],
        verifiedRequests: IDBRequest<ContentObjectRecord | undefined>[],
      ) => {
        const expectedObjects = new Map(
          expectedRequests
            .map((request) => request.result)
            .filter((object): object is ContentObjectDescriptor & { releaseId: string } => Boolean(object))
            .map((object) => [object.sha256, object]),
        );
        const verifiedObjects = new Map(
          verifiedRequests
            .map((request) => request.result)
            .filter((object): object is ContentObjectRecord => Boolean(object))
            .map((object) => [object.sha256, object]),
        );
        for (const digest of digests) {
          const expected = expectedObjects.get(digest);
          const verified = verifiedObjects.get(digest);
          if (!expected || !verified || expected.bytes !== verified.bytes) {
            validationError = new CanonicalResourceStoreError(
              'incomplete',
              `asset mapping references an unverified object ${digest}`,
            );
            transaction.abort();
            return;
          }
        }
        for (const mapping of mappings) {
          for (const span of mapping.spans) {
            const expected = expectedObjects.get(span.objectSha256);
            if (!expected || span.objectOffset + span.bytes > expected.bytes) {
              validationError = new CanonicalResourceStoreError(
                'invalid-content',
                `asset mapping span exceeds object ${span.objectSha256}`,
              );
              transaction.abort();
              return;
            }
          }
        }
        committed = mappings.map(cloneMapping);
        const committedStore = transaction.objectStore('assetMappings');
        for (const mapping of committed) committedStore.put(mapping);
      };
      const finishPlans = () => {
        completedReads += 1;
        if (completedReads !== mappingRequests.length + 1) return;
        const runtime = runtimeRequest.result;
        const mappings = mappingRequests.map((request) => request.result);
        if (!runtime || runtime.state !== 'installing') {
          validationError = new CanonicalResourceStoreError(
            'invalid-state',
            `release ${releaseId} is not accepting mapping writes`,
          );
        } else if (
          mappings.some(
            (mapping, index) => !mapping || mapping.releaseId !== releaseId || mapping.path !== paths[index],
          )
        ) {
          validationError = new CanonicalResourceStoreError(
            'invalid-content',
            `release ${releaseId} has an unknown planned mapping`,
          );
        }
        if (validationError) {
          transaction.abort();
          return;
        }
        const concreteMappings = mappings as AssetContentMapping[];
        const digests = [
          ...new Set(concreteMappings.flatMap((mapping) => mapping.spans.map((span) => span.objectSha256))),
        ];
        if (digests.length === 0) {
          finishObjects(concreteMappings, [], [], []);
          return;
        }
        const expectedStore = transaction.objectStore('releaseObjects');
        const verifiedStore = transaction.objectStore('objects');
        const expectedRequests = digests.map(
          (digest) =>
            expectedStore.get([releaseId, digest]) as IDBRequest<
              (ContentObjectDescriptor & { releaseId: string }) | undefined
            >,
        );
        const verifiedRequests = digests.map(
          (digest) => verifiedStore.get(digest) as IDBRequest<ContentObjectRecord | undefined>,
        );
        let objectReads = 0;
        const objectRead = () => {
          objectReads += 1;
          if (objectReads === digests.length * 2) {
            finishObjects(concreteMappings, digests, expectedRequests, verifiedRequests);
          }
        };
        for (const request of expectedRequests) request.onsuccess = objectRead;
        for (const request of verifiedRequests) request.onsuccess = objectRead;
      };
      runtimeRequest.onsuccess = finishPlans;
      for (const request of mappingRequests) request.onsuccess = finishPlans;
      transaction.oncomplete = () => resolve(committed.map(cloneMapping));
      transaction.onerror = () => reject(validationError || transaction.error || new Error('IndexedDB write failed'));
      transaction.onabort = () => reject(validationError || transaction.error || new Error('IndexedDB write aborted'));
    });
  }

  async prepareRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => ({
      result: journal.prepareRelease(releaseId),
      changedReleases: [releaseId],
    }));
  }

  async activateRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => {
      const previousReleaseId = journal.getActiveRelease()?.releaseId;
      const result = journal.activateRelease(releaseId);
      return {
        result,
        changedReleases: [releaseId, ...(previousReleaseId ? [previousReleaseId] : [])],
        activeChanged: true,
      };
    });
  }

  async reactivateRetainedRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => {
      const previousReleaseId = journal.getActiveRelease()?.releaseId;
      const result = journal.reactivateRetainedRelease(releaseId);
      return {
        result,
        changedReleases: [releaseId, ...(previousReleaseId ? [previousReleaseId] : [])],
        activeChanged: true,
      };
    });
  }

  async rollbackActivation(releaseId: string, failureCode: string): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => {
      const before = journal.getTransaction(releaseId);
      if (!before) {
        throw new CanonicalResourceStoreError('invalid-state', `release ${releaseId} is not journaled`);
      }
      const result = journal.rollbackActivation(releaseId, failureCode);
      return {
        result,
        changedReleases: [releaseId, ...(before.previousActiveReleaseId ? [before.previousActiveReleaseId] : [])],
        activeChanged: true,
      };
    });
  }

  async failRelease(releaseId: string, code: string): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => ({
      result: journal.failRelease(releaseId, code),
      changedReleases: [releaseId],
    }));
  }

  async resumeRelease(releaseId: string): Promise<ReleaseTransactionRecord> {
    return this.mutate((journal) => ({
      result: journal.resumeRelease(releaseId),
      changedReleases: [releaseId],
    }));
  }

  async getResumeState(releaseId: string): Promise<ReleaseResumeState> {
    const memory = MemoryReleaseContentJournal.fromSnapshot(await this.readSnapshot(), this.now);
    return memory.getResumeState(releaseId);
  }

  async getActiveRelease(): Promise<ReleaseTransactionRecord | null> {
    const database = await this.databasePromise;
    const transaction = database.transaction('metadata', 'readonly');
    const activeReleaseId = await requestResult(
      transaction.objectStore('metadata').get(ACTIVE_RELEASE_KEY) as IDBRequest<string | undefined>,
      'IndexedDB active release pointer read',
    );
    await transactionCompletion(transaction);
    if (!activeReleaseId) return null;
    const active = await this.getTransaction(activeReleaseId);
    if (!active || active.state !== 'active') {
      throw new CanonicalResourceStoreError('storage', 'active release pointer and transaction disagree');
    }
    return active;
  }

  async getTransaction(releaseId: string): Promise<ReleaseTransactionRecord | null> {
    const database = await this.databasePromise;
    const transaction = database.transaction(['assetMappings', 'releaseTransactions'], 'readonly');
    const [record, mappings] = await Promise.all([
      requestResult(
        transaction.objectStore('releaseTransactions').get(releaseId) as IDBRequest<
          ReleaseTransactionRecord | undefined
        >,
        'IndexedDB release transaction read',
      ),
      requestResult(
        transaction.objectStore('assetMappings').index('releaseId').getAll(releaseId) as IDBRequest<
          AssetContentMapping[]
        >,
        'IndexedDB release mapping read',
      ),
    ]);
    await transactionCompletion(transaction);
    return record
      ? {
          ...cloneTransaction(record),
          committedMappings: mappings.map(cloneMapping),
        }
      : null;
  }

  async listObjects(): Promise<ContentObjectRecord[]> {
    const database = await this.databasePromise;
    const transaction = database.transaction('objects', 'readonly');
    const objects = await requestResult(
      transaction.objectStore('objects').getAll() as IDBRequest<ContentObjectRecord[]>,
      'IndexedDB object inventory read',
    );
    await transactionCompletion(transaction);
    return objects.map(cloneRecord);
  }

  async getObject(sha256: string): Promise<ContentObjectRecord | null> {
    return (await this.getObjects([sha256]))[0] || null;
  }

  async getObjects(sha256s: readonly string[]): Promise<ContentObjectRecord[]> {
    const digests = [...new Set(sha256s)];
    for (const digest of digests) assertDigest(digest);
    if (digests.length === 0) return [];
    const database = await this.databasePromise;
    const transaction = database.transaction('objects', 'readonly');
    const store = transaction.objectStore('objects');
    const records = await Promise.all(
      digests.map((digest) =>
        requestResult(
          store.get(digest) as IDBRequest<ContentObjectRecord | undefined>,
          'IndexedDB content object read',
        ),
      ),
    );
    await transactionCompletion(transaction);
    return records.filter((record): record is ContentObjectRecord => Boolean(record)).map(cloneRecord);
  }

  async listAssetMappings(releaseId: string): Promise<AssetContentMapping[]> {
    const database = await this.databasePromise;
    const transaction = database.transaction('assetMappings', 'readonly');
    const mappings = await requestResult(
      transaction.objectStore('assetMappings').index('releaseId').getAll(releaseId) as IDBRequest<
        AssetContentMapping[]
      >,
      'IndexedDB release mapping read',
    );
    await transactionCompletion(transaction);
    return mappings.map(cloneMapping).sort((left, right) => left.path.localeCompare(right.path));
  }

  async getReleaseReadView(releaseId?: string): Promise<ReleaseContentReadView> {
    return this.readReleaseView(releaseId);
  }

  async getAssetReadView(releaseId: string, path: string): Promise<CanonicalAssetReadView> {
    const database = await this.databasePromise;
    return new Promise<CanonicalAssetReadView>((resolve, reject) => {
      const transaction = database.transaction(['assetMappings', 'metadata', 'objects', 'releaseRuntime'], 'readonly');
      const runtimeRequest = transaction.objectStore('releaseRuntime').get(releaseId) as IDBRequest<
        CanonicalReleaseRuntimeRecord | undefined
      >;
      const metadataRequest = transaction.objectStore('metadata').get(ACTIVE_RELEASE_KEY) as IDBRequest<
        string | undefined
      >;
      const mappingRequest = transaction.objectStore('assetMappings').get([releaseId, path]) as IDBRequest<
        AssetContentMapping | undefined
      >;
      let baseReads = 0;
      let result: CanonicalAssetReadView = { release: null, mapping: null, objects: [] };
      let validationError: Error | null = null;
      const finishBase = () => {
        baseReads += 1;
        if (baseReads !== 3) return;
        const runtime = runtimeRequest.result;
        const activeReleaseId = metadataRequest.result || null;
        const mapping = mappingRequest.result;
        if (
          runtime &&
          ((runtime.state === 'active' && activeReleaseId !== releaseId) ||
            (activeReleaseId === releaseId && runtime.state !== 'active'))
        ) {
          validationError = new CanonicalResourceStoreError(
            'storage',
            `active release pointer and runtime ${releaseId} disagree`,
          );
          transaction.abort();
          return;
        }
        if (!runtime || !mapping) {
          result = {
            release: runtime ? { ...runtime } : null,
            mapping: null,
            objects: [],
          };
          return;
        }
        if (mapping.releaseId !== releaseId || mapping.path !== path) {
          validationError = new CanonicalResourceStoreError('storage', `asset mapping identity mismatch`);
          transaction.abort();
          return;
        }
        const digests = [...new Set(mapping.spans.map((span) => span.objectSha256))];
        if (digests.length === 0) {
          result = {
            release: { ...runtime },
            mapping: cloneMapping(mapping),
            objects: [],
          };
          return;
        }
        const objectStore = transaction.objectStore('objects');
        const requests = digests.map(
          (digest) => objectStore.get(digest) as IDBRequest<ContentObjectRecord | undefined>,
        );
        let objectReads = 0;
        const finishObject = () => {
          objectReads += 1;
          if (objectReads !== requests.length) return;
          result = {
            release: { ...runtime },
            mapping: cloneMapping(mapping),
            objects: requests
              .map((request) => request.result)
              .filter((record): record is ContentObjectRecord => Boolean(record))
              .map(cloneRecord),
          };
        };
        for (const request of requests) request.onsuccess = finishObject;
      };
      runtimeRequest.onsuccess = finishBase;
      metadataRequest.onsuccess = finishBase;
      mappingRequest.onsuccess = finishBase;
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => reject(validationError || transaction.error || new Error('IndexedDB read failed'));
      transaction.onabort = () => reject(validationError || transaction.error || new Error('IndexedDB read aborted'));
    });
  }

  async getReleaseProbeView(releaseId: string): Promise<CanonicalReleaseProbeView> {
    const database = await this.databasePromise;
    return new Promise<CanonicalReleaseProbeView>((resolve, reject) => {
      const transaction = database.transaction(['assetMappings', 'metadata', 'objects', 'releaseRuntime'], 'readonly');
      const runtimeRequest = transaction.objectStore('releaseRuntime').get(releaseId) as IDBRequest<
        CanonicalReleaseRuntimeRecord | undefined
      >;
      const metadataRequest = transaction.objectStore('metadata').get(ACTIVE_RELEASE_KEY) as IDBRequest<
        string | undefined
      >;
      let runtimeDone = false;
      let metadataDone = false;
      let mappingDone = false;
      let objectDone = false;
      let mapping: AssetContentMapping | null = null;
      let object: ContentObjectRecord | null = null;
      let result: CanonicalReleaseProbeView = { release: null, mapping: null, object: null };
      let validationError: Error | null = null;
      const finish = () => {
        if (!runtimeDone || !metadataDone || !mappingDone || !objectDone) return;
        const runtime = runtimeRequest.result;
        const activeReleaseId = metadataRequest.result || null;
        if (
          runtime &&
          ((runtime.state === 'active' && activeReleaseId !== releaseId) ||
            (activeReleaseId === releaseId && runtime.state !== 'active'))
        ) {
          validationError = new CanonicalResourceStoreError(
            'storage',
            `active release pointer and runtime ${releaseId} disagree`,
          );
          transaction.abort();
          return;
        }
        result = {
          release: runtime ? { ...runtime } : null,
          mapping: mapping ? cloneMapping(mapping) : null,
          object: object ? cloneRecord(object) : null,
        };
      };
      runtimeRequest.onsuccess = () => {
        runtimeDone = true;
        const probePath = runtimeRequest.result?.probePath;
        if (!probePath) {
          mappingDone = true;
          objectDone = true;
          finish();
          return;
        }
        const mappingRequest = transaction.objectStore('assetMappings').get([releaseId, probePath]) as IDBRequest<
          AssetContentMapping | undefined
        >;
        mappingRequest.onsuccess = () => {
          mapping = mappingRequest.result ? cloneMapping(mappingRequest.result) : null;
          mappingDone = true;
          const digest = mapping?.spans[0]?.objectSha256;
          if (!digest) {
            objectDone = true;
            finish();
            return;
          }
          const objectRequest = transaction.objectStore('objects').get(digest) as IDBRequest<
            ContentObjectRecord | undefined
          >;
          objectRequest.onsuccess = () => {
            object = objectRequest.result ? cloneRecord(objectRequest.result) : null;
            objectDone = true;
            finish();
          };
        };
        finish();
      };
      metadataRequest.onsuccess = () => {
        metadataDone = true;
        finish();
      };
      transaction.oncomplete = () => resolve(result);
      transaction.onerror = () => reject(validationError || transaction.error || new Error('IndexedDB read failed'));
      transaction.onabort = () => reject(validationError || transaction.error || new Error('IndexedDB read aborted'));
    });
  }
}

export class CanonicalResourceStore {
  private readonly cacheStorage: CacheStorage;
  private readonly journal: CanonicalResourceJournal;
  private readonly fetchImplementation: typeof fetch;
  private readonly objectUrl: CanonicalResourceStoreOptions['objectUrl'];
  private readonly locks: LockManager | undefined;
  private readonly cacheKeyOrigin: string | URL;
  private readonly maxConcurrentDownloads: number;
  private readonly now: () => number;
  private readonly onObjectProgress?: CanonicalResourceStoreOptions['onObjectProgress'];
  private readonly onObjectVerified?: CanonicalResourceStoreOptions['onObjectVerified'];

  constructor(options: CanonicalResourceStoreOptions) {
    this.cacheStorage = options.cacheStorage;
    this.journal = options.journal;
    this.fetchImplementation = options.fetch;
    this.objectUrl = options.objectUrl;
    this.locks = options.locks ?? getDefaultLocks();
    this.cacheKeyOrigin = options.cacheKeyOrigin || DEFAULT_CACHE_KEY_ORIGIN;
    this.maxConcurrentDownloads = options.maxConcurrentDownloads ?? 4;
    if (
      !Number.isSafeInteger(this.maxConcurrentDownloads) ||
      this.maxConcurrentDownloads < 1 ||
      this.maxConcurrentDownloads > 16
    ) {
      throw new TypeError('maxConcurrentDownloads must be an integer from 1 to 16');
    }
    this.now = options.now || Date.now;
    this.onObjectProgress = options.onObjectProgress;
    this.onObjectVerified = options.onObjectVerified;
  }

  async prepareRelease(
    model: ReleaseContentModel,
    signal?: AbortSignal,
    packageTransport?: CanonicalPackageTransport,
    transferAttempt = 0,
  ): Promise<ReleaseTransactionRecord> {
    return withExclusiveLock(`${RELEASE_LOCK_PREFIX}${model.releaseId}`, this.locks, () =>
      this.prepareReleaseLocked(model, signal, packageTransport, transferAttempt),
    );
  }

  async installAndActivate(
    model: ReleaseContentModel,
    signal?: AbortSignal,
    packageTransport?: CanonicalPackageTransport,
    transferAttempt = 0,
  ): Promise<ReleaseTransactionRecord> {
    return withExclusiveLock(`${RELEASE_LOCK_PREFIX}${model.releaseId}`, this.locks, async () => {
      const prepared = await this.prepareReleaseLocked(model, signal, packageTransport, transferAttempt);
      if (prepared.state === 'active') return prepared;
      if (prepared.state === 'retained') {
        return this.journal.reactivateRetainedRelease(model.releaseId);
      }
      if (prepared.state !== 'prepared') {
        throw new CanonicalResourceStoreError(
          'invalid-state',
          `release ${model.releaseId} is not prepared for activation`,
        );
      }
      try {
        return await this.journal.activateRelease(model.releaseId);
      } catch (error) {
        const current = await this.journal.getTransaction(model.releaseId);
        if (current?.state === 'prepared') {
          try {
            await this.journal.failRelease(model.releaseId, failureCode(error));
          } catch {
            // Preserve the activation failure.
          }
        }
        throw error;
      }
    });
  }

  async rollbackActivation(releaseId: string, failureCode: string): Promise<ReleaseTransactionRecord> {
    return withExclusiveLock(`${RELEASE_LOCK_PREFIX}${releaseId}`, this.locks, () =>
      this.journal.rollbackActivation(releaseId, failureCode),
    );
  }

  /**
   * Current browser releases retain every canonical object. Until mutable
   * stable pointers and Piwork descriptors are available to the same trusted
   * transaction as leases and release records, deletion must fail closed.
   */
  planGarbageCollection(): ReleaseGcPlan {
    return conservativeNoDeleteReleaseGcPlan();
  }

  private async prepareReleaseLocked(
    model: ReleaseContentModel,
    signal?: AbortSignal,
    packageTransport?: CanonicalPackageTransport,
    transferAttempt = 0,
  ): Promise<ReleaseTransactionRecord> {
    let transaction = await this.journal.getTransaction(model.releaseId);
    try {
      if (transaction) {
        if (
          transaction.manifestSha256 !== model.manifestSha256 ||
          transaction.storageSetSha256 !== model.storageSetSha256
        ) {
          throw new CanonicalResourceStoreError(
            'invalid-content',
            `release ${model.releaseId} is already associated with different content`,
          );
        }
        if (transaction.state === 'failed') transaction = await this.journal.resumeRelease(model.releaseId);
      } else {
        const inventory = await this.listHealthyObjects();
        const active = await this.journal.getActiveRelease();
        const previousMappings = active ? await this.journal.listAssetMappings(active.releaseId) : [];
        const plan = planReleaseContent(model, inventory, previousMappings, {
          preferCanonicalForColdInstall: Boolean(packageTransport && !active),
        });
        transaction = await this.journal.beginRelease(plan, model);
      }

      const knownObjects = new Map(
        (await this.journal.getObjects(transaction.requiredObjects.map((object) => object.sha256))).map((object) => [
          object.sha256,
          object,
        ]),
      );
      const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
      const missingObjects: ContentObjectDescriptor[] = [];
      for (const object of transaction.requiredObjects) {
        if (signal?.aborted) throw new DOMException('The operation was aborted', 'AbortError');
        const record = knownObjects.get(object.sha256);
        if (!(await this.hasVerifiedObject(object, record, cache))) missingObjects.push(object);
      }
      const usePackageTransport =
        Boolean(packageTransport) &&
        transaction.previousActiveReleaseId === null &&
        transaction.plannedMappings.every((mapping) => mapping.representationKind !== 'package');
      if (usePackageTransport) {
        await this.installPackageObjects(
          transaction.releaseId,
          model,
          packageTransport!,
          {
            mappings: transaction.plannedMappings,
            missingObjects,
          },
          transaction.state === 'installing',
          signal,
          transferAttempt,
        );
      } else {
        await this.installObjects(
          transaction.releaseId,
          missingObjects,
          transaction.state === 'installing',
          signal,
          transferAttempt,
        );
      }

      transaction = (await this.journal.getTransaction(model.releaseId))!;
      if (transaction.state === 'installing') {
        const resume = await this.journal.getResumeState(model.releaseId);
        await this.journal.commitAssetMappings(model.releaseId, resume.pendingAssetPaths);
        transaction = await this.journal.prepareRelease(model.releaseId);
      }
      return transaction;
    } catch (error) {
      const current = await this.journal.getTransaction(model.releaseId);
      if (current && (current.state === 'installing' || current.state === 'prepared')) {
        try {
          await this.journal.failRelease(model.releaseId, failureCode(error));
        } catch {
          // Preserve the original installation failure.
        }
      }
      throw error;
    }
  }

  async checkHealth(options: { releaseId?: string; probe?: boolean } = {}): Promise<CanonicalResourceHealth> {
    const view = await this.journal.getReleaseReadView(options.releaseId);
    const transaction = view.transaction;
    if (!transaction) {
      return {
        releaseId: options.releaseId || null,
        state: 'missing',
        ready: false,
        missingObjects: [],
        probeSucceeded: false,
      };
    }
    const records = new Map(view.objects.map((record) => [record.sha256, record]));
    const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    const missing = Array.from<boolean>({ length: transaction.requiredObjects.length }).fill(false);
    let probeResponse: Response | undefined;
    let nextObjectIndex = 0;
    const workers = Array.from(
      { length: Math.min(HEALTH_CHECK_CONCURRENCY, transaction.requiredObjects.length) },
      async () => {
        while (true) {
          const index = nextObjectIndex;
          nextObjectIndex += 1;
          if (index >= transaction.requiredObjects.length) return;
          const object = transaction.requiredObjects[index];
          const record = records.get(object.sha256);
          if (!record || record.bytes !== object.bytes) {
            missing[index] = true;
            continue;
          }
          const response = await cache.match(canonicalContentCacheKey(object.sha256, this.cacheKeyOrigin));
          if (!response || !this.responseMatchesDescriptor(response, object)) {
            missing[index] = true;
            continue;
          }
          if (index === 0) probeResponse = response;
        }
      },
    );
    await Promise.all(workers);
    const missingObjects = transaction.requiredObjects.filter((_, index) => missing[index]).map(cloneDescriptor);
    const completeState = ['prepared', 'active', 'retained'].includes(transaction.state);
    const probeSucceeded =
      missingObjects.length === 0 &&
      (!options.probe || (await this.probeMatchedObject(transaction.requiredObjects[0], probeResponse)));
    return {
      releaseId: transaction.releaseId,
      state: transaction.state,
      ready: completeState && missingObjects.length === 0 && probeSucceeded,
      missingObjects,
      probeSucceeded,
    };
  }

  async verifyReleaseIntegrity(
    options: CanonicalResourceIntegrityVerificationOptions = {},
  ): Promise<CanonicalResourceIntegrityVerification> {
    const view = await this.journal.getReleaseReadView(options.releaseId);
    const transaction = view.transaction;
    if (!transaction) {
      return {
        releaseId: options.releaseId || null,
        state: 'missing',
        status: options.signal?.aborted ? 'aborted' : 'complete',
        ready: false,
        checkedObjects: 0,
        checkedBytes: 0,
        verifiedObjects: 0,
        verifiedBytes: 0,
        failures: [],
      };
    }

    const requiredByDigest = new Map(transaction.requiredObjects.map((object) => [object.sha256, object]));
    let selectedObjects = transaction.requiredObjects;
    if (options.objectSha256s) {
      const selectedDigests = new Set<string>();
      for (const digest of options.objectSha256s) {
        assertDigest(digest);
        if (!requiredByDigest.has(digest)) {
          throw new CanonicalResourceStoreError(
            'invalid-content',
            `object ${digest} is not required by release ${transaction.releaseId}`,
          );
        }
        selectedDigests.add(digest);
      }
      selectedObjects = transaction.requiredObjects.filter((object) => selectedDigests.has(object.sha256));
    }

    const records = new Map(view.objects.map((record) => [record.sha256, record]));
    const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    const outcomes: Array<CanonicalObjectIntegrityOutcome | undefined> = Array.from({
      length: selectedObjects.length,
    });
    let nextIndex = 0;
    let aborted = Boolean(options.signal?.aborted);
    await Promise.all(
      Array.from({ length: Math.min(DEEP_VERIFY_CONCURRENCY, selectedObjects.length) }, async () => {
        while (!aborted) {
          if (options.signal?.aborted) {
            aborted = true;
            return;
          }
          const index = nextIndex;
          nextIndex += 1;
          if (index >= selectedObjects.length) return;
          const object = selectedObjects[index];
          const record = records.get(object.sha256);
          if (!record || record.bytes !== object.bytes) {
            outcomes[index] = {
              status: 'failed',
              failure: {
                object: cloneDescriptor(object),
                code: 'missing',
                removed: false,
              },
            };
            continue;
          }
          const outcome = await this.verifyCachedObject(cache, object, options.signal);
          outcomes[index] = outcome;
          if (outcome.status === 'aborted') aborted = true;
        }
      }),
    );

    const failures: CanonicalResourceIntegrityFailure[] = [];
    let checkedObjects = 0;
    let checkedBytes = 0;
    let verifiedObjects = 0;
    let verifiedBytes = 0;
    for (const outcome of outcomes) {
      if (!outcome || outcome.status === 'aborted') continue;
      checkedObjects += 1;
      if (outcome.status === 'verified') {
        checkedBytes += outcome.bytes;
        verifiedObjects += 1;
        verifiedBytes += outcome.bytes;
      } else {
        if (outcome.failure.actualBytes !== undefined) checkedBytes += outcome.failure.actualBytes;
        failures.push({
          ...outcome.failure,
          object: cloneDescriptor(outcome.failure.object),
        });
      }
    }
    const completeState = ['prepared', 'active', 'retained'].includes(transaction.state);
    return {
      releaseId: transaction.releaseId,
      state: transaction.state,
      status: aborted || options.signal?.aborted ? 'aborted' : 'complete',
      ready: !aborted && !options.signal?.aborted && completeState && failures.length === 0,
      checkedObjects,
      checkedBytes,
      verifiedObjects,
      verifiedBytes,
      failures,
    };
  }

  async probeRelease(releaseId: string): Promise<CanonicalResourceProbe> {
    const view = await this.journal.getReleaseProbeView(releaseId);
    const runtime = view.release;
    const mapping = view.mapping;
    const object = view.object;
    if (!runtime) {
      return {
        releaseId,
        state: 'missing',
        ready: false,
        probeSucceeded: false,
        probePath: null,
        probeAssetBytes: null,
        probeAssetSha256: null,
      };
    }
    const identityMatches =
      Boolean(mapping) &&
      runtime.probePath === mapping!.path &&
      runtime.probeAssetBytes === mapping!.assetBytes &&
      runtime.probeAssetSha256 === mapping!.assetSha256;
    const firstSpan = mapping?.spans[0];
    const objectMatches =
      Boolean(firstSpan && object) &&
      firstSpan!.objectSha256 === object!.sha256 &&
      firstSpan!.objectOffset + firstSpan!.bytes <= object!.bytes;
    let probeSucceeded = false;
    if (identityMatches && objectMatches) {
      const response = await this.matchObject(object!, object!);
      probeSucceeded = await this.probeMatchedObject(object!, response);
    }
    const ready = ['active', 'retained'].includes(runtime.state) && identityMatches && objectMatches && probeSucceeded;
    return {
      releaseId: runtime.releaseId,
      state: runtime.state,
      ready,
      probeSucceeded,
      probePath: mapping?.path || null,
      probeAssetBytes: mapping?.assetBytes ?? null,
      probeAssetSha256: mapping?.assetSha256 || null,
    };
  }

  async matchObject(object: ContentObjectDescriptor, knownRecord?: ContentObjectRecord): Promise<Response | undefined> {
    const record =
      knownRecord || (await this.journal.listObjects()).find((candidate) => candidate.sha256 === object.sha256);
    if (!record || record.bytes !== object.bytes) return undefined;
    const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    const response = await cache.match(canonicalContentCacheKey(object.sha256, this.cacheKeyOrigin));
    return response && this.responseMatchesDescriptor(response, object) ? response : undefined;
  }

  private async listHealthyObjects(): Promise<ContentObjectRecord[]> {
    const records = await this.journal.listObjects();
    const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    const healthy = Array.from<boolean>({ length: records.length }).fill(false);
    let nextIndex = 0;
    await Promise.all(
      Array.from({ length: Math.min(HEALTH_CHECK_CONCURRENCY, records.length) }, async () => {
        while (true) {
          const index = nextIndex;
          nextIndex += 1;
          if (index >= records.length) return;
          healthy[index] = await this.hasVerifiedObject(records[index], records[index], cache);
        }
      }),
    );
    return records.filter((_, index) => healthy[index]).map(cloneRecord);
  }

  private async hasVerifiedObject(
    object: ContentObjectDescriptor,
    knownRecord?: ContentObjectRecord,
    knownCache?: Cache,
  ): Promise<boolean> {
    const record = knownRecord || (await this.journal.getObject(object.sha256));
    if (!record || record.bytes !== object.bytes) return false;
    const cache = knownCache || (await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME));
    const response = await cache.match(canonicalContentCacheKey(object.sha256, this.cacheKeyOrigin));
    return Boolean(response && this.responseMatchesDescriptor(response, object));
  }

  private responseMatchesDescriptor(response: Response, object: ContentObjectDescriptor): boolean {
    return (
      response.ok &&
      response.headers.get(CANONICAL_CONTENT_DIGEST_HEADER) === object.sha256 &&
      response.headers.get(CANONICAL_CONTENT_BYTES_HEADER) === String(object.bytes) &&
      response.headers.get('content-length') === String(object.bytes)
    );
  }

  private async verifyCachedObject(
    cache: Cache,
    object: ContentObjectDescriptor,
    signal?: AbortSignal,
  ): Promise<CanonicalObjectIntegrityOutcome> {
    const cacheKey = canonicalContentCacheKey(object.sha256, this.cacheKeyOrigin);
    const remove = async (): Promise<boolean> => {
      try {
        return await cache.delete(cacheKey);
      } catch {
        return false;
      }
    };
    try {
      return await withExclusiveLock(`${OBJECT_LOCK_PREFIX}${object.sha256}`, this.locks, async () => {
        if (signal?.aborted) return { status: 'aborted' };
        const response = await cache.match(cacheKey);
        if (signal?.aborted) {
          await response?.body?.cancel(signal.reason).catch(() => undefined);
          return { status: 'aborted' };
        }
        if (!response) {
          return {
            status: 'failed',
            failure: {
              object: cloneDescriptor(object),
              code: 'missing',
              removed: false,
            },
          };
        }
        if (!response.ok) {
          await response.body?.cancel().catch(() => undefined);
          return {
            status: 'failed',
            failure: {
              object: cloneDescriptor(object),
              code: 'read',
              removed: await remove(),
            },
          };
        }

        const digest = sha256.create();
        const reader = response.body?.getReader();
        let actualBytes = 0;
        let fullyRead = true;
        const abortReader = () => {
          void reader?.cancel(signal?.reason).catch(() => undefined);
        };
        signal?.addEventListener('abort', abortReader, { once: true });
        try {
          if (reader) {
            while (true) {
              if (signal?.aborted) return { status: 'aborted' };
              const next = await reader.read();
              if (signal?.aborted) return { status: 'aborted' };
              if (next.done) break;
              digest.update(next.value);
              actualBytes += next.value.byteLength;
              if (actualBytes > object.bytes) {
                fullyRead = false;
                await reader.cancel().catch(() => undefined);
                break;
              }
            }
          }
        } finally {
          signal?.removeEventListener('abort', abortReader);
          try {
            reader?.releaseLock();
          } catch {
            // The stream can already be detached after cancellation.
          }
        }
        if (signal?.aborted) return { status: 'aborted' };
        const actualSha256 = bytesToHex(digest.digest());
        if (!fullyRead || actualBytes !== object.bytes || actualSha256 !== object.sha256) {
          return {
            status: 'failed',
            failure: {
              object: cloneDescriptor(object),
              code: 'integrity',
              removed: await remove(),
              actualBytes,
              ...(fullyRead ? { actualSha256 } : {}),
            },
          };
        }
        return {
          status: 'verified',
          bytes: actualBytes,
        };
      });
    } catch {
      if (signal?.aborted) return { status: 'aborted' };
      return {
        status: 'failed',
        failure: {
          object: cloneDescriptor(object),
          code: 'read',
          removed: await remove(),
        },
      };
    }
  }

  private async probeMatchedObject(
    object: ContentObjectDescriptor | undefined,
    response: Response | undefined,
  ): Promise<boolean> {
    if (!object) return true;
    if (!response) return false;
    const reader = response.body?.getReader();
    if (!reader) return object.bytes === 0;
    try {
      const first = await reader.read();
      return object.bytes === 0 ? first.done : !first.done && first.value.byteLength > 0;
    } finally {
      // Chromium may keep the cancellation promise for a large persisted Cache
      // Storage response pending while it tears down the backing body. The
      // readiness probe has already consumed a real byte at this point, so its
      // result must not wait on best-effort stream cleanup.
      void reader
        .cancel()
        .catch(() => undefined)
        .finally(() => {
          try {
            reader.releaseLock();
          } catch {
            // The body owns the lock until cancellation finishes.
          }
        });
    }
  }

  private async installPackageObjects(
    releaseId: string,
    model: ReleaseContentModel,
    transport: CanonicalPackageTransport,
    plan: Pick<ReleaseContentPlan, 'mappings' | 'missingObjects'>,
    recordInJournal: boolean,
    signal?: AbortSignal,
    transferAttempt = 0,
  ): Promise<void> {
    if (plan.missingObjects.length === 0) return;
    const ingest = createCanonicalPackageIngestPlan(model, transport, plan);
    const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    let segmentIndex = 0;
    let position = ingest.segments[0]?.offset ?? transport.headerBytes;
    let current:
      | {
          descriptor: CanonicalPackageTransportSegment;
          reader: ReadableStreamDefaultReader<Uint8Array>;
          digest: ReturnType<typeof sha256.create>;
          networkBytes: number;
          consumedBytes: number;
          pending: Uint8Array | null;
          pendingOffset: number;
        }
      | undefined;

    const abortCurrent = async (reason?: unknown): Promise<void> => {
      const active = current;
      current = undefined;
      if (active) await active.reader.cancel(reason).catch(() => undefined);
    };
    const throwIfAborted = (): void => {
      if (signal?.aborted) throw new DOMException('The operation was aborted', 'AbortError');
    };
    const readWithAbort = async (
      reader: ReadableStreamDefaultReader<Uint8Array>,
    ): Promise<ReadableStreamReadResult<Uint8Array>> => {
      if (!signal) return reader.read();
      if (signal.aborted) {
        await reader.cancel(signal.reason).catch(() => undefined);
        throw new DOMException('The operation was aborted', 'AbortError');
      }
      return new Promise<ReadableStreamReadResult<Uint8Array>>((resolve, reject) => {
        let settled = false;
        let onAbort: () => void;
        const cleanup = () => signal.removeEventListener('abort', onAbort);
        onAbort = () => {
          if (settled) return;
          settled = true;
          cleanup();
          void reader.cancel(signal.reason).catch(() => undefined);
          reject(new DOMException('The operation was aborted', 'AbortError'));
        };
        signal.addEventListener('abort', onAbort, { once: true });
        reader.read().then(
          (result) => {
            if (settled) return;
            settled = true;
            cleanup();
            resolve(result);
          },
          (error: unknown) => {
            if (settled) return;
            settled = true;
            cleanup();
            reject(error);
          },
        );
      });
    };
    const openSegment = async (descriptor: CanonicalPackageTransportSegment): Promise<void> => {
      throwIfAborted();
      let response: Response;
      try {
        response = await this.fetchImplementation(
          transferInputForAttempt(this.objectUrl(releaseId, descriptor), transferAttempt),
          {
            cache: 'no-store',
            credentials: 'omit',
            signal,
          },
        );
      } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') throw error;
        throw new CanonicalResourceStoreError('network', `failed to download package segment ${descriptor.sha256}`, {
          cause: error,
        });
      }
      if (!response.ok || response.status !== 200) {
        await response.body?.cancel().catch(() => undefined);
        throw new CanonicalResourceStoreError(
          'http',
          `package segment ${descriptor.sha256} returned HTTP ${response.status}`,
        );
      }
      const declaredLength = response.headers.get('content-length');
      if (
        declaredLength !== null &&
        (!/^(?:0|[1-9][0-9]*)$/.test(declaredLength) || Number(declaredLength) !== descriptor.bytes)
      ) {
        await response.body?.cancel().catch(() => undefined);
        throw new CanonicalResourceStoreError(
          'integrity',
          `package segment ${descriptor.sha256} returned an unexpected Content-Length`,
        );
      }
      const source =
        response.body ||
        new ReadableStream<Uint8Array>({
          start(controller) {
            controller.close();
          },
        });
      current = {
        descriptor,
        reader: source.getReader(),
        digest: sha256.create(),
        networkBytes: 0,
        consumedBytes: 0,
        pending: null,
        pendingOffset: 0,
      };
      position = descriptor.offset;
    };
    const readCurrent = async (maximumBytes: number): Promise<Uint8Array> => {
      throwIfAborted();
      const active = current;
      if (!active || maximumBytes <= 0) return new Uint8Array();
      if (!active.pending || active.pendingOffset >= active.pending.byteLength) {
        const next = await readWithAbort(active.reader);
        throwIfAborted();
        if (next.done) {
          throw new CanonicalResourceStoreError(
            'integrity',
            `package segment ${active.descriptor.sha256} ended before its declared size`,
          );
        }
        active.pending = next.value;
        active.pendingOffset = 0;
        active.digest.update(next.value);
        active.networkBytes += next.value.byteLength;
        if (active.networkBytes > active.descriptor.bytes) {
          throw new CanonicalResourceStoreError(
            'integrity',
            `package segment ${active.descriptor.sha256} exceeded its declared size`,
          );
        }
        await this.onObjectProgress?.({
          releaseId,
          object: cloneDescriptor(active.descriptor),
          chunkBytes: next.value.byteLength,
          loadedBytes: active.networkBytes,
        });
      }
      const available = active.pending.byteLength - active.pendingOffset;
      const take = Math.min(maximumBytes, available);
      const result = active.pending.subarray(active.pendingOffset, active.pendingOffset + take);
      active.pendingOffset += take;
      active.consumedBytes += take;
      position += take;
      if (active.pendingOffset === active.pending.byteLength) {
        active.pending = null;
        active.pendingOffset = 0;
      }
      return result;
    };
    const finishCurrent = async (): Promise<void> => {
      const active = current;
      if (!active) return;
      while (active.consumedBytes < active.descriptor.bytes) {
        await readCurrent(active.descriptor.bytes - active.consumedBytes);
      }
      const overflow = await readWithAbort(active.reader);
      if (!overflow.done) {
        await active.reader.cancel().catch(() => undefined);
        current = undefined;
        throw new CanonicalResourceStoreError(
          'integrity',
          `package segment ${active.descriptor.sha256} exceeded its declared size`,
        );
      }
      const actualDigest = bytesToHex(active.digest.digest());
      current = undefined;
      if (active.networkBytes !== active.descriptor.bytes || actualDigest !== active.descriptor.sha256) {
        throw new CanonicalResourceStoreError(
          'integrity',
          `package segment ${active.descriptor.sha256} failed byte or SHA-256 verification`,
        );
      }
      segmentIndex += 1;
    };
    const advanceTo = async (target: number): Promise<void> => {
      if (!Number.isSafeInteger(target) || target < position) {
        throw new CanonicalResourceStoreError('invalid-content', 'package ingest moved backwards');
      }
      while (true) {
        throwIfAborted();
        if (current) {
          const end = current.descriptor.offset + current.descriptor.bytes;
          if (target <= end) {
            while (position < target) await readCurrent(target - position);
            if (position === end) await finishCurrent();
            return;
          }
          await finishCurrent();
          continue;
        }
        const descriptor = ingest.segments[segmentIndex];
        if (!descriptor) {
          throw new CanonicalResourceStoreError(
            'invalid-content',
            'package ingest target is not in a selected segment',
          );
        }
        if (target < descriptor.offset || target > descriptor.offset + descriptor.bytes) {
          if (target > descriptor.offset + descriptor.bytes) {
            await openSegment(descriptor);
            await finishCurrent();
            continue;
          }
          throw new CanonicalResourceStoreError('invalid-content', 'package ingest target precedes selected content');
        }
        await openSegment(descriptor);
      }
    };
    const readAssetBytes = async (maximumBytes: number): Promise<Uint8Array> => {
      throwIfAborted();
      if (!current) {
        const descriptor = ingest.segments[segmentIndex];
        if (!descriptor || descriptor.offset !== position) {
          throw new CanonicalResourceStoreError('invalid-content', 'package asset crosses an unavailable segment');
        }
        await openSegment(descriptor);
      }
      const active = current!;
      const availableInSegment = active.descriptor.offset + active.descriptor.bytes - position;
      if (availableInSegment === 0) {
        await finishCurrent();
        return readAssetBytes(maximumBytes);
      }
      return readCurrent(Math.min(maximumBytes, availableInSegment));
    };

    try {
      for (const unit of ingest.units) {
        const object = { sha256: unit.sha256, bytes: unit.bytes };
        assertDigest(object.sha256);
        if (object.bytes > 0) await advanceTo(unit.packageOffset);
        await withExclusiveLock(`${OBJECT_LOCK_PREFIX}${object.sha256}`, this.locks, async () => {
          if (await this.hasVerifiedObject(object, undefined, cache)) {
            let skipped = 0;
            while (skipped < object.bytes) {
              skipped += (await readAssetBytes(object.bytes - skipped)).byteLength;
            }
            return;
          }
          const cacheKey = canonicalContentCacheKey(object.sha256, this.cacheKeyOrigin);
          await cache.delete(cacheKey);
          const digest = sha256.create();
          let receivedBytes = 0;
          let streamFailure: unknown;
          const stream = new ReadableStream<Uint8Array>({
            pull: async (controller) => {
              try {
                if (receivedBytes === object.bytes) {
                  controller.close();
                  return;
                }
                const chunk = await readAssetBytes(object.bytes - receivedBytes);
                if (chunk.byteLength === 0) {
                  throw new CanonicalResourceStoreError(
                    'integrity',
                    `package object ${object.sha256} ended before its declared size`,
                  );
                }
                digest.update(chunk);
                receivedBytes += chunk.byteLength;
                controller.enqueue(chunk);
              } catch (error) {
                streamFailure = error;
                controller.error(error);
              }
            },
            cancel: (reason) => abortCurrent(reason),
          });
          const headers = new Headers({
            'cache-control': 'public, max-age=31536000, immutable',
            'content-length': String(object.bytes),
            'content-type': unit.mime || 'application/octet-stream',
            [CANONICAL_CONTENT_DIGEST_HEADER]: object.sha256,
            [CANONICAL_CONTENT_BYTES_HEADER]: String(object.bytes),
          });
          try {
            await cache.put(cacheKey, new Response(stream, { status: 200, headers }));
            const actualDigest = bytesToHex(digest.digest());
            if (receivedBytes !== object.bytes || actualDigest !== object.sha256) {
              await cache.delete(cacheKey);
              throw new CanonicalResourceStoreError(
                'integrity',
                `package object ${object.sha256} failed byte or SHA-256 verification`,
              );
            }
            if (recordInJournal) {
              await this.journal.recordVerifiedObject(releaseId, object, this.now());
            }
            await this.onObjectVerified?.({ releaseId, object: cloneDescriptor(object) });
          } catch (error) {
            await cache.delete(cacheKey).catch(() => undefined);
            // A response can be accepted with HTTP 200 and still fail while
            // its body is being consumed (for example an HTTP/2 stream reset
            // from the edge). Cache.put() reports that as a storage
            // NetworkError, which would otherwise disable the installer's
            // network retry path. Preserve integrity/abort failures, but
            // classify an underlying body-stream failure as retryable network.
            if (streamFailure !== undefined) {
              if (streamFailure instanceof CanonicalResourceStoreError) throw streamFailure;
              if (streamFailure instanceof DOMException && streamFailure.name === 'AbortError') {
                throw streamFailure;
              }
              throw new CanonicalResourceStoreError(
                'network',
                `package segment ${current?.descriptor.sha256 || 'unknown'} stream failed`,
                { cause: streamFailure },
              );
            }
            if (error instanceof CanonicalResourceStoreError) throw error;
            if (error instanceof DOMException && error.name === 'AbortError') throw error;
            throw new CanonicalResourceStoreError('storage', `failed to persist package object ${object.sha256}`, {
              cause: error,
            });
          }
        });
      }
      if (current) await finishCurrent();
      if (segmentIndex !== ingest.segments.length) {
        throw new CanonicalResourceStoreError(
          'invalid-content',
          'package ingest did not consume every selected segment',
        );
      }
      throwIfAborted();
    } catch (error) {
      await abortCurrent(error);
      throw error;
    }
  }

  private async installObject(
    releaseId: string,
    object: ContentObjectDescriptor,
    recordInJournal: boolean,
    signal?: AbortSignal,
    transferAttempt = 0,
  ): Promise<void> {
    assertDigest(object.sha256);
    if (!Number.isSafeInteger(object.bytes) || object.bytes < 0) {
      throw new CanonicalResourceStoreError('invalid-content', `invalid size for object ${object.sha256}`);
    }
    await withExclusiveLock(`${OBJECT_LOCK_PREFIX}${object.sha256}`, this.locks, async () => {
      if (await this.hasVerifiedObject(object)) return;
      const cache = await this.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
      const cacheKey = canonicalContentCacheKey(object.sha256, this.cacheKeyOrigin);
      await cache.delete(cacheKey);

      let response: Response;
      try {
        response = await this.fetchImplementation(
          transferInputForAttempt(this.objectUrl(releaseId, object), transferAttempt),
          {
            cache: 'no-store',
            credentials: 'omit',
            signal,
          },
        );
      } catch (error) {
        if (error instanceof DOMException && error.name === 'AbortError') throw error;
        throw new CanonicalResourceStoreError('network', `failed to download object ${object.sha256}`, {
          cause: error,
        });
      }
      if (!response.ok || response.status !== 200) {
        await response.body?.cancel().catch(() => undefined);
        throw new CanonicalResourceStoreError('http', `object ${object.sha256} returned HTTP ${response.status}`);
      }
      const declaredLength = response.headers.get('content-length');
      if (
        declaredLength !== null &&
        (!/^(?:0|[1-9][0-9]*)$/.test(declaredLength) || Number(declaredLength) !== object.bytes)
      ) {
        await response.body?.cancel().catch(() => undefined);
        throw new CanonicalResourceStoreError(
          'integrity',
          `object ${object.sha256} returned an unexpected Content-Length`,
        );
      }

      const digest = sha256.create();
      let receivedBytes = 0;
      let streamFailure: unknown;
      const source =
        response.body ||
        new ReadableStream<Uint8Array>({
          start(controller) {
            controller.close();
          },
        });
      const verifiedStream = source.pipeThrough(
        new TransformStream<Uint8Array, Uint8Array>({
          transform: async (chunk, controller) => {
            try {
              if (signal?.aborted) throw new DOMException('The operation was aborted', 'AbortError');
              digest.update(chunk);
              receivedBytes += chunk.byteLength;
              if (receivedBytes > object.bytes) {
                throw new CanonicalResourceStoreError(
                  'integrity',
                  `object ${object.sha256} exceeded its declared size`,
                );
              }
              await this.onObjectProgress?.({
                releaseId,
                object: cloneDescriptor(object),
                chunkBytes: chunk.byteLength,
                loadedBytes: receivedBytes,
              });
              controller.enqueue(chunk);
            } catch (error) {
              streamFailure = error;
              throw error;
            }
          },
        }),
      );
      const headers = new Headers();
      const contentType = response.headers.get('content-type');
      if (contentType) headers.set('content-type', contentType);
      headers.set('content-length', String(object.bytes));
      headers.set('cache-control', 'public, max-age=31536000, immutable');
      headers.set(CANONICAL_CONTENT_DIGEST_HEADER, object.sha256);
      headers.set(CANONICAL_CONTENT_BYTES_HEADER, String(object.bytes));

      try {
        await cache.put(
          cacheKey,
          new Response(verifiedStream, {
            status: 200,
            headers,
          }),
        );
        const actualDigest = bytesToHex(digest.digest());
        if (receivedBytes !== object.bytes || actualDigest !== object.sha256) {
          await cache.delete(cacheKey);
          throw new CanonicalResourceStoreError(
            'integrity',
            `object ${object.sha256} failed byte or SHA-256 verification`,
          );
        }
        if (recordInJournal) {
          await this.journal.recordVerifiedObject(releaseId, object, this.now());
        }
        await this.onObjectVerified?.({ releaseId, object: cloneDescriptor(object) });
      } catch (error) {
        await cache.delete(cacheKey).catch(() => undefined);
        if (streamFailure !== undefined) {
          if (streamFailure instanceof CanonicalResourceStoreError) throw streamFailure;
          if (streamFailure instanceof DOMException && streamFailure.name === 'AbortError') throw streamFailure;
          throw new CanonicalResourceStoreError('network', `object ${object.sha256} response stream failed`, {
            cause: streamFailure,
          });
        }
        if (error instanceof CanonicalResourceStoreError) throw error;
        if (error instanceof DOMException && error.name === 'AbortError') throw error;
        throw new CanonicalResourceStoreError('storage', `failed to persist object ${object.sha256}`, {
          cause: error,
        });
      }
    });
  }

  private async installObjects(
    releaseId: string,
    objects: ContentObjectDescriptor[],
    recordInJournal: boolean,
    signal?: AbortSignal,
    transferAttempt = 0,
  ): Promise<void> {
    if (objects.length === 0) return;
    const controller = new AbortController();
    const abortFromParent = () => controller.abort(signal?.reason);
    if (signal?.aborted) abortFromParent();
    else signal?.addEventListener('abort', abortFromParent, { once: true });

    let cursor = 0;
    let firstFailure: unknown;
    const worker = async () => {
      while (!controller.signal.aborted && firstFailure === undefined) {
        const index = cursor;
        cursor += 1;
        const object = objects[index];
        if (!object) return;
        try {
          await this.installObject(releaseId, object, recordInJournal, controller.signal, transferAttempt);
        } catch (error) {
          if (firstFailure === undefined) firstFailure = error;
          return;
        }
      }
    };

    try {
      await Promise.all(Array.from({ length: Math.min(this.maxConcurrentDownloads, objects.length) }, () => worker()));
    } finally {
      signal?.removeEventListener('abort', abortFromParent);
    }
    if (firstFailure !== undefined) throw firstFailure;
    if (signal?.aborted) throw new DOMException('The operation was aborted', 'AbortError');
  }
}
