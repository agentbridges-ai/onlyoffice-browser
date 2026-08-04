import { sha256 } from '@noble/hashes/sha2.js';

export type ResourceReadiness =
  | 'checking'
  | 'needs-download'
  | 'ready'
  | 'update-available'
  | 'updating'
  | 'paused'
  | 'repair-needed'
  | 'error';

export type ResourcePhase = 'idle' | 'planning' | 'downloading' | 'verifying' | 'activating' | 'repairing' | 'paused';

export type ResourceErrorCode =
  | 'offline'
  | 'network'
  | 'timeout'
  | 'integrity'
  | 'quota'
  | 'manifest'
  | 'incompatible'
  | 'storage'
  | 'aborted';

export type ResourceScope = 'recommended' | 'document' | 'all' | 'repair' | 'fonts';
export type ResourceProfile = 'base' | 'word' | 'cell' | 'slide' | 'fonts-basic' | 'fonts-office-compat';

export interface ReleaseAsset {
  path: string;
  bytes: number;
  mime: string;
  sha256: string;
  profile: ResourceProfile;
  chunk: string;
  packageOffset?: number;
  representations?: ReleaseAssetRepresentations;
}

export interface ReleaseWholeRepresentation {
  sha256: string;
  bytes: number;
}

export interface ReleaseFastCdcChunk {
  offset: number;
  bytes: number;
  sha256: string;
}

export interface ReleaseFastCdcRepresentation {
  algorithm: 'fastcdc-v2020';
  minBytes: 65_536;
  averageBytes: 262_144;
  maxBytes: 1_048_576;
  normalization: 1;
  seed: 0;
  chunks: ReleaseFastCdcChunk[];
}

export interface ReleaseAssetRepresentations {
  whole: ReleaseWholeRepresentation;
  fastcdc?: ReleaseFastCdcRepresentation;
}

export type ReleaseAssetV5 = ReleaseAsset & {
  packageOffset: number;
  representations: ReleaseAssetRepresentations;
};

export interface ReleaseChunk {
  id: string;
  profile: ResourceProfile;
  bytes: number;
  paths: string[];
}

export interface ReleasePackageSegment {
  id: string;
  offset: number;
  bytes: number;
  sha256: string;
}

export interface ReleasePackage {
  format: 'onlyoffice-pack-v1';
  path: 'office-resources.oobpack';
  bytes: number;
  sha256: string;
  headerBytes: number;
  segmentBytes: number;
  segments: ReleasePackageSegment[];
}

export interface ReleaseManifestV3 {
  version: 3 | 4 | 5;
  releaseId: string;
  packageVersion: string;
  hostBuildId: string;
  shellRevision: string;
  runtimeManifestSha256: string;
  fontManifestSha256: string;
  x2t: {
    version: string;
    commit: string;
    sha256: string;
  };
  profiles: Record<ResourceProfile, string[]>;
  chunks: ReleaseChunk[];
  package?: ReleasePackage;
  assets: ReleaseAsset[];
  fontFamilies?: Array<{ name: string; paths: string[] }>;
}

export type ReleaseManifestV4 = Omit<ReleaseManifestV3, 'version' | 'package'> & {
  version: 4;
  package: ReleasePackage;
};

export interface ReleaseContentProtocolV1 {
  version: 1;
  digest: 'sha256';
  cacheKeyFormat: 'canonical-sha256-v1';
  storageSetSha256: string;
  fastcdcPolicyId: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0';
}

export type ReleaseManifestV5 = Omit<ReleaseManifestV3, 'version' | 'package' | 'assets'> & {
  version: 5;
  package: ReleasePackage;
  contentProtocol: ReleaseContentProtocolV1;
  assets: ReleaseAssetV5[];
};

export interface ReleaseChannel {
  version: 1;
  releaseId: string;
  manifestUrl?: string;
  manifestSha256?: string;
}

export interface RequiredReleaseIdentity {
  releaseId: string;
  manifestSha256: string;
  packageVersion: string;
  hostBuildId: string;
}

export interface ResourcePlan {
  planId: string;
  releaseId: string;
  scope: ResourceScope;
  profiles: string[];
  totalBytes: number;
  downloadBytes: number;
  reusedBytes: number;
}

export interface FailedResource {
  path: string;
  code: ResourceErrorCode;
  attempts: number;
}

export interface ResourceInstallerSnapshot {
  installedRelease: string | null;
  targetRelease: string | null;
  availableRelease: string | null;
  availablePackageVersion: string | null;
  readiness: ResourceReadiness;
  phase: ResourcePhase;
  storageMode: 'cache-storage' | 'http-cache';
  currentChunk: string | null;
  currentChunkIndex: number;
  currentChunkCount: number;
  downloadedBytes: number;
  downloadBytes: number;
  verifiedBytes: number;
  verifyBytes: number;
  bytesPerSecond: number;
  failedResources: FailedResource[];
  canPause: boolean;
  canResume: boolean;
  canRetry: boolean;
  errorCode: ResourceErrorCode | null;
  installedProfiles: ResourceProfile[];
}

export type ResourcePlanRequest = {
  scope: ResourceScope;
  documentType?: 'word' | 'cell' | 'slide';
  profiles?: ResourceProfile[];
};

export interface OfficeRuntimeResourceInstaller {
  plan(request: ResourcePlanRequest): Promise<ResourcePlan>;
  apply(plan: ResourcePlan): Promise<void>;
  checkForUpdates(): Promise<void>;
  checkHealth(options?: { deep?: boolean }): Promise<void>;
  repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void>;
  pause(): void;
  resume(): Promise<void>;
  cancel(): void;
  getInstallerSnapshot(): ResourceInstallerSnapshot;
  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void;
  getInstalledPaths(): string[];
}

export type JournalAsset = {
  releaseId: string;
  path: string;
  sha256: string;
  bytes: number;
  verifiedAt: number;
};

export type JournalRelease = {
  releaseId: string;
  installedProfiles: ResourceProfile[];
  activatedAt: number;
  packageSha256?: string;
  packageBytes?: number;
};

export type JournalSegment = {
  sha256: string;
  bytes: number;
  verifiedAt: number;
};

export interface InstallationJournal {
  listAssets(releaseId: string): Promise<JournalAsset[]>;
  putAsset(asset: JournalAsset): Promise<void>;
  deleteAsset(releaseId: string, path: string): Promise<void>;
  getActiveRelease(): Promise<JournalRelease | null>;
  activateRelease(release: JournalRelease): Promise<void>;
  listSegments(): Promise<JournalSegment[]>;
  putSegment(segment: JournalSegment): Promise<void>;
  deleteSegment(sha256: string): Promise<void>;
}

const DB_NAME = 'onlyoffice-browser-resources-v3';
const DB_VERSION = 2;
const RELEASE_SEGMENT_CACHE_NAME = 'onlyoffice-release-segments-v1';

export class IndexedDbInstallationJournal implements InstallationJournal {
  private readonly dbPromise: Promise<IDBDatabase>;

  constructor(indexedDb: IDBFactory = indexedDB) {
    this.dbPromise = new Promise((resolve, reject) => {
      const request = indexedDb.open(DB_NAME, DB_VERSION);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains('assets')) {
          const assets = db.createObjectStore('assets', { keyPath: ['releaseId', 'path'] });
          assets.createIndex('releaseId', 'releaseId');
        }
        if (!db.objectStoreNames.contains('metadata')) db.createObjectStore('metadata');
        if (!db.objectStoreNames.contains('segments')) db.createObjectStore('segments', { keyPath: 'sha256' });
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error('IndexedDB open failed'));
    });
  }

  async listAssets(releaseId: string): Promise<JournalAsset[]> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const request = db.transaction('assets').objectStore('assets').index('releaseId').getAll(releaseId);
      request.onsuccess = () => resolve(request.result as JournalAsset[]);
      request.onerror = () => reject(request.error || new Error('IndexedDB read failed'));
    });
  }

  async putAsset(asset: JournalAsset): Promise<void> {
    const db = await this.dbPromise;
    await transactionDone(db, 'assets', 'readwrite', (store) => store.put(asset));
  }

  async deleteAsset(releaseId: string, path: string): Promise<void> {
    const db = await this.dbPromise;
    await transactionDone(db, 'assets', 'readwrite', (store) => store.delete([releaseId, path]));
  }

  async getActiveRelease(): Promise<JournalRelease | null> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const request = db.transaction('metadata').objectStore('metadata').get('active-release');
      request.onsuccess = () => resolve((request.result as JournalRelease | undefined) || null);
      request.onerror = () => reject(request.error || new Error('IndexedDB read failed'));
    });
  }

  async activateRelease(release: JournalRelease): Promise<void> {
    const db = await this.dbPromise;
    await transactionDone(db, 'metadata', 'readwrite', (store) => store.put(release, 'active-release'));
  }

  async listSegments(): Promise<JournalSegment[]> {
    const db = await this.dbPromise;
    return new Promise((resolve, reject) => {
      const request = db.transaction('segments').objectStore('segments').getAll();
      request.onsuccess = () => resolve(request.result as JournalSegment[]);
      request.onerror = () => reject(request.error || new Error('IndexedDB read failed'));
    });
  }

  async putSegment(segment: JournalSegment): Promise<void> {
    const db = await this.dbPromise;
    await transactionDone(db, 'segments', 'readwrite', (store) => store.put(segment));
  }

  async deleteSegment(sha256: string): Promise<void> {
    const db = await this.dbPromise;
    await transactionDone(db, 'segments', 'readwrite', (store) => store.delete(sha256));
  }
}

function transactionDone(
  db: IDBDatabase,
  storeName: string,
  mode: IDBTransactionMode,
  mutate: (store: IDBObjectStore) => IDBRequest,
): Promise<void> {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, mode);
    mutate(transaction.objectStore(storeName));
    transaction.oncomplete = () => resolve();
    transaction.onerror = () => reject(transaction.error || new Error('IndexedDB transaction failed'));
    transaction.onabort = () => reject(transaction.error || new Error('IndexedDB transaction aborted'));
  });
}

export class MemoryInstallationJournal implements InstallationJournal {
  private readonly assets = new Map<string, JournalAsset>();
  private readonly segments = new Map<string, JournalSegment>();
  private active: JournalRelease | null = null;

  async listAssets(releaseId: string): Promise<JournalAsset[]> {
    return [...this.assets.values()].filter((asset) => asset.releaseId === releaseId);
  }

  async putAsset(asset: JournalAsset): Promise<void> {
    this.assets.set(`${asset.releaseId}\0${asset.path}`, { ...asset });
  }

  async deleteAsset(releaseId: string, path: string): Promise<void> {
    this.assets.delete(`${releaseId}\0${path}`);
  }

  async getActiveRelease(): Promise<JournalRelease | null> {
    return this.active ? { ...this.active, installedProfiles: [...this.active.installedProfiles] } : null;
  }

  async activateRelease(release: JournalRelease): Promise<void> {
    this.active = { ...release, installedProfiles: [...release.installedProfiles] };
  }

  async listSegments(): Promise<JournalSegment[]> {
    return [...this.segments.values()].map((segment) => ({ ...segment }));
  }

  async putSegment(segment: JournalSegment): Promise<void> {
    this.segments.set(segment.sha256, { ...segment });
  }

  async deleteSegment(sha256: string): Promise<void> {
    this.segments.delete(sha256);
  }
}

function isSafePath(value: unknown): value is string {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.startsWith('/') ||
    value.includes('\\') ||
    value.includes('?') ||
    value.includes('#') ||
    Array.from(value).some((character) => {
      const code = character.charCodeAt(0);
      return code <= 0x1f || code === 0x7f;
    })
  ) {
    return false;
  }
  let decoded: string;
  try {
    decoded = decodeURIComponent(value);
  } catch {
    return false;
  }
  if (decoded.startsWith('/') || decoded.includes('\\')) return false;
  const segments = decoded.split('/');
  return segments.every((segment) => segment.length > 0 && segment !== '.' && segment !== '..');
}

function isDigest(value: unknown): value is string {
  return typeof value === 'string' && /^[a-f0-9]{64}$/.test(value);
}

export function parseRequiredReleaseIdentity(value: unknown): RequiredReleaseIdentity | null {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return null;
  const prototype = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== null) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    keys.length !== 4 ||
    !Object.hasOwn(record, 'releaseId') ||
    !Object.hasOwn(record, 'manifestSha256') ||
    !Object.hasOwn(record, 'packageVersion') ||
    !Object.hasOwn(record, 'hostBuildId') ||
    typeof record.releaseId !== 'string' ||
    !/^[a-zA-Z0-9._+-]{1,128}$/.test(record.releaseId) ||
    !isDigest(record.manifestSha256) ||
    typeof record.packageVersion !== 'string' ||
    !/^[a-zA-Z0-9][a-zA-Z0-9._+-]{0,127}$/.test(record.packageVersion) ||
    typeof record.hostBuildId !== 'string' ||
    !/^[a-zA-Z0-9][a-zA-Z0-9._+-]{0,127}$/.test(record.hostBuildId)
  ) {
    return null;
  }
  return {
    releaseId: record.releaseId,
    manifestSha256: record.manifestSha256,
    packageVersion: record.packageVersion,
    hostBuildId: record.hostBuildId,
  };
}

export function requiredReleaseIdentitiesEqual(left: RequiredReleaseIdentity, right: RequiredReleaseIdentity): boolean {
  return (
    left.releaseId === right.releaseId &&
    left.manifestSha256 === right.manifestSha256 &&
    left.packageVersion === right.packageVersion &&
    left.hostBuildId === right.hostBuildId
  );
}

function isProfile(value: unknown): value is ResourceProfile {
  return ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'].includes(String(value));
}

function digestBytes(value: string | Uint8Array): string {
  const input = typeof value === 'string' ? new TextEncoder().encode(value) : value;
  return Array.from(sha256(input), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function readManifestResponse(response: Response): Promise<{
  manifest: ReleaseManifestV3;
  manifestSha256: string;
}> {
  const manifestBytes = new Uint8Array(await response.arrayBuffer());
  const manifestSha256 = digestBytes(manifestBytes);
  let manifestValue: unknown;
  try {
    const manifestText = new TextDecoder('utf-8', { fatal: true }).decode(manifestBytes);
    manifestValue = JSON.parse(manifestText);
  } catch {
    throw new ResourceInstallerError('manifest');
  }
  return {
    manifest: parseReleaseManifest(manifestValue),
    manifestSha256,
  };
}

function storageSetDescription(pack: ReleasePackage, assets: ReleaseAssetV5[]) {
  return {
    version: 1,
    packageSegments: pack.segments.map((segment) => ({
      offset: segment.offset,
      bytes: segment.bytes,
      sha256: segment.sha256,
    })),
    assets: [...assets]
      .sort((left, right) => left.path.localeCompare(right.path))
      .map((asset) => ({
        path: asset.path,
        bytes: asset.bytes,
        sha256: asset.sha256,
        whole: {
          bytes: asset.representations.whole.bytes,
          sha256: asset.representations.whole.sha256,
        },
        ...(asset.representations.fastcdc
          ? {
              fastcdc: {
                algorithm: asset.representations.fastcdc.algorithm,
                minBytes: asset.representations.fastcdc.minBytes,
                averageBytes: asset.representations.fastcdc.averageBytes,
                maxBytes: asset.representations.fastcdc.maxBytes,
                normalization: asset.representations.fastcdc.normalization,
                seed: asset.representations.fastcdc.seed,
                chunks: asset.representations.fastcdc.chunks.map((chunk) => ({
                  offset: chunk.offset,
                  bytes: chunk.bytes,
                  sha256: chunk.sha256,
                })),
              },
            }
          : {}),
      })),
  };
}

export function computeStorageSetSha256(pack: ReleasePackage, assets: ReleaseAssetV5[]): string {
  return digestBytes(JSON.stringify(storageSetDescription(pack, assets)));
}

export function parseReleaseManifest(value: unknown): ReleaseManifestV3 {
  const manifest = value as Partial<ReleaseManifestV3>;
  if (
    (manifest.version !== 3 && manifest.version !== 4 && manifest.version !== 5) ||
    typeof manifest.releaseId !== 'string' ||
    !/^[a-zA-Z0-9._+-]{1,128}$/.test(manifest.releaseId) ||
    typeof manifest.packageVersion !== 'string' ||
    typeof manifest.hostBuildId !== 'string' ||
    typeof manifest.shellRevision !== 'string' ||
    !isDigest(manifest.runtimeManifestSha256) ||
    !isDigest(manifest.fontManifestSha256) ||
    !manifest.x2t ||
    typeof manifest.x2t.version !== 'string' ||
    typeof manifest.x2t.commit !== 'string' ||
    !isDigest(manifest.x2t.sha256) ||
    !Array.isArray(manifest.assets) ||
    !Array.isArray(manifest.chunks) ||
    !manifest.profiles
  ) {
    throw new ResourceInstallerError('manifest');
  }
  const assets = manifest.assets.filter(
    (asset): asset is ReleaseAsset =>
      !!asset &&
      isSafePath(asset.path) &&
      Number.isSafeInteger(asset.bytes) &&
      asset.bytes >= 0 &&
      typeof asset.mime === 'string' &&
      isDigest(asset.sha256) &&
      isProfile(asset.profile) &&
      typeof asset.chunk === 'string' &&
      asset.chunk.length > 0 &&
      (manifest.version === 3 || (Number.isSafeInteger(asset.packageOffset) && Number(asset.packageOffset) >= 0)),
  );
  if (assets.length !== manifest.assets.length || new Set(assets.map((asset) => asset.path)).size !== assets.length) {
    throw new ResourceInstallerError('manifest');
  }
  const assetPaths = new Set(assets.map((asset) => asset.path));
  for (const profile of Object.keys(manifest.profiles)) {
    if (!isProfile(profile) || !Array.isArray(manifest.profiles[profile])) throw new ResourceInstallerError('manifest');
    if (!manifest.profiles[profile].every((path) => assetPaths.has(path))) throw new ResourceInstallerError('manifest');
  }
  if (manifest.version === 4 || manifest.version === 5) {
    const pack = manifest.package;
    if (
      !pack ||
      pack.format !== 'onlyoffice-pack-v1' ||
      pack.path !== 'office-resources.oobpack' ||
      !Number.isSafeInteger(pack.bytes) ||
      pack.bytes <= 0 ||
      !isDigest(pack.sha256) ||
      !Number.isSafeInteger(pack.headerBytes) ||
      pack.headerBytes <= 12 ||
      !Number.isSafeInteger(pack.segmentBytes) ||
      pack.segmentBytes <= 0 ||
      !Array.isArray(pack.segments) ||
      pack.segments.length === 0
    ) {
      throw new ResourceInstallerError('manifest');
    }
    let expectedOffset = 0;
    for (const segment of pack.segments) {
      if (
        !segment ||
        !isDigest(segment.id) ||
        segment.id !== segment.sha256 ||
        segment.offset !== expectedOffset ||
        !Number.isSafeInteger(segment.bytes) ||
        segment.bytes <= 0 ||
        !isDigest(segment.sha256)
      ) {
        throw new ResourceInstallerError('manifest');
      }
      expectedOffset += segment.bytes;
    }
    if (expectedOffset !== pack.bytes) throw new ResourceInstallerError('manifest');
    for (const asset of assets) {
      if (
        !Number.isSafeInteger(asset.packageOffset) ||
        Number(asset.packageOffset) < pack.headerBytes ||
        Number(asset.packageOffset) + asset.bytes > pack.bytes
      ) {
        throw new ResourceInstallerError('manifest');
      }
    }
  }
  if (manifest.version === 5) {
    const protocol = (manifest as Partial<ReleaseManifestV5>).contentProtocol;
    if (
      !protocol ||
      protocol.version !== 1 ||
      protocol.digest !== 'sha256' ||
      protocol.cacheKeyFormat !== 'canonical-sha256-v1' ||
      !isDigest(protocol.storageSetSha256) ||
      protocol.fastcdcPolicyId !== 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0'
    ) {
      throw new ResourceInstallerError('manifest');
    }
    const objectSizes = new Map<string, number>();
    const recordObject = (digest: string, bytes: number) => {
      const existing = objectSizes.get(digest);
      if (existing !== undefined && existing !== bytes) throw new ResourceInstallerError('manifest');
      objectSizes.set(digest, bytes);
    };
    for (const segment of manifest.package!.segments) recordObject(segment.sha256, segment.bytes);
    for (const asset of assets) {
      const representations = asset.representations;
      if (
        !representations ||
        !representations.whole ||
        representations.whole.sha256 !== asset.sha256 ||
        representations.whole.bytes !== asset.bytes
      ) {
        throw new ResourceInstallerError('manifest');
      }
      recordObject(representations.whole.sha256, representations.whole.bytes);
      const fastcdc = representations.fastcdc;
      if (!fastcdc) continue;
      if (
        fastcdc.algorithm !== 'fastcdc-v2020' ||
        fastcdc.minBytes !== 65_536 ||
        fastcdc.averageBytes !== 262_144 ||
        fastcdc.maxBytes !== 1_048_576 ||
        fastcdc.normalization !== 1 ||
        fastcdc.seed !== 0 ||
        !Array.isArray(fastcdc.chunks) ||
        fastcdc.chunks.length === 0
      ) {
        throw new ResourceInstallerError('manifest');
      }
      let expectedOffset = 0;
      for (const chunk of fastcdc.chunks) {
        if (
          !chunk ||
          chunk.offset !== expectedOffset ||
          !Number.isSafeInteger(chunk.bytes) ||
          chunk.bytes <= 0 ||
          !isDigest(chunk.sha256)
        ) {
          throw new ResourceInstallerError('manifest');
        }
        expectedOffset += chunk.bytes;
        if (expectedOffset > asset.bytes) throw new ResourceInstallerError('manifest');
        recordObject(chunk.sha256, chunk.bytes);
      }
      if (expectedOffset !== asset.bytes) throw new ResourceInstallerError('manifest');
    }
    if (computeStorageSetSha256(manifest.package!, assets as ReleaseAssetV5[]) !== protocol.storageSetSha256) {
      throw new ResourceInstallerError('manifest');
    }
  }
  return manifest as ReleaseManifestV3;
}

export class ReleaseRepository {
  private readonly baseUrl: URL;
  private readonly fetchImpl: typeof fetch;

  constructor(assetBaseUrl: string | URL, fetchImpl: typeof fetch) {
    this.baseUrl = new URL('/', assetBaseUrl);
    this.fetchImpl = fetchImpl;
  }

  async current(): Promise<{ channel: ReleaseChannel; manifest: ReleaseManifestV3 }> {
    const current = await this.readCurrent('channels/stable.json', false);
    return { channel: current.channel, manifest: current.manifest };
  }

  async currentV5(cached?: { manifest: ReleaseManifestV5; manifestSha256: string }): Promise<{
    channel: ReleaseChannel & { manifestSha256: string };
    manifest: ReleaseManifestV5;
    manifestSha256: string;
  }> {
    if (cached?.manifest.version === 5 && isDigest(cached.manifestSha256) && cached.manifest.releaseId) {
      const channel = await this.readChannel('channels/stable-v5.json', true);
      if (channel.releaseId === cached.manifest.releaseId && channel.manifestSha256 === cached.manifestSha256) {
        return {
          channel: channel as ReleaseChannel & { manifestSha256: string },
          manifest: cached.manifest,
          manifestSha256: cached.manifestSha256,
        };
      }
    }
    const current = await this.readCurrent('channels/stable-v5.json', true);
    if (current.manifest.version !== 5) throw new ResourceInstallerError('incompatible');
    return {
      channel: current.channel as ReleaseChannel & { manifestSha256: string },
      manifest: current.manifest as ReleaseManifestV5,
      manifestSha256: current.manifestSha256,
    };
  }

  async releaseV5(
    releaseId: string,
    expectedManifestSha256: string,
  ): Promise<{ manifest: ReleaseManifestV5; manifestSha256: string }> {
    if (!/^[a-zA-Z0-9._+-]{1,128}$/.test(releaseId) || !isDigest(expectedManifestSha256)) {
      throw new ResourceInstallerError('manifest');
    }
    const manifestUrl = new URL(`releases/${encodeURIComponent(releaseId)}/manifest.json`, this.baseUrl);
    const response = await this.fetchImpl(manifestUrl, {
      cache: 'no-store',
      credentials: 'omit',
      mode: typeof location !== 'undefined' && manifestUrl.origin === location.origin ? 'same-origin' : 'cors',
    });
    if (!response.ok) throw networkError(response.status);
    const { manifest, manifestSha256 } = await readManifestResponse(response);
    if (manifestSha256 !== expectedManifestSha256) throw new ResourceInstallerError('manifest');
    if (manifest.version !== 5 || manifest.releaseId !== releaseId) {
      throw new ResourceInstallerError('incompatible');
    }
    return { manifest: manifest as ReleaseManifestV5, manifestSha256 };
  }

  private async readCurrent(
    channelPath: 'channels/stable.json' | 'channels/stable-v5.json',
    requireManifestSha256: boolean,
  ): Promise<{ channel: ReleaseChannel; manifest: ReleaseManifestV3; manifestSha256: string }> {
    const channel = await this.readChannel(channelPath, requireManifestSha256);
    const channelUrl = new URL(channelPath, this.baseUrl);
    const manifestUrl = channel.manifestUrl
      ? new URL(channel.manifestUrl, channelUrl)
      : new URL(`releases/${encodeURIComponent(channel.releaseId)}/manifest.json`, this.baseUrl);
    const response = await this.fetchImpl(manifestUrl, {
      cache: 'no-store',
      credentials: 'omit',
      mode: typeof location !== 'undefined' && manifestUrl.origin === location.origin ? 'same-origin' : 'cors',
    });
    if (!response.ok) throw networkError(response.status);
    const { manifest, manifestSha256 } = await readManifestResponse(response);
    if (channel.manifestSha256 && manifestSha256 !== channel.manifestSha256) {
      throw new ResourceInstallerError('manifest');
    }
    if (manifest.releaseId !== channel.releaseId) throw new ResourceInstallerError('manifest');
    return { channel, manifest, manifestSha256 };
  }

  private async readChannel(
    channelPath: 'channels/stable.json' | 'channels/stable-v5.json',
    requireManifestSha256: boolean,
  ): Promise<ReleaseChannel> {
    const channelUrl = new URL(channelPath, this.baseUrl);
    const channelResponse = await this.fetchImpl(channelUrl, {
      cache: 'no-store',
      credentials: 'omit',
      mode: typeof location !== 'undefined' && channelUrl.origin === location.origin ? 'same-origin' : 'cors',
    });
    if (!channelResponse.ok) throw networkError(channelResponse.status);
    const channel = (await channelResponse.json()) as Partial<ReleaseChannel>;
    if (
      channel.version !== 1 ||
      typeof channel.releaseId !== 'string' ||
      (requireManifestSha256 && !isDigest(channel.manifestSha256)) ||
      (channel.manifestSha256 !== undefined && !isDigest(channel.manifestSha256))
    ) {
      throw new ResourceInstallerError('manifest');
    }
    return channel as ReleaseChannel;
  }

  assetUrl(releaseId: string, path: string): URL {
    if (!isSafePath(path)) throw new ResourceInstallerError('manifest');
    return new URL(`r/${encodeURIComponent(releaseId)}/${path}`, this.baseUrl);
  }

  packageUrl(releaseId: string): URL {
    return new URL(`p/${encodeURIComponent(releaseId)}/office-resources.oobpack`, this.baseUrl);
  }

  packageSegmentUrl(segmentSha256: string): URL {
    if (!isDigest(segmentSha256)) throw new ResourceInstallerError('manifest');
    return new URL(`segments/sha256/${segmentSha256}`, this.baseUrl);
  }

  contentObjectUrl(releaseId: string, objectSha256: string): URL {
    if (!/^[a-zA-Z0-9._+-]{1,128}$/.test(releaseId) || !isDigest(objectSha256)) {
      throw new ResourceInstallerError('manifest');
    }
    return new URL(`objects/${encodeURIComponent(releaseId)}/sha256/${objectSha256}`, this.baseUrl);
  }
}

export class ResourcePlanner {
  constructor(
    private readonly manifest: ReleaseManifestV3,
    private readonly installed: ReadonlyMap<string, JournalAsset>,
    private readonly activeRelease: JournalRelease | null = null,
  ) {}

  create(request: ResourcePlanRequest): { plan: ResourcePlan; assets: ReleaseAsset[] } {
    const profiles = this.manifest.package
      ? (['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'] satisfies ResourceProfile[])
      : profilesForRequest(request);
    const selectedPaths = new Set(profiles.flatMap((profile) => this.manifest.profiles[profile] || []));
    const assets = this.manifest.assets.filter((asset) => selectedPaths.has(asset.path));
    const packageReady =
      this.manifest.package &&
      this.activeRelease?.releaseId === this.manifest.releaseId &&
      this.activeRelease.packageSha256 === this.manifest.package.sha256 &&
      this.activeRelease.packageBytes === this.manifest.package.bytes;
    const totalBytes = this.manifest.package
      ? this.manifest.package.bytes
      : assets.reduce((total, asset) => total + asset.bytes, 0);
    const reusedBytes = this.manifest.package
      ? packageReady
        ? this.manifest.package.bytes
        : 0
      : assets.reduce((total, asset) => {
          const record = this.installed.get(asset.path);
          return total + (record?.sha256 === asset.sha256 && record.bytes === asset.bytes ? asset.bytes : 0);
        }, 0);
    const plan: ResourcePlan = {
      planId: crypto.randomUUID(),
      releaseId: this.manifest.releaseId,
      scope: request.scope,
      profiles,
      totalBytes,
      downloadBytes: totalBytes - reusedBytes,
      reusedBytes,
    };
    return { plan, assets };
  }
}

function profilesForRequest(request: ResourcePlanRequest): ResourceProfile[] {
  if (request.profiles?.length) return [...new Set(request.profiles)];
  if (request.scope === 'all') {
    return ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'];
  }
  if (request.scope === 'fonts') return ['fonts-basic', 'fonts-office-compat'];
  if (request.scope === 'document') {
    return ['base', 'fonts-basic', request.documentType || 'word'];
  }
  return ['base', 'fonts-basic'];
}

export class ResourceInstallerError extends Error {
  constructor(
    readonly code: ResourceErrorCode,
    readonly path?: string,
  ) {
    super(path ? `${code}: ${path}` : code);
    this.name = 'ResourceInstallerError';
  }
}

function networkError(status?: number): ResourceInstallerError {
  if (typeof navigator !== 'undefined' && navigator.onLine === false) return new ResourceInstallerError('offline');
  if (status === 507) return new ResourceInstallerError('quota');
  return new ResourceInstallerError('network');
}

function bytesToHex(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function readAndVerify(
  response: Response,
  expectedDigest: string,
  onDownload: (bytes: number) => void,
  signal: AbortSignal,
): Promise<Uint8Array> {
  const digest = sha256.create();
  const chunks: Uint8Array[] = [];
  let length = 0;
  const reader = response.body?.getReader();
  if (!reader) {
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (signal.aborted) throw new ResourceInstallerError('aborted');
    digest.update(bytes);
    onDownload(bytes.byteLength);
    chunks.push(bytes);
    length = bytes.byteLength;
  } else {
    while (true) {
      if (signal.aborted) throw new ResourceInstallerError('aborted');
      const next = await reader.read();
      if (next.done) break;
      digest.update(next.value);
      chunks.push(next.value);
      length += next.value.byteLength;
      onDownload(next.value.byteLength);
    }
  }
  if (bytesToHex(digest.digest()) !== expectedDigest) throw new ResourceInstallerError('integrity');
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return result;
}

async function readAndVerifySegment(
  response: Response,
  expectedDigest: string,
  onDownload: (bytes: Uint8Array) => void | Promise<void>,
  signal: AbortSignal,
  waitIfPaused: () => Promise<void>,
): Promise<Uint8Array> {
  const digest = sha256.create();
  let length = 0;
  const chunks: Uint8Array[] = [];
  const reader = response.body?.getReader();
  if (!reader) {
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (signal.aborted) throw new ResourceInstallerError('aborted');
    digest.update(bytes);
    chunks.push(bytes);
    await onDownload(bytes);
    length = bytes.byteLength;
  } else {
    const readWithAbort = async (): Promise<ReadableStreamReadResult<Uint8Array>> => {
      if (signal.aborted) {
        void reader.cancel(signal.reason).catch(() => undefined);
        throw new ResourceInstallerError('aborted');
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
          reject(new ResourceInstallerError('aborted'));
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
    while (true) {
      await waitIfPaused();
      if (signal.aborted) throw new ResourceInstallerError('aborted');
      const next = await readWithAbort();
      if (next.done) break;
      digest.update(next.value);
      chunks.push(next.value);
      length += next.value.byteLength;
      await onDownload(next.value);
    }
  }
  if (bytesToHex(digest.digest()) !== expectedDigest) throw new ResourceInstallerError('integrity');
  const result = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return result;
}

export type ResourceInstallerOptions = {
  assetBaseUrl: string | URL;
  fetch: typeof fetch;
  cacheStorage?: CacheStorage;
  journal: InstallationJournal;
  storageMode: 'cache-storage' | 'http-cache';
  timeoutMs?: number;
  retryDelaysMs?: number[];
  connection?: { saveData?: boolean; effectiveType?: string };
  locks?: LockManager;
  broadcast?: BroadcastChannel;
  legacyStorage?: Storage;
};

const initialSnapshot = (storageMode: 'cache-storage' | 'http-cache'): ResourceInstallerSnapshot => ({
  installedRelease: null,
  targetRelease: null,
  availableRelease: null,
  availablePackageVersion: null,
  readiness: 'checking',
  phase: 'idle',
  storageMode,
  currentChunk: null,
  currentChunkIndex: 0,
  currentChunkCount: 0,
  downloadedBytes: 0,
  downloadBytes: 0,
  verifiedBytes: 0,
  verifyBytes: 0,
  bytesPerSecond: 0,
  failedResources: [],
  canPause: false,
  canResume: false,
  canRetry: false,
  errorCode: null,
  installedProfiles: [],
});

export class TransactionalResourceInstaller implements OfficeRuntimeResourceInstaller {
  private readonly repository: ReleaseRepository;
  private readonly options: ResourceInstallerOptions;
  private readonly listeners = new Set<(snapshot: ResourceInstallerSnapshot) => void>();
  private readonly plans = new Map<
    string,
    { manifest: ReleaseManifestV3; assets: ReleaseAsset[]; plan: ResourcePlan }
  >();
  private snapshot: ResourceInstallerSnapshot;
  private currentManifest: ReleaseManifestV3 | null = null;
  private abortController: AbortController | null = null;
  private paused = false;
  private resumePromise: Promise<void> | null = null;
  private resolveResume: (() => void) | null = null;
  private lastPlan: ResourcePlan | null = null;
  private readonly installedPaths = new Set<string>();

  constructor(options: ResourceInstallerOptions) {
    this.options = options;
    this.repository = new ReleaseRepository(options.assetBaseUrl, options.fetch);
    this.snapshot = initialSnapshot(options.storageMode);
    options.broadcast?.addEventListener('message', (event) => {
      if (event.data?.type !== 'resource-snapshot' || !event.data.snapshot) return;
      const incoming = event.data.snapshot as ResourceInstallerSnapshot;
      const currentRelease = this.currentManifest?.releaseId || this.snapshot.targetRelease;
      const incomingRelease = incoming.targetRelease || incoming.availableRelease;
      if (currentRelease && incomingRelease && incomingRelease !== currentRelease) return;
      this.snapshot = incoming;
      this.publish(false);
    });
  }

  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getInstallerSnapshot(): ResourceInstallerSnapshot {
    return {
      ...this.snapshot,
      failedResources: this.snapshot.failedResources.map((failure) => ({ ...failure })),
    };
  }

  getInstalledPaths(): string[] {
    return [...this.installedPaths];
  }

  async initialize(): Promise<void> {
    const [active, release] = await Promise.all([this.options.journal.getActiveRelease(), this.repository.current()]);
    const installedAssets = await this.options.journal.listAssets(release.manifest.releaseId);
    this.installedPaths.clear();
    for (const asset of installedAssets) this.installedPaths.add(asset.path);
    this.currentManifest = release.manifest;
    const packageLedgerReady =
      !release.manifest.package ||
      (active?.packageSha256 === release.manifest.package.sha256 &&
        active.packageBytes === release.manifest.package.bytes);
    const packageStorageReady =
      !release.manifest.package ||
      (await this.findMissingPackageSegments(release.manifest.releaseId, release.manifest.package)).length === 0;
    const packageReady = packageLedgerReady && packageStorageReady;
    this.snapshot = {
      ...this.snapshot,
      installedRelease: active?.releaseId || null,
      installedProfiles: active?.installedProfiles || [],
      targetRelease: release.manifest.releaseId,
      availableRelease: release.manifest.releaseId,
      availablePackageVersion: release.manifest.packageVersion,
      readiness:
        active?.releaseId === release.manifest.releaseId && packageReady
          ? 'ready'
          : active?.releaseId === release.manifest.releaseId
            ? 'repair-needed'
            : active
              ? 'update-available'
              : 'needs-download',
    };
    this.publish();
  }

  async plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    this.patch({ phase: 'planning', errorCode: null });
    const manifest = await this.ensureManifest();
    const installed = new Map(
      (await this.options.journal.listAssets(manifest.releaseId)).map((item) => [item.path, item]),
    );
    const active = await this.options.journal.getActiveRelease();
    const planned = new ResourcePlanner(manifest, installed, active).create(request);
    if (manifest.package) {
      const missing = await this.findMissingPackageSegments(manifest.releaseId, manifest.package);
      const missingIds = new Set(missing.map((segment) => segment.id));
      const downloadBytes = manifest.package.segments
        .filter((segment) => missingIds.has(segment.id))
        .reduce((sum, segment) => sum + segment.bytes, 0);
      planned.plan.downloadBytes = downloadBytes;
      planned.plan.reusedBytes = manifest.package.bytes - downloadBytes;
    }
    this.plans.set(planned.plan.planId, { manifest, assets: planned.assets, plan: planned.plan });
    this.patch({
      phase: 'idle',
      downloadedBytes: 0,
      downloadBytes: planned.plan.downloadBytes,
      verifiedBytes: 0,
      verifyBytes: planned.plan.downloadBytes,
    });
    return { ...planned.plan, profiles: [...planned.plan.profiles] };
  }

  async apply(plan: ResourcePlan): Promise<void> {
    const internal = this.plans.get(plan.planId);
    if (!internal || internal.plan.releaseId !== plan.releaseId) throw new ResourceInstallerError('incompatible');
    this.lastPlan = plan;
    const run = () => this.applyUnlocked(internal);
    if (this.options.locks) {
      await this.options.locks.request(`onlyoffice-resources:${plan.releaseId}`, { mode: 'exclusive' }, run);
    } else {
      await run();
    }
  }

  async checkForUpdates(): Promise<void> {
    const { manifest } = await this.repository.current();
    const active = await this.options.journal.getActiveRelease();
    const packageLedgerReady =
      !manifest.package ||
      (active?.packageSha256 === manifest.package.sha256 && active.packageBytes === manifest.package.bytes);
    const packageStorageReady =
      !manifest.package || (await this.findMissingPackageSegments(manifest.releaseId, manifest.package)).length === 0;
    const packageReady = packageLedgerReady && packageStorageReady;
    this.currentManifest = manifest;
    this.patch({
      availableRelease: manifest.releaseId,
      targetRelease: manifest.releaseId,
      availablePackageVersion: manifest.packageVersion,
      readiness:
        this.snapshot.installedRelease === manifest.releaseId && packageReady
          ? 'ready'
          : this.snapshot.installedRelease === manifest.releaseId
            ? 'repair-needed'
            : this.snapshot.installedRelease
              ? 'update-available'
              : 'needs-download',
    });
  }

  async checkHealth(): Promise<void> {
    const manifest = await this.ensureManifest();
    if (manifest.package) {
      const active = await this.options.journal.getActiveRelease();
      const ledgerReady =
        active?.releaseId === manifest.releaseId &&
        active.packageSha256 === manifest.package.sha256 &&
        active.packageBytes === manifest.package.bytes;
      const missing = await this.findMissingPackageSegments(manifest.releaseId, manifest.package);
      const ready = ledgerReady && missing.length === 0;
      const failures: FailedResource[] = !ledgerReady
        ? [{ path: manifest.package.path, code: 'integrity', attempts: 0 }]
        : missing.map((segment) => ({
            path: `${manifest.package!.path}?segment=${segment.id}`,
            code: 'storage' as const,
            attempts: 0,
          }));
      this.patch({
        failedResources: failures,
        readiness: ready ? 'ready' : 'repair-needed',
        errorCode: ready ? null : failures[0]?.code || 'integrity',
      });
      return;
    }
    const installed = await this.options.journal.listAssets(manifest.releaseId);
    const expected = new Map(manifest.assets.map((asset) => [asset.path, asset]));
    const failures = installed
      .filter((record) => {
        const asset = expected.get(record.path);
        return !asset || asset.sha256 !== record.sha256 || asset.bytes !== record.bytes;
      })
      .map((record) => ({ path: record.path, code: 'integrity' as const, attempts: 0 }));
    this.patch({
      failedResources: failures,
      readiness: failures.length ? 'repair-needed' : installed.length ? 'ready' : 'needs-download',
      errorCode: failures.length ? 'integrity' : null,
    });
  }

  async repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void> {
    const manifest = await this.ensureManifest();
    if (manifest.package) {
      const missing = await this.findMissingPackageSegments(manifest.releaseId, manifest.package);
      const downloadBytes = missing.reduce((sum, segment) => sum + segment.bytes, 0);
      const plan: ResourcePlan = {
        planId: crypto.randomUUID(),
        releaseId: manifest.releaseId,
        scope: 'repair',
        profiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
        totalBytes: manifest.package.bytes,
        downloadBytes,
        reusedBytes: manifest.package.bytes - downloadBytes,
      };
      this.plans.set(plan.planId, { manifest, assets: manifest.assets, plan });
      await this.apply(plan);
      return;
    }
    const records = await this.options.journal.listAssets(manifest.releaseId);
    const installedPaths = new Set(records.map((record) => record.path));
    const requiredProfiles: ResourceProfile[] =
      options.scope === 'all'
        ? ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat']
        : options.scope === 'required'
          ? ['base', 'fonts-basic']
          : [];
    const candidates =
      options.scope === 'installed'
        ? manifest.assets.filter((asset) => installedPaths.has(asset.path))
        : manifest.assets.filter((asset) => requiredProfiles.includes(asset.profile));
    const corrupt: ReleaseAsset[] = [];
    this.patch({
      phase: 'repairing',
      readiness: 'updating',
      verifyBytes: candidates.reduce((sum, asset) => sum + asset.bytes, 0),
      verifiedBytes: 0,
      failedResources: [],
    });
    for (const asset of candidates) {
      const record = records.find((candidate) => candidate.path === asset.path);
      if (!record || record.sha256 !== asset.sha256 || record.bytes !== asset.bytes) {
        corrupt.push(asset);
        continue;
      }
      try {
        await this.verifyInstalledAsset(asset, manifest.releaseId);
      } catch {
        corrupt.push(asset);
        await this.options.journal.deleteAsset(manifest.releaseId, asset.path);
        this.installedPaths.delete(asset.path);
      }
    }
    if (!corrupt.length) {
      this.patch({ phase: 'idle', readiness: 'ready' });
      return;
    }
    const profiles = [...new Set(corrupt.map((asset) => asset.profile))];
    const totalBytes = corrupt.reduce((sum, asset) => sum + asset.bytes, 0);
    const repairPlan: ResourcePlan = {
      planId: crypto.randomUUID(),
      releaseId: manifest.releaseId,
      scope: 'repair',
      profiles,
      totalBytes,
      downloadBytes: totalBytes,
      reusedBytes: 0,
    };
    this.plans.set(repairPlan.planId, { manifest, assets: corrupt, plan: repairPlan });
    await this.apply(repairPlan);
  }

  pause(): void {
    if (!this.abortController || this.paused) return;
    this.paused = true;
    this.resumePromise = new Promise((resolve) => {
      this.resolveResume = resolve;
    });
    this.patch({ phase: 'paused', readiness: 'paused', canPause: false, canResume: true });
  }

  async resume(): Promise<void> {
    if (!this.paused) {
      if (this.lastPlan && this.snapshot.canRetry) await this.apply(this.lastPlan);
      return;
    }
    this.paused = false;
    this.resolveResume?.();
    this.resolveResume = null;
    this.resumePromise = null;
    this.patch({ phase: 'downloading', readiness: 'updating', canPause: true, canResume: false });
  }

  cancel(): void {
    this.paused = false;
    this.resolveResume?.();
    this.abortController?.abort();
    this.patch({ phase: 'idle', readiness: 'needs-download', canPause: false, canResume: false });
  }

  private async ensureManifest(): Promise<ReleaseManifestV3> {
    if (!this.currentManifest) await this.initialize();
    if (!this.currentManifest) throw new ResourceInstallerError('manifest');
    return this.currentManifest;
  }

  private async findMissingPackageSegments(releaseId: string, pack: ReleasePackage): Promise<ReleasePackageSegment[]> {
    void releaseId;
    const verified = new Map((await this.options.journal.listSegments()).map((segment) => [segment.sha256, segment]));
    const missing: ReleasePackageSegment[] = [];
    for (const segment of pack.segments) {
      const record = verified.get(segment.sha256);
      if (!record || record.bytes !== segment.bytes) {
        missing.push(segment);
        continue;
      }
      if (
        this.options.storageMode === 'cache-storage' &&
        this.options.cacheStorage &&
        typeof this.options.cacheStorage.match === 'function'
      ) {
        const response = await this.options.cacheStorage.match(this.repository.packageSegmentUrl(segment.sha256));
        const contentLength = response ? Number(response.headers.get('content-length')) : Number.NaN;
        if (!response?.ok || !Number.isFinite(contentLength) || contentLength !== segment.bytes) {
          await this.options.journal.deleteSegment(segment.sha256);
          missing.push(segment);
        }
      }
    }
    return missing;
  }

  private async applyUnlocked(internal: {
    manifest: ReleaseManifestV3;
    assets: ReleaseAsset[];
    plan: ResourcePlan;
  }): Promise<void> {
    if (internal.manifest.package) {
      await this.applyPackageUnlocked(internal);
      return;
    }
    const existing = new Map(
      (await this.options.journal.listAssets(internal.manifest.releaseId)).map((asset) => [asset.path, asset]),
    );
    const pending = internal.assets.filter((asset) => existing.get(asset.path)?.sha256 !== asset.sha256);
    const pendingChunks = [...new Set(pending.map((asset) => asset.chunk))];
    const chunkIndexes = new Map(pendingChunks.map((chunk, index) => [chunk, index + 1]));
    this.abortController = new AbortController();
    const startedAt = performance.now();
    this.patch({
      targetRelease: internal.manifest.releaseId,
      readiness: 'updating',
      phase: 'downloading',
      currentChunk: pending[0]?.chunk || null,
      currentChunkIndex: pending.length ? 1 : 0,
      currentChunkCount: pendingChunks.length,
      downloadedBytes: 0,
      downloadBytes: pending.reduce((sum, asset) => sum + asset.bytes, 0),
      verifiedBytes: 0,
      verifyBytes: pending.reduce((sum, asset) => sum + asset.bytes, 0),
      bytesPerSecond: 0,
      failedResources: [],
      canPause: pending.length > 0,
      canResume: false,
      canRetry: false,
      errorCode: null,
    });
    const failures: FailedResource[] = [];
    let cursor = 0;
    const connection = this.options.connection;
    const concurrency = connection?.saveData || ['slow-2g', '2g'].includes(connection?.effectiveType || '') ? 2 : 4;
    const worker = async () => {
      while (cursor < pending.length) {
        const asset = pending[cursor++];
        if (this.paused) await this.resumePromise;
        if (this.abortController?.signal.aborted) throw new ResourceInstallerError('aborted');
        this.patch({
          currentChunk: asset.chunk,
          currentChunkIndex: chunkIndexes.get(asset.chunk) || 0,
          currentChunkCount: pendingChunks.length,
        });
        try {
          await this.transfer(asset, internal.manifest.releaseId);
        } catch (error) {
          const code = error instanceof ResourceInstallerError ? error.code : 'network';
          failures.push({ path: asset.path, code, attempts: this.retryDelays().length + 1 });
        }
        const seconds = Math.max((performance.now() - startedAt) / 1_000, 0.001);
        this.patch({ bytesPerSecond: Math.round(this.snapshot.downloadedBytes / seconds) });
      }
    };
    try {
      await Promise.all(Array.from({ length: Math.min(concurrency, pending.length) }, worker));
      if (failures.length) {
        const first = failures[0];
        this.patch({
          phase: 'idle',
          readiness: 'error',
          failedResources: failures,
          canPause: false,
          canRetry: true,
          errorCode: first.code,
        });
        throw new ResourceInstallerError(first.code, first.path);
      }
      this.patch({ phase: 'activating', canPause: false });
      const installedProfiles = [...new Set(internal.assets.map((asset) => asset.profile))];
      await this.options.journal.activateRelease({
        releaseId: internal.manifest.releaseId,
        installedProfiles,
        activatedAt: Date.now(),
      });
      this.patch({
        installedRelease: internal.manifest.releaseId,
        installedProfiles,
        phase: 'idle',
        currentChunk: null,
        currentChunkIndex: 0,
        currentChunkCount: 0,
        readiness: 'ready',
        canRetry: false,
      });
      this.cleanupLegacyState();
    } finally {
      this.abortController = null;
    }
  }

  private async applyPackageUnlocked(internal: {
    manifest: ReleaseManifestV3;
    assets: ReleaseAsset[];
    plan: ResourcePlan;
  }): Promise<void> {
    const pack = internal.manifest.package;
    if (!pack) throw new ResourceInstallerError('manifest');
    const active = await this.options.journal.getActiveRelease();
    const alreadyInstalled =
      internal.plan.scope !== 'repair' &&
      internal.plan.downloadBytes === 0 &&
      active?.releaseId === internal.manifest.releaseId &&
      active.packageSha256 === pack.sha256 &&
      active.packageBytes === pack.bytes;
    this.abortController = new AbortController();
    const startedAt = performance.now();
    this.patch({
      targetRelease: internal.manifest.releaseId,
      readiness: alreadyInstalled ? 'ready' : 'updating',
      phase: alreadyInstalled ? 'activating' : 'downloading',
      currentChunk: alreadyInstalled ? null : pack.segments[0]?.id || null,
      currentChunkIndex: alreadyInstalled ? 0 : 1,
      currentChunkCount: alreadyInstalled ? 0 : pack.segments.length,
      downloadedBytes: 0,
      downloadBytes: alreadyInstalled ? 0 : internal.plan.downloadBytes,
      verifiedBytes: 0,
      verifyBytes: alreadyInstalled ? 0 : internal.plan.downloadBytes,
      bytesPerSecond: 0,
      failedResources: [],
      canPause: !alreadyInstalled,
      canResume: false,
      canRetry: false,
      errorCode: null,
    });
    try {
      if (!alreadyInstalled) {
        await this.transferPackage(internal.manifest.releaseId, pack, startedAt);
      }
      this.patch({
        phase: 'activating',
        currentChunk: null,
        currentChunkIndex: 0,
        currentChunkCount: 0,
        canPause: false,
      });
      for (let offset = 0; offset < internal.assets.length; offset += 24) {
        await Promise.all(
          internal.assets.slice(offset, offset + 24).map((asset) =>
            this.options.journal.putAsset({
              releaseId: internal.manifest.releaseId,
              path: asset.path,
              sha256: asset.sha256,
              bytes: asset.bytes,
              verifiedAt: Date.now(),
            }),
          ),
        );
      }
      this.installedPaths.clear();
      for (const asset of internal.assets) this.installedPaths.add(asset.path);
      const installedProfiles: ResourceProfile[] = [
        'base',
        'word',
        'cell',
        'slide',
        'fonts-basic',
        'fonts-office-compat',
      ];
      await this.options.journal.activateRelease({
        releaseId: internal.manifest.releaseId,
        installedProfiles,
        activatedAt: Date.now(),
        packageSha256: pack.sha256,
        packageBytes: pack.bytes,
      });
      this.patch({
        installedRelease: internal.manifest.releaseId,
        installedProfiles,
        phase: 'idle',
        currentChunk: null,
        currentChunkIndex: 0,
        currentChunkCount: 0,
        readiness: 'ready',
        downloadedBytes: this.snapshot.downloadBytes,
        verifiedBytes: this.snapshot.verifyBytes,
        bytesPerSecond: Math.round(
          this.snapshot.downloadBytes / Math.max((performance.now() - startedAt) / 1_000, 0.001),
        ),
        canPause: false,
        canResume: false,
        canRetry: false,
      });
      this.cleanupLegacyState();
    } catch (error) {
      const failure =
        error instanceof ResourceInstallerError ? error : new ResourceInstallerError('network', pack.path);
      this.patch({
        phase: 'idle',
        readiness: 'error',
        failedResources: [{ path: pack.path, code: failure.code, attempts: this.retryDelays().length + 1 }],
        canPause: false,
        canResume: false,
        canRetry: true,
        errorCode: failure.code,
      });
      throw failure;
    } finally {
      this.abortController = null;
    }
  }

  private retryDelays(): number[] {
    return this.options.retryDelaysMs || [1_000, 3_000, 10_000];
  }

  private async transferPackage(releaseId: string, pack: ReleasePackage, startedAt: number): Promise<void> {
    const completeDigest = sha256.create();
    const missing = await this.findMissingPackageSegments(releaseId, pack);
    const transferBytes = missing.reduce((sum, segment) => sum + segment.bytes, 0);
    this.patch({
      currentChunk: missing[0]?.id || null,
      currentChunkIndex: missing.length ? 1 : 0,
      currentChunkCount: missing.length,
      downloadedBytes: 0,
      downloadBytes: transferBytes,
      verifiedBytes: 0,
      verifyBytes: transferBytes,
      canPause: missing.length > 0,
    });
    for (const [segmentIndex, segment] of missing.entries()) {
      const delays = [0, ...this.retryDelays()];
      let segmentBytes: Uint8Array | null = null;
      let lastError: unknown;
      const url = this.repository.packageSegmentUrl(segment.sha256);
      this.patch({
        phase: 'downloading',
        currentChunk: segment.id,
        currentChunkIndex: segmentIndex + 1,
        currentChunkCount: missing.length,
      });
      if (
        this.options.storageMode === 'cache-storage' &&
        this.options.cacheStorage &&
        typeof this.options.cacheStorage.match === 'function'
      ) {
        const cached = await this.options.cacheStorage.match(url);
        if (cached) {
          try {
            segmentBytes = await readAndVerifySegment(
              cached,
              segment.sha256,
              () => undefined,
              this.abortController?.signal || new AbortController().signal,
              async () => {
                if (this.paused) await this.resumePromise;
              },
            );
            if (segmentBytes.byteLength !== segment.bytes) {
              throw new ResourceInstallerError('integrity', pack.path);
            }
            this.patch({
              phase: 'verifying',
              currentChunk: segment.id,
              currentChunkIndex: segmentIndex + 1,
              currentChunkCount: missing.length,
              verifiedBytes: this.snapshot.verifiedBytes + segmentBytes.byteLength,
            });
          } catch {
            segmentBytes = null;
            this.patch({
              downloadedBytes: Math.max(0, this.snapshot.downloadedBytes - segment.bytes),
            });
            const cache = await this.options.cacheStorage.open(RELEASE_SEGMENT_CACHE_NAME);
            await cache.delete(url);
          }
        }
      }
      for (let attempt = 0; attempt < delays.length && !segmentBytes; attempt += 1) {
        let attemptDownloaded = 0;
        if (delays[attempt]) await new Promise((resolve) => setTimeout(resolve, delays[attempt]));
        if (this.paused) await this.resumePromise;
        const controller = new AbortController();
        const parentSignal = this.abortController?.signal;
        const abort = () => controller.abort();
        parentSignal?.addEventListener('abort', abort, { once: true });
        let idleTimeout: ReturnType<typeof setTimeout> | undefined;
        const resetIdleTimeout = () => {
          if (idleTimeout) clearTimeout(idleTimeout);
          idleTimeout = setTimeout(() => controller.abort(), this.options.timeoutMs || 30_000);
        };
        resetIdleTimeout();
        try {
          const response = await this.options.fetch(url, {
            cache: attempt === 0 ? 'force-cache' : 'reload',
            credentials: 'omit',
            mode: url.origin === location.origin ? 'same-origin' : 'cors',
            signal: controller.signal,
          });
          if (!response.ok) throw networkError(response.status);
          const cacheResponse = response.clone();
          const contentLength = Number(response.headers.get('content-length'));
          if (Number.isFinite(contentLength) && contentLength !== segment.bytes) {
            throw new ResourceInstallerError('integrity', pack.path);
          }
          const bytes = await readAndVerifySegment(
            response,
            segment.sha256,
            async (chunk) => {
              resetIdleTimeout();
              attemptDownloaded += chunk.byteLength;
              const downloadedBytes = this.snapshot.downloadedBytes + chunk.byteLength;
              const seconds = Math.max((performance.now() - startedAt) / 1_000, 0.001);
              this.patch({
                phase: 'downloading',
                currentChunk: segment.id,
                currentChunkIndex: segmentIndex + 1,
                currentChunkCount: missing.length,
                downloadedBytes,
                bytesPerSecond: Math.round(downloadedBytes / seconds),
              });
            },
            controller.signal,
            async () => {
              if (this.paused) await this.resumePromise;
            },
          );
          if (bytes.byteLength !== segment.bytes) throw new ResourceInstallerError('integrity', pack.path);
          if (this.options.storageMode === 'cache-storage' && this.options.cacheStorage) {
            try {
              const cache = await this.options.cacheStorage.open(RELEASE_SEGMENT_CACHE_NAME);
              await cache.put(url, cacheResponse);
            } catch (error) {
              throw new ResourceInstallerError(
                error instanceof DOMException && error.name === 'QuotaExceededError' ? 'quota' : 'storage',
                pack.path,
              );
            }
          }
          this.patch({
            phase: 'verifying',
            currentChunk: segment.id,
            currentChunkIndex: segmentIndex + 1,
            currentChunkCount: missing.length,
            verifiedBytes: this.snapshot.verifiedBytes + bytes.byteLength,
          });
          segmentBytes = bytes;
        } catch (error) {
          this.patch({
            downloadedBytes: Math.max(0, this.snapshot.downloadedBytes - attemptDownloaded),
            phase: 'downloading',
          });
          lastError =
            controller.signal.aborted && !parentSignal?.aborted
              ? new ResourceInstallerError('timeout', pack.path)
              : error instanceof ResourceInstallerError && !error.path
                ? new ResourceInstallerError(error.code, pack.path)
                : error;
        } finally {
          if (idleTimeout) clearTimeout(idleTimeout);
          parentSignal?.removeEventListener('abort', abort);
        }
      }
      if (!segmentBytes) {
        if (lastError instanceof ResourceInstallerError) throw lastError;
        throw networkError();
      }
      await this.options.journal.putSegment({
        sha256: segment.sha256,
        bytes: segment.bytes,
        verifiedAt: Date.now(),
      });
      completeDigest.update(segmentBytes);
    }
    if (missing.length === pack.segments.length && bytesToHex(completeDigest.digest()) !== pack.sha256) {
      throw new ResourceInstallerError('integrity', pack.path);
    }
    this.patch({
      phase: 'verifying',
      currentChunk: null,
      currentChunkIndex: 0,
      currentChunkCount: missing.length,
      verifiedBytes: transferBytes,
    });
  }

  private async transfer(asset: ReleaseAsset, releaseId: string): Promise<void> {
    const delays = [0, ...this.retryDelays()];
    let lastError: unknown;
    for (let attempt = 0; attempt < delays.length; attempt += 1) {
      let attemptDownloaded = 0;
      let attemptVerified = 0;
      if (delays[attempt]) await new Promise((resolve) => setTimeout(resolve, delays[attempt]));
      if (this.paused) await this.resumePromise;
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), this.options.timeoutMs || 30_000);
      const parentSignal = this.abortController?.signal;
      const abort = () => controller.abort();
      parentSignal?.addEventListener('abort', abort, { once: true });
      try {
        const url = this.repository.assetUrl(releaseId, asset.path);
        const response = await this.options.fetch(url, {
          cache: attempt === 0 ? 'force-cache' : 'reload',
          credentials: 'omit',
          mode: url.origin === location.origin ? 'same-origin' : 'cors',
          signal: controller.signal,
        });
        if (!response.ok) throw networkError(response.status);
        const cacheResponse = response.clone();
        this.patch({ phase: 'downloading' });
        const bytes = await readAndVerify(
          response,
          asset.sha256,
          (count) => {
            attemptDownloaded += count;
            this.patch({ downloadedBytes: this.snapshot.downloadedBytes + count });
          },
          controller.signal,
        );
        if (bytes.byteLength !== asset.bytes) throw new ResourceInstallerError('integrity', asset.path);
        attemptVerified = bytes.byteLength;
        this.patch({
          phase: 'verifying',
          verifiedBytes: this.snapshot.verifiedBytes + attemptVerified,
        });
        if (this.options.storageMode === 'cache-storage' && this.options.cacheStorage) {
          const cache = await this.options.cacheStorage.open(`onlyoffice-release-staging-${releaseId}`);
          await cache.put(url, cacheResponse);
        }
        await this.options.journal.putAsset({
          releaseId,
          path: asset.path,
          sha256: asset.sha256,
          bytes: asset.bytes,
          verifiedAt: Date.now(),
        });
        this.installedPaths.add(asset.path);
        return;
      } catch (error) {
        this.patch({
          downloadedBytes: Math.max(0, this.snapshot.downloadedBytes - attemptDownloaded),
          verifiedBytes: Math.max(0, this.snapshot.verifiedBytes - attemptVerified),
        });
        lastError =
          controller.signal.aborted && !parentSignal?.aborted
            ? new ResourceInstallerError('timeout', asset.path)
            : error;
      } finally {
        clearTimeout(timeout);
        parentSignal?.removeEventListener('abort', abort);
      }
    }
    if (lastError instanceof ResourceInstallerError) throw lastError;
    throw networkError();
  }

  private async verifyInstalledAsset(asset: ReleaseAsset, releaseId: string): Promise<void> {
    const url = this.repository.assetUrl(releaseId, asset.path);
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), this.options.timeoutMs || 30_000);
    try {
      const response = await this.options.fetch(url, {
        cache: 'force-cache',
        credentials: 'omit',
        mode: url.origin === location.origin ? 'same-origin' : 'cors',
        signal: controller.signal,
      });
      if (!response.ok) throw networkError(response.status);
      const bytes = await readAndVerify(
        response,
        asset.sha256,
        (count) => this.patch({ verifiedBytes: this.snapshot.verifiedBytes + count }),
        controller.signal,
      );
      if (bytes.byteLength !== asset.bytes) throw new ResourceInstallerError('integrity', asset.path);
    } finally {
      clearTimeout(timeout);
    }
  }

  private patch(next: Partial<ResourceInstallerSnapshot>): void {
    this.snapshot = { ...this.snapshot, ...next };
    this.publish();
  }

  private cleanupLegacyState(): void {
    try {
      this.options.legacyStorage?.removeItem('onlyoffice-browser:shared-runtime-cache');
      this.options.legacyStorage?.removeItem('onlyoffice-browser:installed-fonts');
    } catch {
      // A completed v3 installation remains authoritative when legacy storage
      // is unavailable or policy-restricted.
    }
    if (!this.options.cacheStorage) return;
    void this.options.cacheStorage
      .keys()
      .then((keys) =>
        Promise.all(
          keys
            .filter(
              (key) =>
                key.startsWith('onlyoffice-browser-font-packages-') || key.startsWith('onlyoffice-release-staging-'),
            )
            .map((key) => this.options.cacheStorage!.delete(key)),
        ),
      );
  }

  private publish(broadcast = true): void {
    const snapshot = this.getInstallerSnapshot();
    for (const listener of this.listeners) listener(snapshot);
    if (broadcast) this.options.broadcast?.postMessage({ type: 'resource-snapshot', snapshot });
  }
}

export async function createTransactionalResourceInstaller(
  options: Partial<ResourceInstallerOptions> & { assetBaseUrl?: string | URL } = {},
): Promise<TransactionalResourceInstaller> {
  const hasIndexedDb = typeof indexedDB !== 'undefined';
  const storageMode =
    options.storageMode ||
    (new URL(options.assetBaseUrl || location.origin).origin === location.origin ? 'cache-storage' : 'http-cache');
  const connection = (navigator as Navigator & { connection?: { saveData?: boolean; effectiveType?: string } })
    .connection;
  const installer = new TransactionalResourceInstaller({
    assetBaseUrl: options.assetBaseUrl || location.origin,
    fetch: options.fetch || window.fetch.bind(window),
    cacheStorage: options.cacheStorage ?? (typeof caches === 'undefined' ? undefined : caches),
    journal: options.journal || (hasIndexedDb ? new IndexedDbInstallationJournal() : new MemoryInstallationJournal()),
    storageMode,
    timeoutMs: options.timeoutMs,
    retryDelaysMs: options.retryDelaysMs,
    connection: options.connection || connection,
    locks: options.locks ?? navigator.locks,
    broadcast:
      options.broadcast ??
      (typeof BroadcastChannel === 'undefined' ? undefined : new BroadcastChannel('onlyoffice-resources-v3')),
    legacyStorage: options.legacyStorage ?? (typeof localStorage === 'undefined' ? undefined : localStorage),
  });
  await installer.initialize();
  return installer;
}
