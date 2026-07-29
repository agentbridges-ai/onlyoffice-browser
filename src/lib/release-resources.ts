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
}

export interface ReleaseChunk {
  id: string;
  profile: ResourceProfile;
  bytes: number;
  paths: string[];
}

export interface ReleaseManifestV3 {
  version: 3;
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
  assets: ReleaseAsset[];
  fontFamilies?: Array<{ name: string; paths: string[] }>;
}

export interface ReleaseChannel {
  version: 1;
  releaseId: string;
  manifestUrl?: string;
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
  readiness: ResourceReadiness;
  phase: ResourcePhase;
  storageMode: 'cache-storage' | 'http-cache';
  currentChunk: string | null;
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
  checkHealth(options?: { deep?: false }): Promise<void>;
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
};

export interface InstallationJournal {
  listAssets(releaseId: string): Promise<JournalAsset[]>;
  putAsset(asset: JournalAsset): Promise<void>;
  deleteAsset(releaseId: string, path: string): Promise<void>;
  getActiveRelease(): Promise<JournalRelease | null>;
  activateRelease(release: JournalRelease): Promise<void>;
}

const DB_NAME = 'onlyoffice-browser-resources-v3';
const DB_VERSION = 1;

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
}

function isSafePath(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    value.length > 0 &&
    !value.startsWith('/') &&
    !value.includes('\\') &&
    !value.split('/').includes('..')
  );
}

function isDigest(value: unknown): value is string {
  return typeof value === 'string' && /^[a-f0-9]{64}$/.test(value);
}

function isProfile(value: unknown): value is ResourceProfile {
  return ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'].includes(String(value));
}

export function parseReleaseManifest(value: unknown): ReleaseManifestV3 {
  const manifest = value as Partial<ReleaseManifestV3>;
  if (
    manifest.version !== 3 ||
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
      asset.chunk.length > 0,
  );
  if (assets.length !== manifest.assets.length || new Set(assets.map((asset) => asset.path)).size !== assets.length) {
    throw new ResourceInstallerError('manifest');
  }
  const assetPaths = new Set(assets.map((asset) => asset.path));
  for (const profile of Object.keys(manifest.profiles)) {
    if (!isProfile(profile) || !Array.isArray(manifest.profiles[profile])) throw new ResourceInstallerError('manifest');
    if (!manifest.profiles[profile].every((path) => assetPaths.has(path))) throw new ResourceInstallerError('manifest');
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
    const channelUrl = new URL('channels/stable.json', this.baseUrl);
    const channelResponse = await this.fetchImpl(channelUrl, {
      cache: 'no-store',
      credentials: 'omit',
      mode: channelUrl.origin === location.origin ? 'same-origin' : 'cors',
    });
    if (!channelResponse.ok) throw networkError(channelResponse.status);
    const channel = (await channelResponse.json()) as Partial<ReleaseChannel>;
    if (channel.version !== 1 || typeof channel.releaseId !== 'string') {
      throw new ResourceInstallerError('manifest');
    }
    const manifestUrl = channel.manifestUrl
      ? new URL(channel.manifestUrl, channelUrl)
      : new URL(`releases/${encodeURIComponent(channel.releaseId)}/manifest.json`, this.baseUrl);
    const response = await this.fetchImpl(manifestUrl, {
      cache: 'no-store',
      credentials: 'omit',
      mode: manifestUrl.origin === location.origin ? 'same-origin' : 'cors',
    });
    if (!response.ok) throw networkError(response.status);
    const manifest = parseReleaseManifest(await response.json());
    if (manifest.releaseId !== channel.releaseId) throw new ResourceInstallerError('manifest');
    return { channel: channel as ReleaseChannel, manifest };
  }

  assetUrl(releaseId: string, path: string): URL {
    if (!isSafePath(path)) throw new ResourceInstallerError('manifest');
    return new URL(`r/${encodeURIComponent(releaseId)}/${path}`, this.baseUrl);
  }
}

export class ResourcePlanner {
  constructor(
    private readonly manifest: ReleaseManifestV3,
    private readonly installed: ReadonlyMap<string, JournalAsset>,
  ) {}

  create(request: ResourcePlanRequest): { plan: ResourcePlan; assets: ReleaseAsset[] } {
    const profiles = profilesForRequest(request);
    const selectedPaths = new Set(profiles.flatMap((profile) => this.manifest.profiles[profile] || []));
    const assets = this.manifest.assets.filter((asset) => selectedPaths.has(asset.path));
    const totalBytes = assets.reduce((total, asset) => total + asset.bytes, 0);
    const reusedBytes = assets.reduce((total, asset) => {
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
  readiness: 'checking',
  phase: 'idle',
  storageMode,
  currentChunk: null,
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
      if (event.data?.type === 'resource-snapshot') this.snapshot = event.data.snapshot;
      this.publish();
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
    this.snapshot = {
      ...this.snapshot,
      installedRelease: active?.releaseId || null,
      installedProfiles: active?.installedProfiles || [],
      targetRelease: release.manifest.releaseId,
      availableRelease: release.manifest.releaseId,
      readiness:
        active?.releaseId === release.manifest.releaseId ? 'ready' : active ? 'update-available' : 'needs-download',
    };
    this.publish();
  }

  async plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    this.patch({ phase: 'planning', errorCode: null });
    const manifest = await this.ensureManifest();
    const installed = new Map(
      (await this.options.journal.listAssets(manifest.releaseId)).map((item) => [item.path, item]),
    );
    const planned = new ResourcePlanner(manifest, installed).create(request);
    this.plans.set(planned.plan.planId, { manifest, assets: planned.assets, plan: planned.plan });
    this.patch({ phase: 'idle' });
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
    this.currentManifest = manifest;
    this.patch({
      availableRelease: manifest.releaseId,
      targetRelease: manifest.releaseId,
      readiness:
        this.snapshot.installedRelease === manifest.releaseId
          ? 'ready'
          : this.snapshot.installedRelease
            ? 'update-available'
            : 'needs-download',
    });
  }

  async checkHealth(): Promise<void> {
    const manifest = await this.ensureManifest();
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

  private async applyUnlocked(internal: {
    manifest: ReleaseManifestV3;
    assets: ReleaseAsset[];
    plan: ResourcePlan;
  }): Promise<void> {
    const existing = new Map(
      (await this.options.journal.listAssets(internal.manifest.releaseId)).map((asset) => [asset.path, asset]),
    );
    const pending = internal.assets.filter((asset) => existing.get(asset.path)?.sha256 !== asset.sha256);
    this.abortController = new AbortController();
    const startedAt = performance.now();
    this.patch({
      targetRelease: internal.manifest.releaseId,
      readiness: 'updating',
      phase: 'downloading',
      currentChunk: pending[0]?.chunk || null,
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
        this.patch({ currentChunk: asset.chunk });
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
        readiness: 'ready',
        canRetry: false,
      });
      this.cleanupLegacyState();
    } finally {
      this.abortController = null;
    }
  }

  private retryDelays(): number[] {
    return this.options.retryDelaysMs || [1_000, 3_000, 10_000];
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
            .filter((key) => key.startsWith('onlyoffice-browser-font-packages-'))
            .map((key) => this.options.cacheStorage!.delete(key)),
        ),
      );
  }

  private publish(): void {
    const snapshot = this.getInstallerSnapshot();
    for (const listener of this.listeners) listener(snapshot);
    this.options.broadcast?.postMessage({ type: 'resource-snapshot', snapshot });
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
