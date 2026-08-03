import {
  CanonicalResourceStore,
  CanonicalResourceStoreError,
  IndexedDbCanonicalResourceJournal,
  planCanonicalPackageTransfer,
  type CanonicalObjectProgress,
  type CanonicalPackageTransport,
  type CanonicalResourceJournal,
} from './canonical-resource-store';
import { releaseManifestV5ToContentModel } from './release-content-manifest';
import {
  planReleaseContent,
  ReleaseContentModelError,
  type ContentObjectDescriptor,
  type ContentObjectRecord,
  type ReleaseContentModel,
  type ReleaseContentPlan,
} from './release-content-model';
import {
  ReleaseRepository,
  ResourceInstallerError,
  parseRequiredReleaseIdentity,
  requiredReleaseIdentitiesEqual,
  type FailedResource,
  type OfficeRuntimeResourceInstaller,
  type ReleaseManifestV5,
  type RequiredReleaseIdentity,
  type ResourceErrorCode,
  type ResourceInstallerSnapshot,
  type ResourcePhase,
  type ResourcePlan,
  type ResourcePlanRequest,
  type ResourceProfile,
} from './release-resources';

const ALL_PROFILES: ResourceProfile[] = ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'];

type CurrentV5Release = {
  manifest: ReleaseManifestV5;
  manifestSha256: string;
  model: ReleaseContentModel;
};

type PreparedCanonicalPlan = {
  publicPlan: ResourcePlan;
  contentPlan: ReleaseContentPlan;
  release: CurrentV5Release;
  packageTransport?: CanonicalPackageTransport;
  transferObjects: ContentObjectDescriptor[];
  verifyBytes: number;
};

type ProgressContext = {
  startedAt: number;
  contentPlan: ReleaseContentPlan;
  transferObjects: ContentObjectDescriptor[];
  objectIndexes: Map<string, number>;
  maxLoadedByDigest: Map<string, number>;
  verifiedDigests: Set<string>;
  currentDigest: string | null;
  networkBytes: number;
};

export type CanonicalResourceInstallerOptions = {
  assetBaseUrl: string | URL;
  fetch: typeof fetch;
  cacheStorage?: CacheStorage;
  indexedDb?: IDBFactory;
  journal?: CanonicalResourceJournal;
  locks?: LockManager;
  broadcast?: BroadcastChannel;
  requiredReleaseIdentity?: RequiredReleaseIdentity;
  maxConcurrentDownloads?: number;
  timeoutMs?: number;
  retryDelaysMs?: number[];
  now?: () => number;
  performanceNow?: () => number;
  online?: () => boolean;
};

function createInitialSnapshot(): ResourceInstallerSnapshot {
  return {
    installedRelease: null,
    targetRelease: null,
    availableRelease: null,
    availablePackageVersion: null,
    readiness: 'checking',
    phase: 'idle',
    storageMode: 'cache-storage',
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
  };
}

function cloneSnapshot(snapshot: ResourceInstallerSnapshot): ResourceInstallerSnapshot {
  return {
    ...snapshot,
    installedProfiles: [...snapshot.installedProfiles],
    failedResources: snapshot.failedResources.map((failure) => ({ ...failure })),
  };
}

function planId(): string {
  return crypto.randomUUID();
}

function isOffline(): boolean {
  return typeof navigator !== 'undefined' && navigator.onLine === false;
}

function defaultDownloadConcurrency(): number {
  if (typeof navigator === 'undefined') return 4;
  const connection = (
    navigator as Navigator & {
      connection?: { saveData?: boolean; effectiveType?: string };
    }
  ).connection;
  return connection?.saveData === true || connection?.effectiveType === 'slow-2g' || connection?.effectiveType === '2g'
    ? 2
    : 4;
}

function hasQuotaCause(error: unknown): boolean {
  let current: unknown = error;
  for (let depth = 0; depth < 4 && current instanceof Error; depth += 1) {
    if (current instanceof DOMException && current.name === 'QuotaExceededError') return true;
    current = current.cause;
  }
  return false;
}

function errorCode(error: unknown): ResourceErrorCode {
  if (error instanceof ResourceInstallerError) return error.code;
  if (error instanceof CanonicalResourceStoreError) {
    if (hasQuotaCause(error)) return 'quota';
    if (error.code === 'network' || error.code === 'http') return isOffline() ? 'offline' : 'network';
    if (error.code === 'integrity' || error.code === 'invalid-content' || error.code === 'incomplete') {
      return 'integrity';
    }
    if (error.code === 'invalid-state') return 'incompatible';
    return 'storage';
  }
  if (error instanceof ReleaseContentModelError) {
    if (error.code === 'integrity-conflict' || error.code === 'incomplete') return 'integrity';
    if (error.code === 'invalid-state') return 'incompatible';
    return 'manifest';
  }
  if (error instanceof DOMException) {
    if (error.name === 'QuotaExceededError') return 'quota';
    if (error.name === 'AbortError') return 'aborted';
  }
  return isOffline() ? 'offline' : 'network';
}

function abortableDelay(delayMs: number, signal: AbortSignal): Promise<void> {
  if (delayMs <= 0) return Promise.resolve();
  return new Promise((resolve, reject) => {
    if (signal.aborted) {
      reject(new DOMException('The operation was aborted', 'AbortError'));
      return;
    }
    const timeout = setTimeout(resolve, delayMs);
    signal.addEventListener(
      'abort',
      () => {
        clearTimeout(timeout);
        reject(new DOMException('The operation was aborted', 'AbortError'));
      },
      { once: true },
    );
  });
}

function failedObjects(plan: ReleaseContentPlan, code: ResourceErrorCode, attempts: number): FailedResource[] {
  const objects = plan.missingObjects.length ? plan.missingObjects : plan.requiredObjects.slice(0, 1);
  return objects.map((object) => ({
    path: `objects/sha256/${object.sha256}`,
    code,
    attempts,
  }));
}

function packageTransportForRelease(release: CurrentV5Release): CanonicalPackageTransport {
  return {
    bytes: release.manifest.package.bytes,
    headerBytes: release.manifest.package.headerBytes,
    segments: release.manifest.package.segments.map((segment) => ({
      offset: segment.offset,
      bytes: segment.bytes,
      sha256: segment.sha256,
    })),
    assets: release.manifest.assets.map((asset) => ({
      path: asset.path,
      packageOffset: asset.packageOffset,
      bytes: asset.bytes,
      sha256: asset.sha256,
      mime: asset.mime,
    })),
  };
}

/**
 * Production v5 installer for the canonical single-copy resource store.
 *
 * Every public scope is deliberately presented as one all-in-one installation.
 * The content planner still chooses whole-file, FastCDC, or package spans per
 * asset, so incremental releases download only missing immutable objects.
 */
export class CanonicalResourceInstaller implements OfficeRuntimeResourceInstaller {
  private readonly repository: ReleaseRepository;
  private readonly journal: CanonicalResourceJournal;
  private readonly store: CanonicalResourceStore;
  private readonly listeners = new Set<(snapshot: ResourceInstallerSnapshot) => void>();
  private readonly plans = new Map<string, PreparedCanonicalPlan>();
  private readonly releases = new Map<string, CurrentV5Release>();
  private readonly installedPaths = new Set<string>();
  private readonly retryDelaysMs: number[];
  private readonly timeoutMs: number;
  private readonly performanceNow: () => number;
  private readonly online: () => boolean;
  private readonly broadcast?: BroadcastChannel;
  private readonly requiredReleaseIdentity: RequiredReleaseIdentity | null;
  private readonly instanceId = planId();
  private snapshot = createInitialSnapshot();
  private currentRelease: CurrentV5Release | null = null;
  private lastPlan: PreparedCanonicalPlan | null = null;
  private operation: Promise<void> | null = null;
  private attemptController: AbortController | null = null;
  private idleTimeout: ReturnType<typeof setTimeout> | null = null;
  private interruption: 'pause' | 'cancel' | 'timeout' | null = null;
  private attemptInterruption: 'pause' | 'cancel' | 'timeout' | null = null;
  private resumeGate: Promise<void> | null = null;
  private resolveResumeGate: (() => void) | null = null;
  private progress: ProgressContext | null = null;
  private operationKind: 'apply' | 'repair' | null = null;
  private lastOperationKind: 'apply' | 'repair' | null = null;
  private lastRepairOptions: { scope: 'required' | 'installed' | 'all' } | null = null;
  private pausedResumePhase: ResourcePhase = 'downloading';

  constructor(options: CanonicalResourceInstallerOptions) {
    const cacheStorage = options.cacheStorage ?? (typeof caches === 'undefined' ? undefined : caches);
    if (!cacheStorage) throw new ResourceInstallerError('storage');
    const indexedDb = options.indexedDb ?? (typeof indexedDB === 'undefined' ? undefined : indexedDB);
    if (!options.journal && !indexedDb) throw new ResourceInstallerError('storage');

    this.repository = new ReleaseRepository(options.assetBaseUrl, options.fetch);
    this.requiredReleaseIdentity = options.requiredReleaseIdentity
      ? parseRequiredReleaseIdentity(options.requiredReleaseIdentity)
      : null;
    if (options.requiredReleaseIdentity && !this.requiredReleaseIdentity) {
      throw new ResourceInstallerError('incompatible');
    }
    if (this.requiredReleaseIdentity) {
      this.snapshot.targetRelease = this.requiredReleaseIdentity.releaseId;
    }
    this.journal =
      options.journal ||
      new IndexedDbCanonicalResourceJournal(indexedDb!, {
        locks: options.locks,
        now: options.now,
      });
    this.retryDelaysMs = options.retryDelaysMs || [1_000, 3_000, 10_000];
    this.timeoutMs = options.timeoutMs || 30_000;
    this.performanceNow = options.performanceNow || (() => performance.now());
    this.online = options.online || (() => typeof navigator === 'undefined' || navigator.onLine !== false);
    this.broadcast = options.broadcast;
    this.broadcast?.addEventListener('message', (event) => {
      const message = event.data as {
        type?: unknown;
        sourceId?: unknown;
        snapshot?: unknown;
        requiredReleaseIdentity?: unknown;
      };
      if (
        message?.type !== 'canonical-resource-snapshot-v1' ||
        message.sourceId === this.instanceId ||
        this.operation ||
        !this.isSnapshot(message.snapshot)
      ) {
        return;
      }
      const messageIdentity = parseRequiredReleaseIdentity(message.requiredReleaseIdentity);
      if (
        this.requiredReleaseIdentity &&
        (!messageIdentity ||
          !requiredReleaseIdentitiesEqual(this.requiredReleaseIdentity, messageIdentity) ||
          message.snapshot.targetRelease !== this.requiredReleaseIdentity.releaseId)
      ) {
        return;
      }
      void this.acceptBroadcastSnapshot(message.snapshot).catch(() => undefined);
    });
    const cacheKeyOrigin = new URL('/', options.assetBaseUrl).origin;
    this.store = new CanonicalResourceStore({
      cacheStorage,
      journal: this.journal,
      fetch: options.fetch,
      locks: options.locks,
      cacheKeyOrigin,
      maxConcurrentDownloads: options.maxConcurrentDownloads ?? defaultDownloadConcurrency(),
      now: options.now,
      objectUrl: (releaseId, object) => this.repository.contentObjectUrl(releaseId, object.sha256),
      onObjectProgress: (progress) => this.onObjectProgress(progress),
      onObjectVerified: ({ releaseId, object }) => this.onObjectVerified(releaseId, object),
    });
  }

  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getInstallerSnapshot(): ResourceInstallerSnapshot {
    return cloneSnapshot(this.snapshot);
  }

  getInstalledPaths(): string[] {
    return [...this.installedPaths].sort();
  }

  private async acceptBroadcastSnapshot(snapshot: ResourceInstallerSnapshot): Promise<void> {
    const active = await this.journal.getActiveRelease();
    const activeReleaseId = active?.releaseId ?? null;
    if (this.operation || activeReleaseId !== snapshot.installedRelease) return;
    if (
      snapshot.phase === 'idle' &&
      ((snapshot.targetRelease !== null && snapshot.targetRelease !== activeReleaseId) ||
        (snapshot.availableRelease !== null && snapshot.availableRelease !== activeReleaseId))
    ) {
      return;
    }
    this.snapshot = cloneSnapshot(snapshot);
    for (const listener of this.listeners) listener(cloneSnapshot(this.snapshot));
    if (snapshot.installedRelease) {
      await this.refreshInstalledPaths(snapshot.installedRelease);
    } else {
      this.installedPaths.clear();
    }
  }

  async initialize(): Promise<void> {
    this.patch({ readiness: 'checking', phase: 'idle', errorCode: null });
    let targetError: unknown;
    const existingActiveRelease = await this.journal.getActiveRelease();
    const canBootstrapFromExisting =
      !!existingActiveRelease &&
      (!this.requiredReleaseIdentity || this.activeMatchesRequiredIdentity(existingActiveRelease));
    const online = this.online();
    if (!canBootstrapFromExisting && online && this.requiredReleaseIdentity) {
      try {
        await this.refreshAvailableRelease();
      } catch {
        // The mutable channel is display-only for a pinned Piwork release.
      }
    }
    if (!canBootstrapFromExisting && online) {
      try {
        await this.loadCurrentRelease(true);
      } catch (error) {
        targetError = error;
      }
    } else if (!canBootstrapFromExisting) {
      targetError = new ResourceInstallerError('offline');
    }
    try {
      await this.refreshHealth(true);
    } catch (error) {
      this.setFailure(error, undefined, 1);
      throw this.asInstallerError(error);
    }
    if (targetError) {
      const active = await this.journal.getActiveRelease();
      const activeIsUsable = !!active && (!this.requiredReleaseIdentity || this.activeMatchesRequiredIdentity(active));
      if (!activeIsUsable) {
        const failure =
          active &&
          this.requiredReleaseIdentity &&
          active.releaseId === this.requiredReleaseIdentity.releaseId &&
          active.manifestSha256 !== this.requiredReleaseIdentity.manifestSha256
            ? new ResourceInstallerError('incompatible', 'release/identity')
            : targetError;
        this.setFailure(failure, undefined, 1);
        throw this.asInstallerError(failure);
      }
    }
  }

  async plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    this.patch({ phase: 'planning', errorCode: null, failedResources: [] });
    try {
      const release = await this.loadCurrentRelease(false);
      const prepared = await this.createPlan(release, request.scope);
      this.plans.set(prepared.publicPlan.planId, prepared);
      this.patch({
        phase: 'idle',
        targetRelease: release.manifest.releaseId,
        downloadedBytes: 0,
        downloadBytes: prepared.publicPlan.downloadBytes,
        verifiedBytes: 0,
        verifyBytes: prepared.verifyBytes,
        currentChunk: null,
        currentChunkIndex: 0,
        currentChunkCount: prepared.transferObjects.length,
        canPause: false,
        canResume: false,
        canRetry: false,
      });
      return {
        ...prepared.publicPlan,
        profiles: [...prepared.publicPlan.profiles],
      };
    } catch (error) {
      try {
        await this.refreshHealth(true);
        const active = await this.journal.getActiveRelease();
        if (!active || (this.requiredReleaseIdentity && !this.activeMatchesRequiredIdentity(active))) {
          this.setFailure(error, undefined, 1);
        }
      } catch {
        this.setFailure(error, undefined, 1);
      }
      throw this.asInstallerError(error);
    }
  }

  async apply(plan: ResourcePlan): Promise<void> {
    if (this.requiredReleaseIdentity && plan.releaseId !== this.requiredReleaseIdentity.releaseId) {
      throw new ResourceInstallerError('incompatible');
    }
    const prepared = this.plans.get(plan.planId);
    if (!prepared || !this.samePlan(prepared.publicPlan, plan)) {
      throw new ResourceInstallerError('incompatible');
    }
    if (this.operation) throw new ResourceInstallerError('incompatible');
    this.lastPlan = prepared;
    await this.startOperation(prepared);
  }

  async checkForUpdates(): Promise<void> {
    try {
      if (!this.online()) {
        await this.refreshHealth(true);
        return;
      }
      if (this.requiredReleaseIdentity) {
        try {
          await this.refreshAvailableRelease();
        } catch {
          // Upstream availability never changes the authorized Piwork target.
        }
      }
      await this.loadCurrentRelease(true);
      await this.refreshHealth(true);
    } catch (error) {
      try {
        await this.refreshHealth(true);
        if (await this.journal.getActiveRelease()) return;
      } catch {
        // Preserve the original availability error when local health also fails.
      }
      this.setFailure(error, undefined, 1);
      throw this.asInstallerError(error);
    }
  }

  async checkHealth(): Promise<void> {
    try {
      await this.refreshHealth(true);
    } catch (error) {
      this.setFailure(error, undefined, 1);
      throw this.asInstallerError(error);
    }
  }

  async repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void> {
    if (this.operation) throw new ResourceInstallerError('incompatible');
    this.lastRepairOptions = { ...options };
    this.lastOperationKind = 'repair';
    await this.startTrackedOperation('repair', () => this.runRepair(options));
  }

  pause(): void {
    if (!this.operation || !this.snapshot.canPause || this.interruption === 'pause') return;
    this.pausedResumePhase = this.snapshot.phase === 'paused' ? 'downloading' : this.snapshot.phase;
    this.interruption = 'pause';
    if (!this.resumeGate) {
      this.resumeGate = new Promise((resolve) => {
        this.resolveResumeGate = resolve;
      });
    }
    this.clearIdleTimeout();
    this.attemptInterruption = 'pause';
    this.attemptController?.abort();
    this.patch({
      phase: 'paused',
      readiness: 'paused',
      canPause: false,
      canResume: true,
      canRetry: false,
      errorCode: null,
    });
  }

  async resume(): Promise<void> {
    if (this.interruption === 'pause') {
      this.interruption = null;
      this.resolveResumeGate?.();
      this.resolveResumeGate = null;
      this.resumeGate = null;
      this.patch({
        phase: this.pausedResumePhase,
        readiness: 'updating',
        canPause: true,
        canResume: false,
      });
      if (this.operation) await this.operation;
      return;
    }
    if (!this.operation && this.snapshot.canRetry) {
      if (this.lastOperationKind === 'repair' && this.lastRepairOptions) {
        await this.startTrackedOperation('repair', () => this.runRepair(this.lastRepairOptions!));
      } else if (this.lastPlan) {
        await this.startOperation(this.lastPlan);
      }
    }
  }

  cancel(): void {
    const cancellingRepair = this.operationKind === 'repair';
    this.interruption = 'cancel';
    this.attemptInterruption = 'cancel';
    this.clearIdleTimeout();
    this.attemptController?.abort();
    this.resolveResumeGate?.();
    this.resolveResumeGate = null;
    this.resumeGate = null;
    this.patch({
      phase: 'idle',
      readiness: cancellingRepair
        ? this.snapshot.installedRelease
          ? 'repair-needed'
          : 'needs-download'
        : this.snapshot.installedRelease
          ? 'update-available'
          : 'needs-download',
      canPause: false,
      canResume: false,
      canRetry: true,
      errorCode: 'aborted',
    });
  }

  async rollbackActivation(releaseId: string, failure: { code: ResourceErrorCode; path: string }): Promise<void> {
    if (this.operation) throw new ResourceInstallerError('incompatible', 'release/activation-in-progress');
    if (this.requiredReleaseIdentity && releaseId !== this.requiredReleaseIdentity.releaseId) {
      throw new ResourceInstallerError('incompatible', 'release/identity');
    }
    await this.store.rollbackActivation(releaseId, failure.code);
    await this.refreshHealth(true);
    const active = await this.journal.getActiveRelease();
    this.patch({
      targetRelease: this.requiredReleaseIdentity?.releaseId ?? releaseId,
      readiness: active ? 'update-available' : 'error',
      phase: 'idle',
      currentChunk: null,
      currentChunkIndex: 0,
      currentChunkCount: 0,
      failedResources: [
        {
          path: failure.path,
          code: failure.code,
          attempts: 1,
        },
      ],
      canPause: false,
      canResume: false,
      canRetry: true,
      errorCode: failure.code,
    });
  }

  private startOperation(prepared: PreparedCanonicalPlan): Promise<void> {
    this.lastOperationKind = 'apply';
    return this.startTrackedOperation('apply', () => this.runOperation(prepared));
  }

  private startTrackedOperation(kind: 'apply' | 'repair', run: () => Promise<void>): Promise<void> {
    let operation!: Promise<void>;
    this.operationKind = kind;
    operation = run().finally(() => {
      if (this.operation === operation) this.operation = null;
      if (this.operationKind === kind) this.operationKind = null;
      this.attemptController = null;
      this.clearIdleTimeout();
      this.progress = null;
      this.pausedResumePhase = 'downloading';
    });
    this.operation = operation;
    return operation;
  }

  private async runRepair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void> {
    this.interruption = null;
    const active = await this.journal.getActiveRelease();
    if (active) {
      this.patch({
        installedRelease: active.releaseId,
        installedProfiles: [...ALL_PROFILES],
      });
    }
    if (
      active &&
      this.requiredReleaseIdentity &&
      active.releaseId === this.requiredReleaseIdentity.releaseId &&
      active.manifestSha256 !== this.requiredReleaseIdentity.manifestSha256
    ) {
      const failure = new ResourceInstallerError('incompatible', 'release/identity');
      this.setFailure(failure, undefined, 1);
      throw failure;
    }

    if (active && (!this.requiredReleaseIdentity || this.activeMatchesRequiredIdentity(active))) {
      let verification: Awaited<ReturnType<CanonicalResourceStore['verifyReleaseIntegrity']>>;
      try {
        verification = await this.verifyActiveRelease(active);
      } catch (error) {
        const failure = this.asInstallerError(error);
        if (!(failure.code === 'aborted' && this.hasInterruption('cancel'))) {
          this.setFailure(failure, undefined, 1);
        }
        throw failure;
      }
      if (verification.ready) {
        await this.refreshHealth(false);
        return;
      }
      const failures: FailedResource[] = verification.failures.map((failure) => ({
        path: `objects/sha256/${failure.object.sha256}`,
        code: failure.code === 'integrity' ? 'integrity' : 'storage',
        attempts: 1,
      }));
      if (failures.length === 0) {
        const failure = new ResourceInstallerError('storage', `releases/${active.releaseId}`);
        this.setFailure(failure, undefined, 1);
        throw failure;
      }
      this.patch({
        phase: 'repairing',
        readiness: 'updating',
        failedResources: failures,
        errorCode: failures.some((failure) => failure.code === 'integrity') ? 'integrity' : 'storage',
        canPause: false,
        canResume: false,
        canRetry: false,
      });
    }

    if (this.interruption === 'cancel') {
      throw this.cancelledRepair(active !== null);
    }
    if (this.interruption === 'pause') await this.waitForResume();

    let release: CurrentV5Release;
    try {
      release =
        active && (!this.requiredReleaseIdentity || active.releaseId === this.requiredReleaseIdentity.releaseId)
          ? await this.loadRelease(active.releaseId, active.manifestSha256)
          : await this.loadCurrentRelease(false);
    } catch (error) {
      this.setFailure(error, undefined, 1);
      throw this.asInstallerError(error);
    }

    if (this.interruption === 'cancel') {
      throw this.cancelledRepair(active !== null);
    }
    this.patch({ phase: 'repairing', readiness: 'updating', canPause: false, errorCode: null });
    let prepared: PreparedCanonicalPlan;
    try {
      prepared = await this.createPlan(release, 'repair');
    } catch (error) {
      this.setFailure(error, undefined, 1);
      throw this.asInstallerError(error);
    }
    this.plans.set(prepared.publicPlan.planId, prepared);
    this.lastPlan = prepared;
    if (this.interruption === 'cancel') throw this.cancelledRepair(active !== null);
    await this.runOperation(prepared, true);
    await this.refreshHealth(true);
    if (this.snapshot.readiness === 'repair-needed' || this.snapshot.readiness === 'error') {
      const first = this.snapshot.failedResources[0];
      throw new ResourceInstallerError(first?.code || this.snapshot.errorCode || 'storage', first?.path);
    }

    void options.scope;
  }

  private async verifyActiveRelease(
    active: NonNullable<Awaited<ReturnType<CanonicalResourceJournal['getActiveRelease']>>>,
  ): ReturnType<CanonicalResourceStore['verifyReleaseIntegrity']> {
    const verifyBytes = active.requiredObjects.reduce((sum, object) => sum + object.bytes, 0);
    while (true) {
      if (this.interruption === 'cancel') throw new ResourceInstallerError('aborted');
      if (this.interruption === 'pause') await this.waitForResume();
      const attempt = new AbortController();
      this.attemptController = attempt;
      this.attemptInterruption = null;
      this.patch({
        phase: 'repairing',
        readiness: 'updating',
        currentChunk: null,
        currentChunkIndex: 0,
        currentChunkCount: active.requiredObjects.length,
        verifiedBytes: 0,
        verifyBytes,
        failedResources: [],
        canPause: active.requiredObjects.length > 0,
        canResume: false,
        canRetry: false,
        errorCode: null,
      });
      const verification = await this.store.verifyReleaseIntegrity({
        releaseId: active.releaseId,
        signal: attempt.signal,
      });
      this.attemptController = null;
      const attemptInterruption = this.attemptInterruption;
      this.attemptInterruption = null;
      if (verification.status === 'aborted') {
        if (attemptInterruption === 'pause') {
          if (this.interruption === 'pause') await this.waitForResume();
          continue;
        }
        if (attemptInterruption === 'cancel' || this.hasInterruption('cancel')) {
          throw new ResourceInstallerError('aborted');
        }
        throw new ResourceInstallerError('storage', `releases/${active.releaseId}`);
      }
      this.patch({
        phase: 'repairing',
        verifiedBytes: verification.verifiedBytes,
        verifyBytes,
        canPause: false,
      });
      return verification;
    }
  }

  private async runOperation(prepared: PreparedCanonicalPlan, preserveInterruption = false): Promise<void> {
    const { contentPlan, publicPlan, release, packageTransport, transferObjects, verifyBytes } = prepared;
    if (!preserveInterruption) this.interruption = null;
    this.progress = {
      startedAt: this.performanceNow(),
      contentPlan,
      transferObjects,
      objectIndexes: new Map(transferObjects.map((object, index) => [object.sha256, index + 1])),
      maxLoadedByDigest: new Map(),
      verifiedDigests: new Set(),
      currentDigest: null,
      networkBytes: 0,
    };
    this.patch({
      targetRelease: release.manifest.releaseId,
      readiness: 'updating',
      phase: transferObjects.length ? 'downloading' : 'verifying',
      currentChunk: transferObjects[0]?.sha256 || null,
      currentChunkIndex: transferObjects.length ? 1 : 0,
      currentChunkCount: transferObjects.length,
      downloadedBytes: 0,
      downloadBytes: publicPlan.downloadBytes,
      verifiedBytes: 0,
      verifyBytes,
      bytesPerSecond: 0,
      failedResources: [],
      canPause: contentPlan.missingObjects.length > 0,
      canResume: false,
      canRetry: false,
      errorCode: null,
    });

    let retryIndex = 0;
    while (true) {
      if (this.interruption === 'cancel') throw new ResourceInstallerError('aborted');
      if (this.interruption === 'pause') await this.waitForResume();
      const attempt = new AbortController();
      this.attemptController = attempt;
      this.attemptInterruption = null;
      let timeoutTriggered = false;
      try {
        if (retryIndex > 0) {
          await abortableDelay(this.retryDelaysMs[retryIndex - 1], attempt.signal);
        }
        this.armIdleTimeout(() => {
          timeoutTriggered = true;
          this.interruption = 'timeout';
          this.attemptInterruption = 'timeout';
          attempt.abort();
        });
        // A failed large HTTP/2 stream can leave an intermediary-specific
        // connection state behind. Keep the first URL canonical, then add a
        // deterministic retry query on subsequent installer attempts so the
        // edge opens a fresh response without creating another CAS key.
        await this.store.installAndActivate(release.model, attempt.signal, packageTransport, retryIndex);
        this.clearIdleTimeout();
        this.patch({
          phase: 'verifying',
          verifiedBytes: verifyBytes,
          downloadedBytes: publicPlan.downloadBytes,
          canPause: false,
        });
        this.patch({ phase: 'activating' });
        await this.refreshInstalledPaths(release.manifest.releaseId);
        const readiness = this.requiredReleaseIdentity
          ? 'ready'
          : this.snapshot.availableRelease && this.snapshot.availableRelease !== release.manifest.releaseId
            ? 'update-available'
            : 'ready';
        this.patch({
          installedRelease: release.manifest.releaseId,
          installedProfiles: [...ALL_PROFILES],
          targetRelease:
            this.requiredReleaseIdentity?.releaseId || this.snapshot.availableRelease || release.manifest.releaseId,
          readiness,
          phase: 'idle',
          currentChunk: null,
          currentChunkIndex: 0,
          currentChunkCount: 0,
          verifiedBytes: verifyBytes,
          downloadedBytes: publicPlan.downloadBytes,
          canPause: false,
          canResume: false,
          canRetry: false,
          errorCode: null,
          failedResources: [],
        });
        return;
      } catch (error) {
        this.clearIdleTimeout();
        this.attemptController = null;
        const attemptInterruption = this.attemptInterruption;
        this.attemptInterruption = null;
        if (attemptInterruption === 'pause') {
          if (this.interruption === 'pause') {
            this.patch({
              phase: 'paused',
              readiness: 'paused',
              canPause: false,
              canResume: true,
              canRetry: false,
            });
            await this.waitForResume();
          }
          continue;
        }
        if (attemptInterruption === 'cancel' || this.hasInterruption('cancel')) {
          const cancelled = new ResourceInstallerError('aborted');
          this.setFailure(cancelled, contentPlan, retryIndex + 1, false);
          throw cancelled;
        }
        const code: ResourceErrorCode =
          timeoutTriggered || attemptInterruption === 'timeout' || this.interruption === 'timeout'
            ? 'timeout'
            : errorCode(error);
        if (this.interruption === 'timeout') this.interruption = null;
        if ((code === 'network' || code === 'timeout') && retryIndex < this.retryDelaysMs.length) {
          retryIndex += 1;
          continue;
        }
        const failure = new ResourceInstallerError(code);
        this.setFailure(failure, contentPlan, retryIndex + 1);
        throw failure;
      }
    }
  }

  private async waitForResume(): Promise<void> {
    if (this.interruption !== 'pause') return;
    if (!this.resumeGate) {
      this.resumeGate = new Promise((resolve) => {
        this.resolveResumeGate = resolve;
      });
    }
    await this.resumeGate;
  }

  private async createPlan(
    release: CurrentV5Release,
    scope: ResourcePlanRequest['scope'],
  ): Promise<PreparedCanonicalPlan> {
    const inventory = await this.healthyInventory();
    const active = await this.journal.getActiveRelease();
    const previousMappings = active ? await this.journal.listAssetMappings(active.releaseId) : [];
    const packageTransport =
      !active && release.manifest.assets.some((asset) => asset.representations.fastcdc)
        ? packageTransportForRelease(release)
        : undefined;
    const contentPlan = planReleaseContent(release.model, inventory, previousMappings, {
      preferCanonicalForColdInstall: Boolean(packageTransport),
    });
    const transferObjects = packageTransport
      ? planCanonicalPackageTransfer(release.model, packageTransport, contentPlan)
      : contentPlan.missingObjects;
    const downloadBytes = transferObjects.reduce((sum, object) => sum + object.bytes, 0);
    const verifyBytes = contentPlan.missingObjects.reduce((sum, object) => sum + object.bytes, 0);
    const publicPlan: ResourcePlan = {
      planId: planId(),
      releaseId: release.manifest.releaseId,
      scope,
      profiles: [...ALL_PROFILES],
      totalBytes: contentPlan.requiredObjectBytes,
      downloadBytes,
      reusedBytes: contentPlan.reusedObjectBytes,
    };
    return {
      publicPlan,
      contentPlan,
      release,
      ...(packageTransport ? { packageTransport } : {}),
      transferObjects,
      verifyBytes,
    };
  }

  private async healthyInventory(): Promise<ContentObjectRecord[]> {
    const records = await this.journal.listObjects();
    const healthy: ContentObjectRecord[] = [];
    const concurrency = 16;
    let cursor = 0;
    const worker = async () => {
      while (cursor < records.length) {
        const record = records[cursor++];
        const response = await this.store.matchObject(record);
        if (!response) continue;
        healthy.push(record);
        await response.body?.cancel().catch(() => undefined);
      }
    };
    await Promise.all(Array.from({ length: Math.min(concurrency, records.length) }, () => worker()));
    return healthy;
  }

  private async loadCurrentRelease(force: boolean): Promise<CurrentV5Release> {
    if (this.currentRelease && !force) return this.currentRelease;
    const required = this.requiredReleaseIdentity;
    const current = required
      ? await this.repository.releaseV5(required.releaseId, required.manifestSha256)
      : await this.repository.currentV5();
    if (
      required &&
      (current.manifest.releaseId !== required.releaseId ||
        current.manifestSha256 !== required.manifestSha256 ||
        current.manifest.packageVersion !== required.packageVersion ||
        current.manifest.hostBuildId !== required.hostBuildId)
    ) {
      throw new ResourceInstallerError('incompatible');
    }
    const release: CurrentV5Release = {
      manifest: current.manifest,
      manifestSha256: current.manifestSha256,
      model: releaseManifestV5ToContentModel(current.manifest, current.manifestSha256),
    };
    this.currentRelease = release;
    this.releases.set(release.manifest.releaseId, release);
    this.patch({
      targetRelease: release.manifest.releaseId,
      ...(!required
        ? {
            availableRelease: release.manifest.releaseId,
            availablePackageVersion: release.manifest.packageVersion,
          }
        : {}),
    });
    return release;
  }

  private async refreshAvailableRelease(): Promise<void> {
    const upstream = await this.repository.currentV5();
    this.patch({
      availableRelease: upstream.manifest.releaseId,
      availablePackageVersion: upstream.manifest.packageVersion,
      targetRelease: this.requiredReleaseIdentity?.releaseId ?? upstream.manifest.releaseId,
    });
  }

  private async loadRelease(releaseId: string, manifestSha256: string): Promise<CurrentV5Release> {
    if (
      this.requiredReleaseIdentity &&
      (releaseId !== this.requiredReleaseIdentity.releaseId ||
        manifestSha256 !== this.requiredReleaseIdentity.manifestSha256)
    ) {
      throw new ResourceInstallerError('incompatible');
    }
    const cached = this.releases.get(releaseId);
    if (cached) {
      if (cached.manifestSha256 !== manifestSha256) throw new ResourceInstallerError('integrity');
      return cached;
    }
    const loaded = await this.repository.releaseV5(releaseId, manifestSha256);
    if (
      this.requiredReleaseIdentity &&
      (loaded.manifest.packageVersion !== this.requiredReleaseIdentity.packageVersion ||
        loaded.manifest.hostBuildId !== this.requiredReleaseIdentity.hostBuildId)
    ) {
      throw new ResourceInstallerError('incompatible');
    }
    const release: CurrentV5Release = {
      manifest: loaded.manifest,
      manifestSha256: loaded.manifestSha256,
      model: releaseManifestV5ToContentModel(loaded.manifest, loaded.manifestSha256),
    };
    this.releases.set(releaseId, release);
    return release;
  }

  private async refreshHealth(probe: boolean): Promise<void> {
    const active = await this.journal.getActiveRelease();
    const health = await this.store.checkHealth({ releaseId: active?.releaseId, probe });
    if (!active) {
      this.installedPaths.clear();
      this.patch({
        installedRelease: null,
        installedProfiles: [],
        readiness: 'needs-download',
        phase: 'idle',
        failedResources: [],
        errorCode: null,
        canPause: false,
        canResume: false,
      });
      return;
    }
    await this.refreshInstalledPaths(active.releaseId);
    const failures: FailedResource[] = health.missingObjects.map((object) => ({
      path: `objects/sha256/${object.sha256}`,
      code: 'storage' as const,
      attempts: 1,
    }));
    if (!health.probeSucceeded && failures.length === 0) {
      failures.push({ path: 'broker/probe', code: 'storage', attempts: 1 });
    }
    const identityMatches = !this.requiredReleaseIdentity || this.activeMatchesRequiredIdentity(active);
    if (health.ready && !identityMatches) {
      failures.push({
        path: 'release/identity',
        code: 'incompatible',
        attempts: 1,
      });
    }
    const targetIsInstalled = active.releaseId === this.requiredReleaseIdentity?.releaseId;
    const readiness =
      !health.ready || (targetIsInstalled && !identityMatches)
        ? 'repair-needed'
        : this.requiredReleaseIdentity
          ? identityMatches
            ? 'ready'
            : 'update-available'
          : this.snapshot.availableRelease && this.snapshot.availableRelease !== active.releaseId
            ? 'update-available'
            : 'ready';
    const healthyForTarget = health.ready && identityMatches;
    this.patch({
      installedRelease: active.releaseId,
      installedProfiles: [...ALL_PROFILES],
      targetRelease: this.requiredReleaseIdentity?.releaseId ?? this.snapshot.targetRelease,
      readiness,
      phase: 'idle',
      failedResources: failures,
      errorCode: healthyForTarget || readiness === 'update-available' ? null : failures[0]?.code || 'storage',
      canPause: false,
      canResume: false,
      canRetry: !healthyForTarget && readiness !== 'update-available',
    });
  }

  private async refreshInstalledPaths(releaseId: string): Promise<void> {
    const mappings = await this.journal.listAssetMappings(releaseId);
    this.installedPaths.clear();
    for (const mapping of mappings) this.installedPaths.add(mapping.path);
  }

  private activeMatchesRequiredIdentity(
    active: Awaited<ReturnType<CanonicalResourceJournal['getActiveRelease']>>,
  ): boolean {
    const required = this.requiredReleaseIdentity;
    return !required || (active?.releaseId === required.releaseId && active.manifestSha256 === required.manifestSha256);
  }

  private onObjectProgress(progress: CanonicalObjectProgress): void {
    const context = this.progress;
    if (!context || progress.releaseId !== context.contentPlan.releaseId) return;
    this.armIdleTimeout(() => {
      this.interruption = 'timeout';
      this.attemptInterruption = 'timeout';
      this.attemptController?.abort();
    });
    if (context.currentDigest && context.currentDigest !== progress.object.sha256) {
      const previousBytes =
        context.contentPlan.missingObjects.find((object) => object.sha256 === context.currentDigest)?.bytes || 0;
      this.patch({
        verifiedBytes: Math.min(this.snapshot.verifyBytes, this.snapshot.verifiedBytes + previousBytes),
      });
    }
    context.currentDigest = progress.object.sha256;
    context.networkBytes += progress.chunkBytes;
    context.maxLoadedByDigest.set(
      progress.object.sha256,
      Math.max(context.maxLoadedByDigest.get(progress.object.sha256) || 0, progress.loadedBytes),
    );
    const downloadedBytes = [...context.maxLoadedByDigest.values()].reduce((sum, bytes) => sum + bytes, 0);
    const seconds = Math.max((this.performanceNow() - context.startedAt) / 1_000, 0.001);
    this.patch({
      phase: progress.loadedBytes === progress.object.bytes ? 'verifying' : 'downloading',
      currentChunk: progress.object.sha256,
      currentChunkIndex: context.objectIndexes.get(progress.object.sha256) || 0,
      currentChunkCount: context.transferObjects.length,
      downloadedBytes: Math.min(this.snapshot.downloadBytes, downloadedBytes),
      bytesPerSecond: Math.round(context.networkBytes / seconds),
      canPause: progress.loadedBytes < progress.object.bytes,
    });
  }

  private onObjectVerified(releaseId: string, object: ContentObjectDescriptor): void {
    const context = this.progress;
    if (!context || releaseId !== context.contentPlan.releaseId || context.verifiedDigests.has(object.sha256)) return;
    context.verifiedDigests.add(object.sha256);
    const verifiedBytes = context.contentPlan.missingObjects.reduce(
      (sum, candidate) => sum + (context.verifiedDigests.has(candidate.sha256) ? candidate.bytes : 0),
      0,
    );
    this.patch({
      phase: 'verifying',
      currentChunk: object.sha256,
      currentChunkIndex: context.objectIndexes.get(object.sha256) || 0,
      verifiedBytes: Math.min(this.snapshot.verifyBytes, verifiedBytes),
    });
  }

  private armIdleTimeout(onTimeout: () => void): void {
    this.clearIdleTimeout();
    this.idleTimeout = setTimeout(onTimeout, this.timeoutMs);
  }

  private hasInterruption(interruption: 'pause' | 'cancel' | 'timeout'): boolean {
    return this.interruption === interruption;
  }

  private clearIdleTimeout(): void {
    if (this.idleTimeout) clearTimeout(this.idleTimeout);
    this.idleTimeout = null;
  }

  private cancelledRepair(hasActiveRelease: boolean): ResourceInstallerError {
    const failure = new ResourceInstallerError('aborted');
    this.patch({
      phase: 'idle',
      readiness: hasActiveRelease ? 'repair-needed' : 'needs-download',
      canPause: false,
      canResume: false,
      canRetry: true,
      errorCode: 'aborted',
    });
    return failure;
  }

  private samePlan(expected: ResourcePlan, received: ResourcePlan): boolean {
    return (
      expected.planId === received.planId &&
      expected.releaseId === received.releaseId &&
      expected.scope === received.scope &&
      expected.totalBytes === received.totalBytes &&
      expected.downloadBytes === received.downloadBytes &&
      expected.reusedBytes === received.reusedBytes &&
      expected.profiles.length === received.profiles.length &&
      expected.profiles.every((profile, index) => profile === received.profiles[index])
    );
  }

  private setFailure(error: unknown, contentPlan?: ReleaseContentPlan, attempts = 1, readinessError = true): void {
    const code = errorCode(error);
    const directFailure =
      error instanceof ResourceInstallerError && error.path ? [{ path: error.path, code, attempts }] : [];
    this.patch({
      phase: 'idle',
      readiness: readinessError ? 'error' : this.snapshot.readiness,
      failedResources: contentPlan ? failedObjects(contentPlan, code, attempts) : directFailure,
      canPause: false,
      canResume: false,
      canRetry: true,
      errorCode: code,
    });
  }

  private asInstallerError(error: unknown): ResourceInstallerError {
    return error instanceof ResourceInstallerError ? error : new ResourceInstallerError(errorCode(error));
  }

  private patch(patch: Partial<ResourceInstallerSnapshot>): void {
    this.snapshot = { ...this.snapshot, ...patch };
    const snapshot = cloneSnapshot(this.snapshot);
    for (const listener of this.listeners) listener(snapshot);
    try {
      this.broadcast?.postMessage({
        type: 'canonical-resource-snapshot-v1',
        sourceId: this.instanceId,
        snapshot,
        requiredReleaseIdentity: this.requiredReleaseIdentity ? { ...this.requiredReleaseIdentity } : null,
      });
    } catch {
      // Cross-tab progress is advisory; a closed channel must not fail installation.
    }
  }

  private isSnapshot(value: unknown): value is ResourceInstallerSnapshot {
    if (!value || typeof value !== 'object') return false;
    const snapshot = value as Partial<ResourceInstallerSnapshot>;
    return (
      snapshot.storageMode === 'cache-storage' &&
      typeof snapshot.readiness === 'string' &&
      typeof snapshot.phase === 'string' &&
      (snapshot.installedRelease === null || typeof snapshot.installedRelease === 'string') &&
      (snapshot.availableRelease === null || typeof snapshot.availableRelease === 'string') &&
      Array.isArray(snapshot.installedProfiles) &&
      Array.isArray(snapshot.failedResources)
    );
  }
}
