import {
  RuntimeCacheController,
  type RuntimeCacheCategory,
  type RuntimeCacheProgress,
  type RuntimeFontFamily,
} from './runtime-cache';
import { ONLYOFFICE_BROWSER_VERSION } from '../version';
import {
  createTransactionalResourceInstaller,
  type OfficeRuntimeResourceInstaller,
  type ResourceErrorCode,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
  type ResourcePlanRequest,
  type ResourceReadiness,
} from './release-resources';

export type OfficeRuntimeResourceOperation =
  | 'prepare-document'
  | 'prefetch-recommended'
  | 'load-all'
  | 'check-health'
  | 'download-font'
  | 'install-font-preset'
  | 'remove-font';

export type OfficeRuntimeReadiness = ResourceReadiness;
export type OfficeDocumentResourceType = 'word' | 'cell' | 'slide';
export type OfficeFontPreset = 'basic' | 'office-compatibility';

export type OfficeRuntimePackSnapshot = {
  id: RuntimeCacheCategory;
  ready: boolean;
  completedBytes: number;
  totalBytes: number;
};

export type OfficeRuntimeResourceSnapshot = {
  packageVersion: string;
  assetVersion: string;
  readiness: OfficeRuntimeReadiness;
  packs: OfficeRuntimePackSnapshot[];
  progress: RuntimeCacheProgress;
  fonts: RuntimeFontFamily[];
  verifiedFontPaths: string[];
  operation: OfficeRuntimeResourceOperation | null;
  error: { code: ResourceErrorCode; path?: string } | null;
  installedRelease: string | null;
  targetRelease: string | null;
  availableRelease: string | null;
  storageMode: 'cache-storage' | 'http-cache';
  phase: ResourceInstallerSnapshot['phase'];
  currentChunk: string | null;
  currentChunkIndex: number;
  currentChunkCount: number;
  downloadedBytes: number;
  downloadBytes: number;
  verifiedBytes: number;
  verifyBytes: number;
  bytesPerSecond: number;
  failedResources: ResourceInstallerSnapshot['failedResources'];
  canPause: boolean;
  canResume: boolean;
  canRetry: boolean;
};

export type OfficeRuntimeResourceListener = (snapshot: OfficeRuntimeResourceSnapshot) => void;

export type OfficeRuntimeResourceManagerOptions = {
  storage?: Storage;
  fetch?: typeof fetch;
  cacheStorage?: CacheStorage;
  assetBaseUrl?: string | URL;
  releaseInstaller?: OfficeRuntimeResourceInstaller;
};

class VolatileStorage implements Storage {
  private readonly values = new Map<string, string>();
  get length(): number {
    return this.values.size;
  }
  clear(): void {
    this.values.clear();
  }
  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }
  key(index: number): string | null {
    return [...this.values.keys()][index] ?? null;
  }
  removeItem(key: string): void {
    this.values.delete(key);
  }
  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}

function cloneProgress(progress: RuntimeCacheProgress): RuntimeCacheProgress {
  return {
    ...progress,
    failures: (progress.failures || []).map((failure) => ({ ...failure })),
    categories: progress.categories.map((category) => ({ ...category })),
  };
}

export class OfficeRuntimeResourceManager {
  private readonly controller: RuntimeCacheController;
  private readonly installer: OfficeRuntimeResourceInstaller | null;
  private readonly listeners = new Set<OfficeRuntimeResourceListener>();
  private readonly inFlight = new Map<string, Promise<RuntimeCacheProgress>>();
  private queue: Promise<unknown> = Promise.resolve();
  private snapshot: OfficeRuntimeResourceSnapshot;

  private constructor(controller: RuntimeCacheController, installer: OfficeRuntimeResourceInstaller | null) {
    this.controller = controller;
    this.installer = installer;
    const progress = controller.getProgress(controller.isComplete() ? 'complete' : 'ready');
    this.snapshot = this.buildSnapshot(progress, null, null);
    this.installer?.subscribeInstaller(() => {
      this.publish(this.snapshot.progress, this.snapshot.operation, this.snapshot.error);
    });
  }

  static async create(options: OfficeRuntimeResourceManagerOptions = {}): Promise<OfficeRuntimeResourceManager> {
    let installer = options.releaseInstaller || null;
    if (!installer) {
      try {
        installer = await createTransactionalResourceInstaller({
          assetBaseUrl: options.assetBaseUrl,
          fetch: options.fetch,
          cacheStorage: options.cacheStorage,
        });
      } catch {
        // Release Manifest v3 is deployed additively. Keep the v2 compatibility
        // path available until the stable channel pointer exists.
      }
    }
    const controller = await RuntimeCacheController.create(
      installer ? new VolatileStorage() : options.storage,
      options.fetch,
      options.cacheStorage,
      options.assetBaseUrl,
    );
    return new OfficeRuntimeResourceManager(controller, installer);
  }

  getSnapshot(): OfficeRuntimeResourceSnapshot {
    return this.snapshot;
  }

  subscribe(listener: OfficeRuntimeResourceListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getVerifiedFontPaths(): string[] {
    return [
      ...new Set([
        ...(this.installer?.getInstalledPaths().filter((path) => path.startsWith('fonts/')) || []),
        ...this.controller
          .listFonts()
          .filter((font) => font.downloaded)
          .flatMap((font) => font.paths),
      ]),
    ];
  }

  remainingBytes(): number {
    return this.controller.remainingBytes();
  }

  shouldCheckHealth(now = Date.now()): boolean {
    return this.controller.shouldCheckHealth(now);
  }

  loadAll(): Promise<RuntimeCacheProgress> {
    if (this.installer) {
      return this.run('load-all', 'load-all', async () => {
        const plan = await this.installer!.plan({ scope: 'all' });
        await this.installer!.apply(plan);
        return this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready');
      });
    }
    return this.run('load-all', 'load-all', (notify) => this.controller.loadAll(notify));
  }

  prepareForDocumentType(type: OfficeDocumentResourceType): Promise<RuntimeCacheProgress> {
    if (this.installer) {
      return this.run(`prepare-document:${type}`, 'prepare-document', async () => {
        const plan = await this.installer!.plan({ scope: 'document', documentType: type });
        await this.installer!.apply(plan);
        return this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready');
      });
    }
    return this.run(`prepare-document:${type}`, 'prepare-document', (notify) =>
      this.controller.loadCategories(['fonts', 'core', type], notify),
    );
  }

  prefetchRecommended(): Promise<RuntimeCacheProgress> {
    if (this.installer) {
      return this.run('prefetch-recommended', 'prefetch-recommended', async () => {
        const plan = await this.installer!.plan({ scope: 'recommended' });
        await this.installer!.apply(plan);
        return this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready');
      });
    }
    return this.run('prefetch-recommended', 'prefetch-recommended', (notify) =>
      this.controller.loadCategories(['fonts', 'core'], notify),
    );
  }

  checkHealth(): Promise<RuntimeCacheProgress> {
    if (this.installer) {
      return this.run('check-health', 'check-health', async () => {
        await this.installer!.checkHealth();
        return this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready');
      });
    }
    return this.run('check-health', 'check-health', (notify) => this.controller.checkHealth(notify));
  }

  downloadFontFamily(id: string): Promise<RuntimeCacheProgress> {
    return this.run(`download-font:${id}`, 'download-font', (notify) => this.controller.downloadFontFamily(id, notify));
  }

  uninstallFontFamily(id: string): Promise<RuntimeCacheProgress> {
    return this.run(`remove-font:${id}`, 'remove-font', async () => this.controller.uninstallFontFamily(id));
  }

  installFontPreset(preset: OfficeFontPreset): Promise<RuntimeCacheProgress> {
    if (preset === 'basic') return this.prefetchRecommended();
    if (this.installer) {
      return this.run('font-preset:office-compatibility', 'install-font-preset', async () => {
        const plan = await this.installer!.plan({ scope: 'fonts' });
        await this.installer!.apply(plan);
        return this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready');
      });
    }
    return this.run('font-preset:office-compatibility', 'install-font-preset', async (notify) => {
      const preferredFamilies = [
        'Aptos',
        'DengXian',
        'Arial',
        'Calibri',
        'Cambria',
        'Times New Roman',
        'Microsoft YaHei',
        'SimSun',
      ];
      let progress = await this.controller.loadCategories(['fonts', 'core'], notify);
      const available = this.controller.listFonts();
      for (const name of preferredFamilies) {
        const family = available.find((candidate) => candidate.name.toLocaleLowerCase() === name.toLocaleLowerCase());
        if (family && !family.downloaded) {
          progress = await this.controller.downloadFontFamily(family.id, notify);
        }
      }
      return progress;
    });
  }

  repair(options: { scope: 'required' | 'installed' | 'all' } = { scope: 'installed' }): Promise<RuntimeCacheProgress> {
    if (this.installer) {
      return this.run(`repair:${options.scope}`, 'check-health', async () => {
        await this.installer!.repair(options);
        return this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready');
      });
    }
    return this.run('check-health', 'check-health', (notify) => this.controller.repairInstalled(notify));
  }

  plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    if (!this.installer) {
      throw new Error('Release Manifest v3 is not available');
    }
    return this.installer.plan(request);
  }

  async apply(plan: ResourcePlan): Promise<void> {
    if (!this.installer) throw new Error('Release Manifest v3 is not available');
    await this.installer.apply(plan);
    this.publish(this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready'), null, null);
  }

  async checkForUpdates(): Promise<void> {
    if (!this.installer) return;
    await this.installer.checkForUpdates();
    this.publish(this.controller.getProgress(this.controller.isComplete() ? 'complete' : 'ready'), null, null);
  }

  pause(): void {
    this.installer?.pause();
    this.publish(this.snapshot.progress, this.snapshot.operation, null);
  }

  async resume(): Promise<void> {
    await this.installer?.resume();
    this.publish(this.snapshot.progress, this.snapshot.operation, null);
  }

  cancel(): void {
    this.installer?.cancel();
    this.publish(this.snapshot.progress, null, null);
  }

  private buildSnapshot(
    progress: RuntimeCacheProgress,
    operation: OfficeRuntimeResourceOperation | null,
    error: { code: ResourceErrorCode; path?: string } | null,
  ): OfficeRuntimeResourceSnapshot {
    const installer = this.installer?.getInstallerSnapshot();
    const installedProfiles = new Set(installer?.installedProfiles || []);
    const packs = progress.categories.map((category) => ({
      id: category.category,
      ready:
        (category.category === 'core' && installedProfiles.has('base')) ||
        (category.category === 'fonts' && installedProfiles.has('fonts-basic')) ||
        (category.category !== 'core' && category.category !== 'fonts' && installedProfiles.has(category.category)) ||
        (category.totalFiles > 0 && category.completedFiles === category.totalFiles),
      completedBytes: category.completedBytes,
      totalBytes: category.totalBytes,
    }));
    return {
      packageVersion: ONLYOFFICE_BROWSER_VERSION,
      assetVersion: this.controller.version,
      readiness:
        installer?.readiness ||
        (error || progress.phase === 'error'
          ? 'error'
          : operation || progress.phase === 'checking' || progress.phase === 'loading'
            ? 'updating'
            : packs.every((pack) => pack.ready)
              ? 'ready'
              : 'needs-download'),
      packs,
      progress: cloneProgress(progress),
      fonts: this.controller.listFonts(),
      verifiedFontPaths: this.getVerifiedFontPaths(),
      operation,
      error,
      installedRelease: installer?.installedRelease || null,
      targetRelease: installer?.targetRelease || this.controller.version,
      availableRelease: installer?.availableRelease || null,
      storageMode: installer?.storageMode || 'http-cache',
      phase: installer?.phase || (operation ? 'downloading' : 'idle'),
      currentChunk: installer?.currentChunk || null,
      currentChunkIndex: installer?.currentChunkIndex || 0,
      currentChunkCount: installer?.currentChunkCount || 0,
      downloadedBytes: installer?.downloadedBytes || progress.completedBytes,
      downloadBytes: installer?.downloadBytes || progress.totalBytes,
      verifiedBytes: installer?.verifiedBytes || progress.completedBytes,
      verifyBytes: installer?.verifyBytes || progress.totalBytes,
      bytesPerSecond: installer?.bytesPerSecond || 0,
      failedResources:
        installer?.failedResources ||
        (progress.failures || []).map((failure) => ({
          path: failure.path,
          code: 'integrity' as const,
          attempts: 1,
        })),
      canPause: installer?.canPause || false,
      canResume: installer?.canResume || false,
      canRetry: installer?.canRetry || false,
    };
  }

  private publish(
    progress: RuntimeCacheProgress,
    operation: OfficeRuntimeResourceOperation | null,
    error: { code: ResourceErrorCode; path?: string } | null,
  ): void {
    this.snapshot = this.buildSnapshot(progress, operation, error);
    for (const listener of this.listeners) listener(this.snapshot);
  }

  private run(
    key: string,
    operation: OfficeRuntimeResourceOperation,
    task: (notify: (progress: RuntimeCacheProgress) => void) => Promise<RuntimeCacheProgress>,
  ): Promise<RuntimeCacheProgress> {
    const existing = this.inFlight.get(key);
    if (existing) return existing;

    const queued = this.queue
      .catch(() => undefined)
      .then(async () => {
        this.publish(this.snapshot.progress, operation, null);
        try {
          const progress = await task((next) => this.publish(next, operation, null));
          this.publish(progress, null, null);
          return progress;
        } catch (error) {
          const nextError =
            error && typeof error === 'object' && 'code' in error
              ? {
                  code: (error as { code: ResourceErrorCode }).code,
                  path: (error as { path?: string }).path,
                }
              : { code: 'network' as const };
          this.publish(
            this.controller.getProgress('error', [
              {
                path: 'runtime',
                reason: nextError.code,
              },
            ]),
            null,
            nextError,
          );
          throw nextError;
        }
      });
    this.queue = queued;
    this.inFlight.set(key, queued);
    void queued.then(
      () => {
        if (this.inFlight.get(key) === queued) this.inFlight.delete(key);
      },
      () => {
        if (this.inFlight.get(key) === queued) this.inFlight.delete(key);
      },
    );
    return queued;
  }
}

export async function createOfficeRuntimeResourceManager(
  options: OfficeRuntimeResourceManagerOptions = {},
): Promise<OfficeRuntimeResourceManager> {
  return OfficeRuntimeResourceManager.create(options);
}
