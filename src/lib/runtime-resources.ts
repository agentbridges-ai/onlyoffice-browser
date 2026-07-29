import {
  RuntimeCacheController,
  type RuntimeCacheCategory,
  type RuntimeCacheProgress,
  type RuntimeFontFamily,
} from './runtime-cache';
import { ONLYOFFICE_BROWSER_VERSION } from '../version';

export type OfficeRuntimeResourceOperation =
  | 'prepare-document'
  | 'prefetch-recommended'
  | 'load-all'
  | 'check-health'
  | 'download-font'
  | 'install-font-preset'
  | 'remove-font';

export type OfficeRuntimeReadiness = 'ready' | 'needs-download' | 'updating' | 'error';
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
  error: Error | null;
};

export type OfficeRuntimeResourceListener = (snapshot: OfficeRuntimeResourceSnapshot) => void;

export type OfficeRuntimeResourceManagerOptions = {
  storage?: Storage;
  fetch?: typeof fetch;
  cacheStorage?: CacheStorage;
  assetBaseUrl?: string | URL;
};

function cloneProgress(progress: RuntimeCacheProgress): RuntimeCacheProgress {
  return {
    ...progress,
    categories: progress.categories.map((category) => ({ ...category })),
  };
}

export class OfficeRuntimeResourceManager {
  private readonly controller: RuntimeCacheController;
  private readonly listeners = new Set<OfficeRuntimeResourceListener>();
  private readonly inFlight = new Map<string, Promise<RuntimeCacheProgress>>();
  private queue: Promise<unknown> = Promise.resolve();
  private snapshot: OfficeRuntimeResourceSnapshot;

  private constructor(controller: RuntimeCacheController) {
    this.controller = controller;
    const progress = controller.getProgress(controller.isComplete() ? 'complete' : 'ready');
    this.snapshot = this.buildSnapshot(progress, null, null);
  }

  static async create(options: OfficeRuntimeResourceManagerOptions = {}): Promise<OfficeRuntimeResourceManager> {
    const controller = await RuntimeCacheController.create(
      options.storage,
      options.fetch,
      options.cacheStorage,
      options.assetBaseUrl,
    );
    return new OfficeRuntimeResourceManager(controller);
  }

  getSnapshot(): OfficeRuntimeResourceSnapshot {
    return {
      ...this.snapshot,
      progress: cloneProgress(this.snapshot.progress),
      fonts: this.snapshot.fonts.map((font) => ({ ...font, paths: [...font.paths] })),
      verifiedFontPaths: [...this.snapshot.verifiedFontPaths],
    };
  }

  subscribe(listener: OfficeRuntimeResourceListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getVerifiedFontPaths(): string[] {
    return [
      ...new Set(
        this.controller
          .listFonts()
          .filter((font) => font.downloaded)
          .flatMap((font) => font.paths),
      ),
    ];
  }

  remainingBytes(): number {
    return this.controller.remainingBytes();
  }

  shouldCheckHealth(now = Date.now()): boolean {
    return this.controller.shouldCheckHealth(now);
  }

  loadAll(): Promise<RuntimeCacheProgress> {
    return this.run('load-all', 'load-all', (notify) => this.controller.loadAll(notify));
  }

  prepareForDocumentType(type: OfficeDocumentResourceType): Promise<RuntimeCacheProgress> {
    return this.run(`prepare-document:${type}`, 'prepare-document', (notify) =>
      this.controller.loadCategories(['fonts', 'core', type], notify),
    );
  }

  prefetchRecommended(): Promise<RuntimeCacheProgress> {
    return this.run('prefetch-recommended', 'prefetch-recommended', (notify) =>
      this.controller.loadCategories(['fonts', 'core'], notify),
    );
  }

  checkHealth(): Promise<RuntimeCacheProgress> {
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
    return this.run('font-preset:office-compatibility', 'install-font-preset', async (notify) => {
      const preferredFamilies = ['Arial', 'Calibri', 'Cambria', 'Times New Roman', 'Microsoft YaHei', 'SimSun'];
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

  repair(): Promise<RuntimeCacheProgress> {
    return this.checkHealth();
  }

  private buildSnapshot(
    progress: RuntimeCacheProgress,
    operation: OfficeRuntimeResourceOperation | null,
    error: Error | null,
  ): OfficeRuntimeResourceSnapshot {
    const packs = progress.categories.map((category) => ({
      id: category.category,
      ready: category.totalFiles > 0 && category.completedFiles === category.totalFiles,
      completedBytes: category.completedBytes,
      totalBytes: category.totalBytes,
    }));
    return {
      packageVersion: ONLYOFFICE_BROWSER_VERSION,
      assetVersion: this.controller.version,
      readiness:
        error || progress.phase === 'error'
          ? 'error'
          : operation || progress.phase === 'checking' || progress.phase === 'loading'
            ? 'updating'
            : packs.every((pack) => pack.ready)
              ? 'ready'
              : 'needs-download',
      packs,
      progress: cloneProgress(progress),
      fonts: this.controller.listFonts(),
      verifiedFontPaths: this.getVerifiedFontPaths(),
      operation,
      error,
    };
  }

  private publish(
    progress: RuntimeCacheProgress,
    operation: OfficeRuntimeResourceOperation | null,
    error: Error | null,
  ): void {
    this.snapshot = this.buildSnapshot(progress, operation, error);
    const snapshot = this.getSnapshot();
    for (const listener of this.listeners) listener(snapshot);
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
          const nextError = error instanceof Error ? error : new Error(String(error));
          this.publish(
            this.controller.getProgress('error', Math.max(1, this.snapshot.progress.failedFiles)),
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
