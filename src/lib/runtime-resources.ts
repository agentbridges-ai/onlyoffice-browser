import { RuntimeCacheController, type RuntimeCacheProgress, type RuntimeFontFamily } from './runtime-cache';

export type OfficeRuntimeResourceOperation = 'load-all' | 'check-health' | 'download-font' | 'remove-font';

export type OfficeRuntimeResourceSnapshot = {
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

  checkHealth(): Promise<RuntimeCacheProgress> {
    return this.run('check-health', 'check-health', (notify) => this.controller.checkHealth(notify));
  }

  downloadFontFamily(id: string): Promise<RuntimeCacheProgress> {
    return this.run(`download-font:${id}`, 'download-font', (notify) => this.controller.downloadFontFamily(id, notify));
  }

  uninstallFontFamily(id: string): Promise<RuntimeCacheProgress> {
    return this.run(`remove-font:${id}`, 'remove-font', async () => this.controller.uninstallFontFamily(id));
  }

  private buildSnapshot(
    progress: RuntimeCacheProgress,
    operation: OfficeRuntimeResourceOperation | null,
    error: Error | null,
  ): OfficeRuntimeResourceSnapshot {
    return {
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
