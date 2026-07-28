export type RuntimeCacheAsset = {
  path: string;
  bytes: number;
  category: RuntimeCacheCategory;
};

export type RuntimeCacheCategory = 'fonts' | 'core' | 'word' | 'cell' | 'slide';

export type RuntimeCacheCategoryProgress = {
  category: RuntimeCacheCategory;
  completedFiles: number;
  totalFiles: number;
  completedBytes: number;
  totalBytes: number;
};

export type RuntimeCacheProgress = {
  phase: 'checking' | 'ready' | 'loading' | 'complete' | 'error';
  completedFiles: number;
  totalFiles: number;
  completedBytes: number;
  totalBytes: number;
  failedFiles: number;
  categories: RuntimeCacheCategoryProgress[];
};

type RuntimeManifest = {
  version: number;
  generatedAt: string;
  assets: Array<{ path: string; bytes: number; pack: RuntimeCacheCategory }>;
};

type FontManifest = {
  allFonts?: string;
  fontSelection?: string;
  fontSourceMap?: string;
  fontThumbnails?: string[];
  fonts?: string[];
  assets?: Array<{ path: string; bytes: number }>;
};

type StoredProgress = {
  version: string;
  completed: string[];
  assets?: RuntimeCacheAsset[];
};

const RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const FONT_MANIFEST_PATH = '/onlyoffice-browser-font-assets.json';
const STORAGE_KEY = 'onlyoffice-browser:shared-runtime-cache';

function isSafeAssetPath(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && !value.startsWith('/') && !value.includes('..');
}

function parseRuntimeManifest(value: unknown): RuntimeManifest {
  const manifest = value as Partial<RuntimeManifest>;
  if (
    !manifest ||
    typeof manifest.version !== 'number' ||
    typeof manifest.generatedAt !== 'string' ||
    !Array.isArray(manifest.assets)
  ) {
    throw new Error('The runtime asset manifest is missing its asset inventory.');
  }
  const assets = manifest.assets.filter(
    (asset): asset is { path: string; bytes: number; pack: RuntimeCacheCategory } =>
      Boolean(asset) &&
      isSafeAssetPath(asset.path) &&
      Number.isSafeInteger(asset.bytes) &&
      asset.bytes >= 0 &&
      ['core', 'word', 'cell', 'slide'].includes(asset.pack),
  );
  if (assets.length !== manifest.assets.length) {
    throw new Error('The runtime asset manifest contains an invalid asset.');
  }
  return { version: manifest.version, generatedAt: manifest.generatedAt, assets };
}

function fontAssets(manifest: FontManifest): Array<{ path: string; bytes: number }> {
  if (
    Array.isArray(manifest.assets) &&
    manifest.assets.every(
      (asset) => asset && isSafeAssetPath(asset.path) && Number.isSafeInteger(asset.bytes) && asset.bytes >= 0,
    )
  ) {
    return manifest.assets;
  }
  const candidates = [
    manifest.fontSourceMap,
    manifest.allFonts,
    manifest.fontSelection,
    ...(manifest.fontThumbnails || []),
    ...(manifest.fonts || []),
  ];
  return [...new Set(candidates.filter(isSafeAssetPath))].map((path) => ({ path, bytes: 0 }));
}

function parseStoredAssets(value: unknown): RuntimeCacheAsset[] | null {
  if (!Array.isArray(value)) return null;
  const assets = value.filter(
    (asset): asset is RuntimeCacheAsset =>
      Boolean(asset) &&
      isSafeAssetPath(asset.path) &&
      Number.isSafeInteger(asset.bytes) &&
      asset.bytes >= 0 &&
      ['fonts', 'core', 'word', 'cell', 'slide'].includes(asset.category),
  );
  return assets.length === value.length && assets.length > 0 ? assets : null;
}

function readStoredProgress(
  storage: Storage,
  version: string,
): { completed: Set<string>; assets: RuntimeCacheAsset[] | null } {
  try {
    const parsed = JSON.parse(storage.getItem(STORAGE_KEY) || 'null') as StoredProgress | null;
    if (parsed?.version !== version || !Array.isArray(parsed.completed)) {
      return { completed: new Set(), assets: null };
    }
    return {
      completed: new Set(parsed.completed.filter(isSafeAssetPath)),
      assets: parseStoredAssets(parsed.assets),
    };
  } catch {
    return { completed: new Set(), assets: null };
  }
}

function writeStoredProgress(
  storage: Storage,
  version: string,
  completed: Set<string>,
  assets: RuntimeCacheAsset[],
): void {
  storage.setItem(STORAGE_KEY, JSON.stringify({ version, completed: [...completed], assets } satisfies StoredProgress));
}

async function responseBytes(response: Response, onChunk: (bytes: number) => void): Promise<void> {
  if (!response.body) {
    const body = await response.arrayBuffer();
    onChunk(body.byteLength);
    return;
  }
  const reader = response.body.getReader();
  while (true) {
    const result = await reader.read();
    if (result.done) return;
    onChunk(result.value.byteLength);
  }
}

export class RuntimeCacheController {
  readonly assets: RuntimeCacheAsset[];
  readonly version: string;
  readonly completed: Set<string>;
  private readonly storage: Storage;
  private readonly fetchImpl: typeof fetch;

  private constructor(assets: RuntimeCacheAsset[], version: string, storage: Storage, fetchImpl: typeof fetch) {
    this.assets = assets;
    this.version = version;
    this.storage = storage;
    this.fetchImpl = fetchImpl;
    this.completed = readStoredProgress(storage, version).completed;
  }

  static async create(
    storage: Storage = window.localStorage,
    fetchImpl: typeof fetch = window.fetch.bind(window),
  ): Promise<RuntimeCacheController> {
    const manifestUrl = new URL(RUNTIME_MANIFEST_PATH, window.location.origin);
    const versionResponse = await fetchImpl(manifestUrl.href, {
      method: 'HEAD',
      cache: 'no-cache',
      credentials: 'omit',
    });
    if (!versionResponse.ok) {
      throw new Error(`Runtime manifest version request failed (${versionResponse.status}).`);
    }
    const deployedVersion =
      versionResponse.headers.get('X-OnlyOffice-Asset-Version') ||
      versionResponse.headers.get('etag')?.replaceAll('"', '') ||
      '';
    const stored = deployedVersion ? readStoredProgress(storage, deployedVersion) : null;
    if (stored?.assets) {
      return new RuntimeCacheController(stored.assets, deployedVersion, storage, fetchImpl);
    }

    if (deployedVersion) manifestUrl.searchParams.set('__oobv', deployedVersion);
    let runtimeResponse = await fetchImpl(manifestUrl.href, {
      cache: deployedVersion ? 'force-cache' : 'reload',
      credentials: 'omit',
    });
    if (!runtimeResponse.ok) throw new Error(`Runtime manifest request failed (${runtimeResponse.status}).`);
    let runtimeValue = await runtimeResponse.json();
    let runtime: RuntimeManifest;
    try {
      runtime = parseRuntimeManifest(runtimeValue);
    } catch {
      const recoveryUrl = new URL(RUNTIME_MANIFEST_PATH, window.location.origin);
      recoveryUrl.searchParams.set('__cache_status', String(Date.now()));
      recoveryUrl.searchParams.set('retry', '1');
      runtimeResponse = await fetchImpl(recoveryUrl.href, { cache: 'reload', credentials: 'omit' });
      if (!runtimeResponse.ok) {
        throw new Error(`Runtime manifest recovery request failed (${runtimeResponse.status}).`);
      }
      runtimeValue = await runtimeResponse.json();
      runtime = parseRuntimeManifest(runtimeValue);
    }
    const version =
      deployedVersion ||
      runtimeResponse.headers.get('X-OnlyOffice-Asset-Version') ||
      runtimeResponse.headers.get('etag')?.replaceAll('"', '') ||
      runtime.generatedAt;

    const fontManifestUrl = new URL(FONT_MANIFEST_PATH, window.location.origin);
    if (version) fontManifestUrl.searchParams.set('__oobv', version);
    const fontResponse = await fetchImpl(fontManifestUrl.href, {
      cache: version ? 'force-cache' : 'reload',
      credentials: 'omit',
    });
    if (!fontResponse.ok) throw new Error(`Font manifest request failed (${fontResponse.status}).`);
    const fonts = (await fontResponse.json()) as FontManifest;
    const byPath = new Map(
      runtime.assets.map((asset) => [
        asset.path,
        { path: asset.path, bytes: asset.bytes, category: asset.pack } satisfies RuntimeCacheAsset,
      ]),
    );
    byPath.set(FONT_MANIFEST_PATH.slice(1), {
      path: FONT_MANIFEST_PATH.slice(1),
      bytes: Number(fontResponse.headers.get('content-length')) || 0,
      category: 'fonts',
    });
    for (const asset of fontAssets(fonts)) {
      byPath.set(asset.path, { ...asset, category: 'fonts' });
    }
    const assets = [...byPath.values()];
    const controller = new RuntimeCacheController(assets, version || runtime.generatedAt, storage, fetchImpl);
    writeStoredProgress(storage, controller.version, controller.completed, controller.assets);
    return controller;
  }

  getProgress(phase: RuntimeCacheProgress['phase'] = 'ready', failedFiles = 0): RuntimeCacheProgress {
    const completedAssets = this.assets.filter((asset) => this.completed.has(asset.path));
    return {
      phase,
      completedFiles: completedAssets.length,
      totalFiles: this.assets.length,
      completedBytes: completedAssets.reduce((total, asset) => total + asset.bytes, 0),
      totalBytes: this.assets.reduce((total, asset) => total + asset.bytes, 0),
      failedFiles,
      categories: this.categoryProgress(completedAssets),
    };
  }

  private categoryProgress(completedAssets: RuntimeCacheAsset[]): RuntimeCacheCategoryProgress[] {
    return (['fonts', 'core', 'word', 'cell', 'slide'] as const).map((category) => {
      const assets = this.assets.filter((asset) => asset.category === category);
      const completed = completedAssets.filter((asset) => asset.category === category);
      return {
        category,
        completedFiles: completed.length,
        totalFiles: assets.length,
        completedBytes: completed.reduce((total, asset) => total + asset.bytes, 0),
        totalBytes: assets.reduce((total, asset) => total + asset.bytes, 0),
      };
    });
  }

  isComplete(): boolean {
    return this.assets.length > 0 && this.assets.every((asset) => this.completed.has(asset.path));
  }

  private versionedUrl(path: string): string {
    const url = new URL(path, window.location.origin);
    url.searchParams.set('__oobv', this.version);
    return url.href;
  }

  private async resolveUnknownSizes(onProgress: (progress: RuntimeCacheProgress) => void): Promise<void> {
    const unknown = this.assets.filter((asset) => asset.bytes === 0);
    let cursor = 0;
    const worker = async () => {
      while (cursor < unknown.length) {
        const asset = unknown[cursor++];
        try {
          const response = await this.fetchImpl(this.versionedUrl(asset.path), {
            method: 'HEAD',
            cache: 'no-cache',
          });
          const bytes = Number(response.headers.get('content-length'));
          if (response.ok && Number.isSafeInteger(bytes) && bytes >= 0) asset.bytes = bytes;
        } finally {
          onProgress(this.getProgress('checking'));
        }
      }
    };
    await Promise.all(Array.from({ length: Math.min(6, unknown.length) }, worker));
  }

  async loadAll(onProgress: (progress: RuntimeCacheProgress) => void): Promise<RuntimeCacheProgress> {
    await this.resolveUnknownSizes(onProgress);
    const pending = this.assets.filter((asset) => !this.completed.has(asset.path));
    let cursor = 0;
    let failedFiles = 0;
    const initialProgress = this.getProgress();
    let completedFiles = initialProgress.completedFiles;
    let downloadedBytes = initialProgress.completedBytes;
    const totalFiles = initialProgress.totalFiles;
    const totalBytes = initialProgress.totalBytes;
    let writesSinceFlush = 0;
    onProgress(this.getProgress('loading'));

    const worker = async () => {
      while (cursor < pending.length) {
        const asset = pending[cursor++];
        try {
          const response = await this.fetchImpl(this.versionedUrl(asset.path), { cache: 'force-cache' });
          if (!response.ok) throw new Error(`${asset.path}: HTTP ${response.status}`);
          let received = 0;
          await responseBytes(response, (bytes) => {
            received += bytes;
            downloadedBytes += bytes;
            onProgress({
              phase: 'loading',
              completedFiles,
              totalFiles,
              completedBytes: Math.min(downloadedBytes, totalBytes),
              totalBytes,
              failedFiles,
              categories: this.categoryProgress(this.assets.filter((candidate) => this.completed.has(candidate.path))),
            });
          });
          if (asset.bytes === 0) asset.bytes = received;
          this.completed.add(asset.path);
          completedFiles += 1;
          writesSinceFlush += 1;
          if (writesSinceFlush >= 10) {
            writeStoredProgress(this.storage, this.version, this.completed, this.assets);
            writesSinceFlush = 0;
          }
        } catch {
          failedFiles += 1;
        }
        onProgress({
          phase: 'loading',
          completedFiles,
          totalFiles,
          completedBytes: Math.min(downloadedBytes, totalBytes),
          totalBytes,
          failedFiles,
          categories: this.categoryProgress(this.assets.filter((candidate) => this.completed.has(candidate.path))),
        });
      }
    };
    await Promise.all(Array.from({ length: Math.min(4, pending.length) }, worker));
    writeStoredProgress(this.storage, this.version, this.completed, this.assets);
    const progress = this.getProgress(failedFiles === 0 && this.isComplete() ? 'complete' : 'error', failedFiles);
    onProgress(progress);
    return progress;
  }
}
