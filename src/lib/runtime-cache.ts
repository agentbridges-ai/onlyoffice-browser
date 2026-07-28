export type RuntimeCacheAsset = {
  path: string;
  bytes: number;
  category: RuntimeCacheCategory;
  revision: string;
};

export type RuntimeCacheCategory = 'fonts' | 'core' | 'word' | 'cell' | 'slide';

export type RuntimeFontPackage = RuntimeCacheAsset & {
  families: string[];
};

export type RuntimeFontFamily = {
  id: string;
  name: string;
  bytes: number;
  paths: string[];
  downloaded: boolean;
  removable: boolean;
};

type RuntimeFontFamilyDefinition = {
  name: string;
  paths: string[];
};

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
  assets: Array<{ path: string; bytes: number; pack: RuntimeCacheCategory; revision: string }>;
};

type FontManifest = {
  defaultFont?: string;
  defaultFonts?: string[];
  builtInFonts?: string[];
  allFonts?: string;
  fontSelection?: string;
  fontSourceMap?: string;
  fontThumbnails?: string[];
  fonts?: string[];
  fontFamilies?: RuntimeFontFamilyDefinition[];
  assets?: Array<{ path: string; bytes: number; revision: string; families?: string[] }>;
};

type StoredProgress = {
  version: string;
  completed: string[];
  assets?: RuntimeCacheAsset[];
  lastVerifiedAt?: number;
  fontCatalog?: RuntimeFontPackage[];
  fontFamilies?: RuntimeFontFamilyDefinition[];
  requiredFonts?: string[];
};

const RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const FONT_MANIFEST_PATH = '/onlyoffice-browser-font-assets.json';
const STORAGE_KEY = 'onlyoffice-browser:shared-runtime-cache';
const INSTALLED_FONTS_STORAGE_KEY = 'onlyoffice-browser:installed-fonts';
const FONT_PACKAGE_CACHE_PREFIX = 'onlyoffice-browser-font-packages-';

function isSafeAssetPath(value: unknown): value is string {
  return typeof value === 'string' && value.length > 0 && !value.startsWith('/') && !value.includes('..');
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === 'string');
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
    (asset): asset is { path: string; bytes: number; pack: RuntimeCacheCategory; revision: string } =>
      Boolean(asset) &&
      isSafeAssetPath(asset.path) &&
      Number.isSafeInteger(asset.bytes) &&
      asset.bytes >= 0 &&
      typeof asset.revision === 'string' &&
      asset.revision.length > 0 &&
      ['core', 'word', 'cell', 'slide'].includes(asset.pack),
  );
  if (assets.length !== manifest.assets.length) {
    throw new Error('The runtime asset manifest contains an invalid asset.');
  }
  return { version: manifest.version, generatedAt: manifest.generatedAt, assets };
}

function fontAssets(
  manifest: FontManifest,
  fallbackRevision: string,
): Array<{ path: string; bytes: number; revision: string; families?: string[] }> {
  if (
    Array.isArray(manifest.assets) &&
    manifest.assets.every(
      (asset) =>
        asset &&
        isSafeAssetPath(asset.path) &&
        Number.isSafeInteger(asset.bytes) &&
        asset.bytes >= 0 &&
        typeof asset.revision === 'string' &&
        asset.revision.length > 0 &&
        (asset.families === undefined ||
          (Array.isArray(asset.families) && asset.families.every((name) => typeof name === 'string'))),
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
  return [...new Set(candidates.filter(isSafeAssetPath))].map((path) => ({
    path,
    bytes: 0,
    revision: fallbackRevision,
  }));
}

function parseStoredAssets(value: unknown): RuntimeCacheAsset[] | null {
  if (!Array.isArray(value)) return null;
  const assets = value.filter(
    (asset): asset is RuntimeCacheAsset =>
      Boolean(asset) &&
      isSafeAssetPath(asset.path) &&
      Number.isSafeInteger(asset.bytes) &&
      asset.bytes >= 0 &&
      typeof asset.revision === 'string' &&
      asset.revision.length > 0 &&
      ['fonts', 'core', 'word', 'cell', 'slide'].includes(asset.category),
  );
  return assets.length === value.length && assets.length > 0 ? assets : null;
}

function parseStoredFontCatalog(value: unknown): RuntimeFontPackage[] | null {
  if (!Array.isArray(value)) return null;
  const fonts = value.filter(
    (asset): asset is RuntimeFontPackage =>
      Boolean(asset) &&
      isSafeAssetPath(asset.path) &&
      asset.path.startsWith('fonts/') &&
      Number.isSafeInteger(asset.bytes) &&
      asset.bytes >= 0 &&
      typeof asset.revision === 'string' &&
      asset.revision.length > 0 &&
      Array.isArray(asset.families) &&
      asset.families.every((name: unknown) => typeof name === 'string'),
  );
  return fonts.length === value.length ? fonts : null;
}

function parseFontFamilies(value: unknown): RuntimeFontFamilyDefinition[] | null {
  if (!Array.isArray(value)) return null;
  const families = value.filter(
    (family): family is RuntimeFontFamilyDefinition =>
      Boolean(family) &&
      typeof family.name === 'string' &&
      family.name.length > 0 &&
      Array.isArray(family.paths) &&
      family.paths.length > 0 &&
      family.paths.every((path: unknown) => isSafeAssetPath(path) && path.startsWith('fonts/')),
  );
  return families.length === value.length ? families : null;
}

function readStoredProgress(
  storage: Storage,
  version?: string,
): {
  version: string;
  completed: Set<string>;
  assets: RuntimeCacheAsset[] | null;
  lastVerifiedAt: number;
  fontCatalog: RuntimeFontPackage[] | null;
  fontFamilies: RuntimeFontFamilyDefinition[] | null;
  requiredFonts: string[] | null;
} {
  try {
    const parsed = JSON.parse(storage.getItem(STORAGE_KEY) || 'null') as StoredProgress | null;
    if (!parsed || (version && parsed.version !== version) || !Array.isArray(parsed.completed)) {
      return {
        version: '',
        completed: new Set(),
        assets: null,
        lastVerifiedAt: 0,
        fontCatalog: null,
        fontFamilies: null,
        requiredFonts: null,
      };
    }
    return {
      version: parsed.version,
      completed: new Set(parsed.completed.filter(isSafeAssetPath)),
      assets: parseStoredAssets(parsed.assets),
      lastVerifiedAt: Number.isFinite(parsed.lastVerifiedAt) ? parsed.lastVerifiedAt || 0 : 0,
      fontCatalog: parseStoredFontCatalog(parsed.fontCatalog),
      fontFamilies: parseFontFamilies(parsed.fontFamilies),
      requiredFonts: isStringArray(parsed.requiredFonts) ? parsed.requiredFonts.filter(isSafeAssetPath) : null,
    };
  } catch {
    return {
      version: '',
      completed: new Set(),
      assets: null,
      lastVerifiedAt: 0,
      fontCatalog: null,
      fontFamilies: null,
      requiredFonts: null,
    };
  }
}

function writeStoredProgress(
  storage: Storage,
  version: string,
  completed: Set<string>,
  assets: RuntimeCacheAsset[],
  lastVerifiedAt = 0,
  fontCatalog: RuntimeFontPackage[] = [],
  fontFamilies: RuntimeFontFamilyDefinition[] = [],
  requiredFonts: string[] = [],
): void {
  storage.setItem(
    STORAGE_KEY,
    JSON.stringify({
      version,
      completed: [...completed],
      assets,
      lastVerifiedAt,
      fontCatalog,
      fontFamilies,
      requiredFonts,
    } satisfies StoredProgress),
  );
}

function readInstalledFonts(
  storage: Storage,
  requiredFonts: string[],
): { installed: Set<string>; downloaded: Set<string> } {
  try {
    const stored = JSON.parse(storage.getItem(INSTALLED_FONTS_STORAGE_KEY) || 'null');
    if (stored?.version === 2 && Array.isArray(stored.downloaded)) {
      const downloaded = new Set<string>(stored.downloaded.filter(isSafeAssetPath));
      for (const required of requiredFonts) downloaded.delete(required);
      writeDownloadedFonts(storage, downloaded);
      return {
        installed: new Set([...requiredFonts.filter(isSafeAssetPath), ...downloaded]),
        downloaded,
      };
    }
  } catch {
    // Fall through to the compact required set.
  }
  // Version 1 stored required and user-selected fonts in one array, so it
  // cannot distinguish an intentional download from a formerly over-broad
  // built-in set. Reset that ambiguous ledger once and let the current
  // manifest define required fonts precisely.
  const downloaded = new Set<string>();
  writeDownloadedFonts(storage, downloaded);
  return {
    installed: new Set(requiredFonts.filter(isSafeAssetPath)),
    downloaded,
  };
}

function writeDownloadedFonts(storage: Storage, downloaded: Set<string>): void {
  storage.setItem(
    INSTALLED_FONTS_STORAGE_KEY,
    JSON.stringify({
      version: 2,
      downloaded: [...downloaded],
    }),
  );
}

function bytesToHex(bytes: Uint8Array): string {
  return Array.from(bytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function responseBytes(
  response: Response,
  onChunk: (bytes: number) => void,
  expectedRevision?: string,
): Promise<number> {
  const digest = expectedRevision ? sha256.create() : null;
  let received = 0;
  if (!response.body) {
    const body = await response.arrayBuffer();
    const bytes = new Uint8Array(body);
    received = bytes.byteLength;
    digest?.update(bytes);
    onChunk(bytes.byteLength);
  } else {
    const reader = response.body.getReader();
    while (true) {
      const result = await reader.read();
      if (result.done) break;
      received += result.value.byteLength;
      digest?.update(result.value);
      onChunk(result.value.byteLength);
    }
  }
  if (digest && !bytesToHex(digest.digest()).startsWith(expectedRevision!)) {
    throw new Error('Asset integrity verification failed.');
  }
  return received;
}

const HEALTH_CHECK_INTERVAL_MS = 7 * 24 * 60 * 60 * 1000;

export class RuntimeCacheController {
  readonly assets: RuntimeCacheAsset[];
  readonly fontCatalog: RuntimeFontPackage[];
  readonly fontFamilies: RuntimeFontFamilyDefinition[];
  readonly fontPackageSharedPaths: string[];
  readonly requiredFontPaths: string[];
  readonly version: string;
  readonly completed: Set<string>;
  private readonly storage: Storage;
  private readonly fetchImpl: typeof fetch;
  private readonly assetBaseUrl: URL;
  private readonly installedFonts: Set<string>;
  private readonly downloadedFonts: Set<string>;
  private readonly cacheStorage?: CacheStorage;
  private fontPackageCachePromise?: Promise<Cache>;
  private lastVerifiedAt: number;

  private constructor(
    assets: RuntimeCacheAsset[],
    version: string,
    storage: Storage,
    fetchImpl: typeof fetch,
    completed?: Set<string>,
    lastVerifiedAt = 0,
    fontCatalog: RuntimeFontPackage[] = [],
    fontFamilies: RuntimeFontFamilyDefinition[] = [],
    installedFonts = new Set<string>(),
    downloadedFonts = new Set<string>(),
    requiredFontPaths: string[] = [],
    cacheStorage?: CacheStorage,
    assetBaseUrl: string | URL = window.location.origin,
  ) {
    this.assets = assets;
    this.fontCatalog = fontCatalog;
    this.fontFamilies = fontFamilies;
    this.requiredFontPaths = [...new Set(requiredFontPaths)];
    this.fontPackageSharedPaths = assets
      .filter(
        (asset) =>
          asset.category === 'fonts' &&
          (!asset.path.startsWith('fonts/') || this.requiredFontPaths.includes(asset.path)),
      )
      .map((asset) => asset.path);
    this.version = version;
    this.storage = storage;
    this.fetchImpl = fetchImpl;
    this.assetBaseUrl = new URL('/', assetBaseUrl);
    this.completed = completed || readStoredProgress(storage, version).completed;
    this.lastVerifiedAt = lastVerifiedAt;
    this.installedFonts = installedFonts;
    this.downloadedFonts = downloadedFonts;
    this.cacheStorage = cacheStorage;
  }

  static async create(
    storage: Storage = window.localStorage,
    fetchImpl: typeof fetch = window.fetch.bind(window),
    cacheStorage: CacheStorage | undefined = window.caches,
    assetBaseUrl: string | URL = window.location.origin,
  ): Promise<RuntimeCacheController> {
    const baseUrl = new URL('/', assetBaseUrl);
    const manifestUrl = new URL(RUNTIME_MANIFEST_PATH, baseUrl);
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
    if (stored?.assets && stored.fontCatalog && stored.fontFamilies && stored.requiredFonts) {
      const { installed: installedFonts, downloaded: downloadedFonts } = readInstalledFonts(
        storage,
        stored.requiredFonts,
      );
      const controller = new RuntimeCacheController(
        stored.assets,
        deployedVersion,
        storage,
        fetchImpl,
        stored.completed,
        stored.lastVerifiedAt,
        stored.fontCatalog,
        stored.fontFamilies,
        installedFonts,
        downloadedFonts,
        stored.requiredFonts,
        cacheStorage,
        baseUrl,
      );
      await controller.reconcileFontPackageCache();
      return controller;
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
      const recoveryUrl = new URL(RUNTIME_MANIFEST_PATH, baseUrl);
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

    const fontManifestUrl = new URL(FONT_MANIFEST_PATH, baseUrl);
    if (version) fontManifestUrl.searchParams.set('__oobv', version);
    const fontResponse = await fetchImpl(fontManifestUrl.href, {
      cache: version ? 'force-cache' : 'reload',
      credentials: 'omit',
    });
    if (!fontResponse.ok) throw new Error(`Font manifest request failed (${fontResponse.status}).`);
    const fonts = (await fontResponse.json()) as FontManifest;
    const requiredFontPaths = [
      ...(fonts.defaultFonts || (fonts.defaultFont ? [fonts.defaultFont] : [])),
      ...(fonts.builtInFonts || []),
    ];
    const { installed: installedFonts, downloaded: downloadedFonts } = readInstalledFonts(storage, requiredFontPaths);
    const fontInventory = fontAssets(fonts, version);
    const fontCatalog = fontInventory
      .filter((asset) => asset.path.startsWith('fonts/'))
      .map(
        (asset) =>
          ({
            path: asset.path,
            bytes: asset.bytes,
            revision: asset.revision,
            category: 'fonts',
            families: asset.families || [],
          }) satisfies RuntimeFontPackage,
      );
    const fontFamilies = parseFontFamilies(fonts.fontFamilies) || [];
    const byPath = new Map(
      runtime.assets.map((asset) => [
        asset.path,
        {
          path: asset.path,
          bytes: asset.bytes,
          category: asset.pack,
          revision: asset.revision,
        } satisfies RuntimeCacheAsset,
      ]),
    );
    byPath.set(FONT_MANIFEST_PATH.slice(1), {
      path: FONT_MANIFEST_PATH.slice(1),
      bytes: Number(fontResponse.headers.get('content-length')) || 0,
      category: 'fonts',
      revision: version,
    });
    for (const asset of fontInventory.filter(
      (asset) => !asset.path.startsWith('fonts/') || installedFonts.has(asset.path),
    )) {
      byPath.set(asset.path, { ...asset, category: 'fonts' });
    }
    const assets = [...byPath.values()];
    const previous = readStoredProgress(storage);
    const previousByPath = new Map((previous.assets || []).map((asset) => [asset.path, asset]));
    const completed = new Set(
      assets
        .filter((asset) => {
          const oldAsset = previousByPath.get(asset.path);
          return previous.completed.has(asset.path) && oldAsset?.revision === asset.revision;
        })
        .map((asset) => asset.path),
    );
    const controller = new RuntimeCacheController(
      assets,
      version || runtime.generatedAt,
      storage,
      fetchImpl,
      completed,
      previous.version === version ? previous.lastVerifiedAt : 0,
      fontCatalog,
      fontFamilies,
      installedFonts,
      downloadedFonts,
      requiredFontPaths,
      cacheStorage,
      baseUrl,
    );
    await controller.reconcileFontPackageCache();
    writeStoredProgress(
      storage,
      controller.version,
      controller.completed,
      controller.assets,
      controller.lastVerifiedAt,
      controller.fontCatalog,
      controller.fontFamilies,
      controller.requiredFontPaths,
    );
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

  shouldCheckHealth(now = Date.now()): boolean {
    return this.isComplete() && now - this.lastVerifiedAt >= HEALTH_CHECK_INTERVAL_MS;
  }

  remainingBytes(): number {
    return this.assets
      .filter((asset) => !this.completed.has(asset.path))
      .reduce((total, asset) => total + asset.bytes, 0);
  }

  listFonts(): RuntimeFontFamily[] {
    const fontsByPath = new Map(this.fontCatalog.map((font) => [font.path, font]));
    return this.fontFamilies
      .map(({ name, paths }) => {
        const fonts = paths.map((path) => fontsByPath.get(path)).filter((font): font is RuntimeFontPackage => !!font);
        return {
          id: name.toLocaleLowerCase(),
          name,
          bytes: fonts.reduce((total, font) => total + font.bytes, 0),
          paths: fonts.map((font) => font.path),
          downloaded:
            fonts.length > 0 &&
            [...fonts.map((font) => font.path), ...this.fontPackageSharedPaths].every((path) =>
              this.completed.has(path),
            ),
          removable: fonts.some((font) => !this.fontPackageSharedPaths.includes(font.path)),
        };
      })
      .filter((family) => family.paths.length > 0);
  }

  async downloadFontFamily(
    id: string,
    onProgress: (progress: RuntimeCacheProgress) => void,
  ): Promise<RuntimeCacheProgress> {
    const family = this.listFonts().find((candidate) => candidate.id === id);
    if (!family) throw new Error(`Unknown font family: ${id}`);
    let failedFiles = 0;
    const packagePaths = [...new Set([...this.fontPackageSharedPaths, ...family.paths])];
    for (const path of packagePaths) {
      const asset =
        this.assets.find((candidate) => candidate.path === path) ||
        this.fontCatalog.find((candidate) => candidate.path === path);
      if (!asset) {
        failedFiles += 1;
        continue;
      }
      if (!this.assets.some((candidate) => candidate.path === path)) this.assets.push(asset);
      if (this.completed.has(path)) continue;

      onProgress(this.getProgress('loading'));
      let verified = false;
      for (const cacheMode of ['force-cache', 'reload'] as const) {
        try {
          await this.downloadAndVerify(asset, cacheMode, () => onProgress(this.getProgress('loading')));
          verified = true;
          break;
        } catch {
          // Retry once while bypassing a missing or corrupt HTTP-cache entry.
        }
      }
      if (!verified) {
        const index = this.assets.findIndex((asset) => asset.path === path);
        if (index >= 0) this.assets.splice(index, 1);
        failedFiles += 1;
        continue;
      }
      this.completed.add(path);
      if (path.startsWith('fonts/')) {
        this.installedFonts.add(path);
        if (!this.fontPackageSharedPaths.includes(path)) this.downloadedFonts.add(path);
      }
    }
    writeDownloadedFonts(this.storage, this.downloadedFonts);
    writeStoredProgress(
      this.storage,
      this.version,
      this.completed,
      this.assets,
      this.lastVerifiedAt,
      this.fontCatalog,
      this.fontFamilies,
      this.requiredFontPaths,
    );
    const progress = this.getProgress(
      failedFiles === 0 ? (this.isComplete() ? 'complete' : 'ready') : 'error',
      failedFiles,
    );
    onProgress(progress);
    return progress;
  }

  async uninstallFontFamily(id: string): Promise<RuntimeCacheProgress> {
    const family = this.listFonts().find((candidate) => candidate.id === id);
    if (!family) throw new Error(`Unknown font family: ${id}`);
    if (!family.downloaded || !family.removable) return this.getProgress(this.isComplete() ? 'complete' : 'ready');

    const removablePaths = family.paths.filter((path) => !this.fontPackageSharedPaths.includes(path));
    const cache = await this.openFontPackageCache();
    for (const path of removablePaths) {
      const asset = this.fontCatalog.find((candidate) => candidate.path === path);
      if (!asset) continue;
      if (cache) {
        await Promise.all([
          cache.delete(this.versionedUrl(asset)),
          cache.delete(new URL(asset.path, this.assetBaseUrl).href),
        ]);
      }
      this.completed.delete(path);
      this.installedFonts.delete(path);
      this.downloadedFonts.delete(path);
      const assetIndex = this.assets.findIndex((candidate) => candidate.path === path);
      if (assetIndex >= 0) this.assets.splice(assetIndex, 1);
    }
    writeDownloadedFonts(this.storage, this.downloadedFonts);
    writeStoredProgress(
      this.storage,
      this.version,
      this.completed,
      this.assets,
      this.lastVerifiedAt,
      this.fontCatalog,
      this.fontFamilies,
      this.requiredFontPaths,
    );
    return this.getProgress(this.isComplete() ? 'complete' : 'ready');
  }

  private async openFontPackageCache(): Promise<Cache | undefined> {
    if (!this.cacheStorage) return undefined;
    this.fontPackageCachePromise ||= this.cacheStorage.open(`${FONT_PACKAGE_CACHE_PREFIX}${this.version}`);
    return this.fontPackageCachePromise;
  }

  private async reconcileFontPackageCache(): Promise<void> {
    if (!this.cacheStorage) return;
    const cache = await this.openFontPackageCache();
    if (!cache) return;
    for (const asset of this.assets.filter((candidate) => candidate.category === 'fonts')) {
      const requestUrl = this.versionedUrl(asset);
      const cached = (await cache.match(requestUrl)) || (await this.cacheStorage.match(requestUrl));
      if (cached) {
        this.completed.add(asset.path);
        if (!(await cache.match(requestUrl))) await cache.put(requestUrl, cached.clone());
      } else {
        this.completed.delete(asset.path);
      }
    }
    const cacheNames = await this.cacheStorage.keys();
    await Promise.all(
      cacheNames
        .filter(
          (name) =>
            name.startsWith(FONT_PACKAGE_CACHE_PREFIX) && name !== `${FONT_PACKAGE_CACHE_PREFIX}${this.version}`,
        )
        .map((name) => this.cacheStorage!.delete(name)),
    );
    writeStoredProgress(
      this.storage,
      this.version,
      this.completed,
      this.assets,
      this.lastVerifiedAt,
      this.fontCatalog,
      this.fontFamilies,
      this.requiredFontPaths,
    );
  }

  private expectedIntegrity(asset: RuntimeCacheAsset): string | undefined {
    return asset.path === FONT_MANIFEST_PATH.slice(1) ? undefined : asset.revision;
  }

  private versionedUrl(asset: RuntimeCacheAsset): string {
    const url = new URL(asset.path, this.assetBaseUrl);
    url.searchParams.set('__oobv', asset.revision || this.version);
    return url.href;
  }

  private async resolveUnknownSizes(onProgress: (progress: RuntimeCacheProgress) => void): Promise<void> {
    const unknown = this.assets.filter((asset) => asset.bytes === 0);
    let cursor = 0;
    const worker = async () => {
      while (cursor < unknown.length) {
        const asset = unknown[cursor++];
        try {
          const response = await this.fetchImpl(this.versionedUrl(asset), {
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

  private async downloadAndVerify(
    asset: RuntimeCacheAsset,
    cache: RequestCache,
    onChunk: (bytes: number) => void,
  ): Promise<number> {
    const response = await this.fetchImpl(this.versionedUrl(asset), {
      cache,
      credentials: 'omit',
      mode: this.assetBaseUrl.origin === window.location.origin ? 'same-origin' : 'cors',
    });
    if (!response.ok) throw new Error(`${asset.path}: HTTP ${response.status}`);
    const cacheResponse = asset.category === 'fonts' ? response.clone() : null;
    const received = await responseBytes(response, onChunk, this.expectedIntegrity(asset));
    if (cacheResponse) {
      const fontCache = await this.openFontPackageCache();
      await fontCache?.put(this.versionedUrl(asset), cacheResponse);
    }
    return received;
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
          let received = 0;
          let verified = false;
          for (const cacheMode of ['force-cache', 'reload'] as const) {
            let attemptReceived = 0;
            try {
              received = await this.downloadAndVerify(asset, cacheMode, (bytes) => {
                attemptReceived += bytes;
                downloadedBytes += bytes;
                onProgress({
                  phase: 'loading',
                  completedFiles,
                  totalFiles,
                  completedBytes: Math.min(downloadedBytes, totalBytes),
                  totalBytes,
                  failedFiles,
                  categories: this.categoryProgress(
                    this.assets.filter((candidate) => this.completed.has(candidate.path)),
                  ),
                });
              });
              verified = true;
              break;
            } catch {
              downloadedBytes = Math.max(0, downloadedBytes - attemptReceived);
              received = 0;
            }
          }
          if (!verified) throw new Error(`${asset.path}: integrity verification failed`);
          if (asset.bytes === 0) asset.bytes = received;
          this.completed.add(asset.path);
          completedFiles += 1;
          writesSinceFlush += 1;
          if (writesSinceFlush >= 10) {
            writeStoredProgress(
              this.storage,
              this.version,
              this.completed,
              this.assets,
              this.lastVerifiedAt,
              this.fontCatalog,
              this.fontFamilies,
              this.requiredFontPaths,
            );
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
    if (failedFiles === 0 && this.isComplete()) this.lastVerifiedAt = Date.now();
    writeStoredProgress(
      this.storage,
      this.version,
      this.completed,
      this.assets,
      this.lastVerifiedAt,
      this.fontCatalog,
      this.fontFamilies,
      this.requiredFontPaths,
    );
    const progress = this.getProgress(failedFiles === 0 && this.isComplete() ? 'complete' : 'error', failedFiles);
    onProgress(progress);
    return progress;
  }

  async checkHealth(onProgress: (progress: RuntimeCacheProgress) => void): Promise<RuntimeCacheProgress> {
    if (!this.isComplete()) return this.loadAll(onProgress);

    let failedFiles = 0;
    onProgress(this.getProgress('checking'));
    let cursor = 0;
    const worker = async () => {
      while (cursor < this.assets.length) {
        const asset = this.assets[cursor++];
        try {
          await this.downloadAndVerify(asset, 'only-if-cached', () => undefined);
        } catch {
          this.completed.delete(asset.path);
          writeStoredProgress(
            this.storage,
            this.version,
            this.completed,
            this.assets,
            this.lastVerifiedAt,
            this.fontCatalog,
            this.fontFamilies,
            this.requiredFontPaths,
          );
          try {
            await this.downloadAndVerify(asset, 'reload', () => undefined);
            this.completed.add(asset.path);
          } catch {
            failedFiles += 1;
          }
        }
        onProgress(this.getProgress(failedFiles === 0 ? 'checking' : 'error', failedFiles));
      }
    };
    await Promise.all(Array.from({ length: Math.min(3, this.assets.length) }, worker));
    if (failedFiles === 0 && this.isComplete()) this.lastVerifiedAt = Date.now();
    writeStoredProgress(
      this.storage,
      this.version,
      this.completed,
      this.assets,
      this.lastVerifiedAt,
      this.fontCatalog,
      this.fontFamilies,
      this.requiredFontPaths,
    );
    const progress = this.getProgress(failedFiles === 0 ? 'complete' : 'error', failedFiles);
    onProgress(progress);
    return progress;
  }
}
import { sha256 } from '@noble/hashes/sha2.js';
