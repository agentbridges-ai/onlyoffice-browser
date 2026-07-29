import { beforeEach, describe, expect, it, vi } from 'vitest';
import { RuntimeCacheController, type RuntimeCacheProgress } from '../../src/lib/runtime-cache';

const runtimeAssets = [
  { path: 'sdkjs/common/core.js', bytes: 4, pack: 'core', revision: '9f64a747e1b97f13' },
  { path: 'sdkjs/word/word.js', bytes: 4, pack: 'word', revision: '9f64a747e1b97f13' },
  { path: 'sdkjs/cell/cell.js', bytes: 4, pack: 'cell', revision: '9f64a747e1b97f13' },
  { path: 'sdkjs/slide/slide.js', bytes: 4, pack: 'slide', revision: '9f64a747e1b97f13' },
];

function response(body: BodyInit | null, init: ResponseInit = {}): Response {
  return new Response(body, { status: 200, ...init });
}

function memoryCacheStorage() {
  const buckets = new Map<string, Map<string, Response>>();
  const deleted: string[] = [];
  const requestUrl = (request: RequestInfo | URL) =>
    typeof request === 'string' ? request : request instanceof URL ? request.href : request.url;
  const cacheFor = (bucket: Map<string, Response>) =>
    ({
      match: async (request: RequestInfo | URL) => bucket.get(requestUrl(request))?.clone(),
      put: async (request: RequestInfo | URL, value: Response) => {
        bucket.set(requestUrl(request), value.clone());
      },
      delete: async (request: RequestInfo | URL) => {
        const url = requestUrl(request);
        deleted.push(url);
        return bucket.delete(url);
      },
    }) as unknown as Cache;
  const storage = {
    open: async (name: string) => {
      const bucket = buckets.get(name) || new Map<string, Response>();
      buckets.set(name, bucket);
      return cacheFor(bucket);
    },
    match: async (request: RequestInfo | URL) => {
      for (const bucket of buckets.values()) {
        const value = bucket.get(requestUrl(request));
        if (value) return value.clone();
      }
      return undefined;
    },
    keys: async () => [...buckets.keys()],
    delete: async (name: string) => buckets.delete(name),
  } as unknown as CacheStorage;
  return { storage, buckets, deleted };
}

describe('RuntimeCacheController', () => {
  beforeEach(() => {
    window.localStorage.clear();
  });

  it('classifies runtime and generated font assets and persists a completed version', async () => {
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-28T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'asset-v1', 'Content-Type': 'application/json' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(
          JSON.stringify({
            defaultFonts: ['fonts/000.ttf', 'fonts/001.ttf'],
            builtInFonts: ['fonts/002.ttf'],
            allFonts: 'sdkjs/common/AllFonts.js',
            fontSelection: 'server/FileConverter/bin/font_selection.bin',
            fontSourceMap: 'onlyoffice-browser-font-source-map.json',
            fontThumbnails: ['sdkjs/common/Images/fonts_thumbnail.png'],
            fonts: ['fonts/000.ttf', 'fonts/001.ttf', 'fonts/002.ttf'],
            fontFamilies: [{ name: 'Microsoft YaHei', paths: ['fonts/000.ttf', 'fonts/001.ttf'] }],
            assets: [
              { path: 'sdkjs/common/AllFonts.js', bytes: 4, revision: '9f64a747e1b97f13' },
              { path: 'server/FileConverter/bin/font_selection.bin', bytes: 4, revision: '9f64a747e1b97f13' },
              { path: 'onlyoffice-browser-font-source-map.json', bytes: 4, revision: '9f64a747e1b97f13' },
              { path: 'sdkjs/common/Images/fonts_thumbnail.png', bytes: 4, revision: '9f64a747e1b97f13' },
              {
                path: 'fonts/000.ttf',
                bytes: 4,
                revision: '9f64a747e1b97f13',
                families: ['Microsoft YaHei'],
              },
              {
                path: 'fonts/001.ttf',
                bytes: 4,
                revision: '9f64a747e1b97f13',
                families: ['Microsoft YaHei Bold'],
              },
              {
                path: 'fonts/002.ttf',
                bytes: 4,
                revision: '9f64a747e1b97f13',
                families: ['Wingdings'],
              },
            ],
          }),
          { headers: { 'Content-Type': 'application/json' } },
        );
      }
      if (init?.method === 'HEAD') {
        return response(null, { headers: { 'Content-Length': '4' } });
      }
      return response(new Uint8Array([1, 2, 3, 4]));
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    const initial = controller.getProgress();

    expect(initial.totalFiles).toBe(12);
    expect(initial.categories.map((category) => [category.category, category.totalFiles])).toEqual([
      ['fonts', 8],
      ['core', 1],
      ['word', 1],
      ['cell', 1],
      ['slide', 1],
    ]);
    expect(initial.totalBytes).toBe(44);
    expect(controller.listFonts()).toEqual([
      expect.objectContaining({
        name: 'Microsoft YaHei',
        bytes: 8,
        paths: ['fonts/000.ttf', 'fonts/001.ttf'],
      }),
    ]);

    let finalProgress: RuntimeCacheProgress | undefined;
    await controller.loadAll((progress) => {
      finalProgress = progress;
    });
    expect(finalProgress?.phase).toBe('complete');
    expect(finalProgress?.completedFiles).toBe(12);
    expect(fetchMock.mock.calls.some(([, init]) => init?.cache === 'force-cache')).toBe(true);

    const restored = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    expect(restored.isComplete()).toBe(true);
    const restoredRequests = fetchMock.mock.calls.slice(-1);
    expect(restoredRequests).toHaveLength(1);
    expect(restoredRequests[0]?.[1]).toMatchObject({ method: 'HEAD', cache: 'no-cache' });
  });

  it('migrates unchanged files and downloads only changed revisions after a release', async () => {
    const oldAssets = runtimeAssets.map((asset) => ({
      path: asset.path,
      bytes: asset.bytes,
      category: asset.pack,
      revision: asset.revision,
    }));
    window.localStorage.setItem(
      'onlyoffice-browser:shared-runtime-cache',
      JSON.stringify({
        version: 'old',
        completed: oldAssets.map((asset) => asset.path),
        assets: oldAssets,
      }),
    );
    const changedAssets = runtimeAssets.map((asset) =>
      asset.pack === 'word' ? { ...asset, revision: 'ee10da4aefe61a37' } : asset,
    );
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-28T00:00:00.000Z',
            assets: changedAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'new' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(JSON.stringify({ assets: [] }));
      }
      if (init?.method === 'HEAD') return response(null, { headers: { 'Content-Length': '4' } });
      return response(
        url.searchParams.get('__oobv') === 'ee10da4aefe61a37'
          ? new Uint8Array([4, 3, 2, 1])
          : new Uint8Array([1, 2, 3, 4]),
      );
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    expect(controller.isComplete()).toBe(false);
    expect(controller.getProgress().completedFiles).toBe(3);
    await controller.loadAll(() => undefined);
    const downloadedAssetUrls = fetchMock.mock.calls
      .filter(([, init]) => init?.cache === 'force-cache')
      .map(([input]) => String(input))
      .filter((url) => !url.includes('.json'));
    expect(downloadedAssetUrls).toEqual([expect.stringContaining('sdkjs/word/word.js?__oobv=ee10da4aefe61a37')]);
  });

  it('drops the ambiguous v1 font ledger when a narrower built-in set is released', async () => {
    window.localStorage.setItem(
      'onlyoffice-browser:installed-fonts',
      JSON.stringify(['fonts/default.ttf', 'fonts/formerly-built-in.ttf']),
    );
    const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-28T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'font-ledger-v2' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(
          JSON.stringify({
            defaultFonts: ['fonts/default.ttf'],
            builtInFonts: ['fonts/symbol.ttf'],
            fontFamilies: [],
            assets: [
              { path: 'fonts/default.ttf', bytes: 4, revision: '9f64a747e1b97f13' },
              { path: 'fonts/symbol.ttf', bytes: 4, revision: '9f64a747e1b97f13' },
              { path: 'fonts/formerly-built-in.ttf', bytes: 4, revision: '9f64a747e1b97f13' },
            ],
          }),
        );
      }
      return response(null);
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);

    expect(controller.assets.map((asset) => asset.path)).toContain('fonts/default.ttf');
    expect(controller.assets.map((asset) => asset.path)).toContain('fonts/symbol.ttf');
    expect(controller.assets.map((asset) => asset.path)).not.toContain('fonts/formerly-built-in.ttf');
    expect(JSON.parse(window.localStorage.getItem('onlyoffice-browser:installed-fonts') || 'null')).toEqual({
      version: 2,
      downloaded: [],
    });
  });

  it('does not report the whole runtime complete after downloading one optional font family', async () => {
    const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-28T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'optional-font-v1' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(
          JSON.stringify({
            defaultFonts: [],
            builtInFonts: ['fonts/symbol.ttf'],
            fontFamilies: [{ name: 'Aptos', paths: ['fonts/aptos.ttf'] }],
            assets: [
              {
                path: 'sdkjs/common/Images/fonts_thumbnail.png',
                bytes: 4,
                revision: '9f64a747e1b97f13',
              },
              {
                path: 'fonts/aptos.ttf',
                bytes: 4,
                revision: '9f64a747e1b97f13',
                families: ['Aptos'],
              },
              {
                path: 'fonts/symbol.ttf',
                bytes: 4,
                revision: '9f64a747e1b97f13',
                families: ['OpenSymbol'],
              },
            ],
          }),
        );
      }
      return response(new Uint8Array([1, 2, 3, 4]));
    });
    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);

    const result = await controller.downloadFontFamily('aptos', () => undefined);

    expect(result.phase).toBe('ready');
    expect(result.completedFiles).toBe(4);
    expect(result.totalFiles).toBe(runtimeAssets.length + 4);
    expect(
      fetchMock.mock.calls.some(([input]) => String(input).includes('sdkjs/common/Images/fonts_thumbnail.png')),
    ).toBe(true);
    expect(fetchMock.mock.calls.some(([input]) => String(input).includes('fonts/symbol.ttf'))).toBe(true);
    expect(controller.listFonts()).toEqual([
      expect.objectContaining({
        name: 'Aptos',
        downloaded: true,
      }),
    ]);
    expect(controller.isComplete()).toBe(false);
  });

  it('precisely removes an optional family from the parent font package cache', async () => {
    const cacheStorage = memoryCacheStorage();
    const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-28T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'font-remove-v1' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(
          JSON.stringify({
            defaultFonts: ['fonts/default.ttf'],
            builtInFonts: [],
            fontFamilies: [
              { name: 'Microsoft YaHei', paths: ['fonts/default.ttf'] },
              { name: 'Arial', paths: ['fonts/arial.ttf'] },
            ],
            assets: [
              { path: 'fonts/default.ttf', bytes: 4, revision: '9f64a747e1b97f13' },
              { path: 'fonts/arial.ttf', bytes: 4, revision: '9f64a747e1b97f13' },
            ],
          }),
        );
      }
      return response(new Uint8Array([1, 2, 3, 4]));
    });
    const controller = await RuntimeCacheController.create(
      window.localStorage,
      fetchMock as unknown as typeof fetch,
      cacheStorage.storage,
    );

    await controller.downloadFontFamily('arial', () => undefined);
    expect(controller.listFonts().find((font) => font.name === 'Arial')).toMatchObject({
      downloaded: true,
      removable: true,
    });

    await controller.uninstallFontFamily('arial');

    expect(controller.listFonts().find((font) => font.name === 'Arial')).toMatchObject({
      downloaded: false,
      removable: true,
    });
    expect(controller.assets.some((asset) => asset.path === 'fonts/arial.ttf')).toBe(false);
    expect(cacheStorage.deleted.some((url) => url.includes('/fonts/arial.ttf?__oobv=9f64a747e1b97f13'))).toBe(true);
    expect(JSON.parse(window.localStorage.getItem('onlyoffice-browser:installed-fonts') || 'null')).toEqual({
      version: 2,
      downloaded: [],
    });
  });

  it('recovers from a stale v1 manifest cached by an older service worker', async () => {
    let runtimeRequests = 0;
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json' && init?.method === 'HEAD') {
        return response(null, { headers: { 'X-OnlyOffice-Asset-Version': 'recovered' } });
      }
      if (url.pathname === '/onlyoffice-runtime-assets.json' && runtimeRequests++ > 0) {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-28T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'recovered' } },
        );
      }
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(JSON.stringify({ version: 1, selected: 4 }));
      }
      return response(JSON.stringify({ fonts: [] }));
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    expect(controller.version).toBe('recovered');
    expect(controller.getProgress().totalFiles).toBe(5);
    expect(fetchMock.mock.calls.some(([input]) => String(input).includes('__cache_status='))).toBe(true);
  });

  it('checks only the release header when a complete inventory is already stored', async () => {
    const assets = runtimeAssets.map((asset) => ({
      path: asset.path,
      bytes: asset.bytes,
      category: asset.pack,
      revision: asset.revision,
    }));
    window.localStorage.setItem(
      'onlyoffice-browser:shared-runtime-cache',
      JSON.stringify({
        version: 'asset-v1',
        completed: assets.map((asset) => asset.path),
        assets,
        fontCatalog: [],
        fontFamilies: [],
        requiredFonts: [],
      }),
    );
    const fetchMock = vi.fn(async (_input: RequestInfo | URL, _init?: RequestInit) =>
      response(null, { headers: { 'X-OnlyOffice-Asset-Version': 'asset-v1' } }),
    );

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);

    expect(controller.isComplete()).toBe(true);
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0]?.[1]).toMatchObject({
      method: 'HEAD',
      cache: 'no-cache',
      credentials: 'omit',
    });
  });

  it('detects a corrupt cached body and repairs only that asset from the network', async () => {
    const assets = runtimeAssets.map((asset) => ({
      path: asset.path,
      bytes: asset.bytes,
      category: asset.pack,
      revision: asset.revision,
    }));
    window.localStorage.setItem(
      'onlyoffice-browser:shared-runtime-cache',
      JSON.stringify({
        version: 'asset-v1',
        completed: assets.map((asset) => asset.path),
        assets,
        lastVerifiedAt: 0,
        fontCatalog: [],
        fontFamilies: [],
        requiredFonts: [],
      }),
    );
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (init?.method === 'HEAD') {
        return response(null, { headers: { 'X-OnlyOffice-Asset-Version': 'asset-v1' } });
      }
      if (init?.cache === 'only-if-cached' && url.pathname.endsWith('/word.js')) {
        return response(new Uint8Array([9, 9, 9, 9]));
      }
      return response(new Uint8Array([1, 2, 3, 4]));
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    expect(controller.shouldCheckHealth()).toBe(true);
    const result = await controller.checkHealth(() => undefined);

    expect(result.phase).toBe('complete');
    expect(controller.isComplete()).toBe(true);
    expect(
      fetchMock.mock.calls.some(([input, init]) => String(input).includes('/word.js') && init?.cache === 'reload'),
    ).toBe(true);
    expect(controller.shouldCheckHealth()).toBe(false);
  });

  it('repairs an asset evicted from the HTTP cache', async () => {
    const asset = {
      path: runtimeAssets[0].path,
      bytes: runtimeAssets[0].bytes,
      category: runtimeAssets[0].pack,
      revision: runtimeAssets[0].revision,
    };
    window.localStorage.setItem(
      'onlyoffice-browser:shared-runtime-cache',
      JSON.stringify({
        version: 'asset-v1',
        completed: [asset.path],
        assets: [asset],
        lastVerifiedAt: 0,
        fontCatalog: [],
        fontFamilies: [],
        requiredFonts: [],
      }),
    );
    const fetchMock = vi.fn(async (_input: RequestInfo | URL, init?: RequestInit) => {
      if (init?.method === 'HEAD') {
        return response(null, { headers: { 'X-OnlyOffice-Asset-Version': 'asset-v1' } });
      }
      if (init?.cache === 'only-if-cached') return response(null, { status: 504 });
      return response(new Uint8Array([1, 2, 3, 4]));
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    const result = await controller.checkHealth(() => undefined);

    expect(result.phase).toBe('complete');
    expect(fetchMock.mock.calls.some(([, init]) => init?.cache === 'reload')).toBe(true);
  });

  it('verifies installed fonts from the dedicated Cache Storage before using the network', async () => {
    const cacheStorage = memoryCacheStorage();
    let thumbnailOnline = true;
    const thumbnailPath = '/sdkjs/common/Images/fonts_thumbnail@1.75x.png';
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-29T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'font-cache-v1' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(
          JSON.stringify({
            defaultFonts: [],
            builtInFonts: [],
            fontFamilies: [],
            assets: [
              {
                path: thumbnailPath.slice(1),
                bytes: 4,
                revision: '9f64a747e1b97f13',
              },
            ],
          }),
        );
      }
      if (url.pathname === thumbnailPath && !thumbnailOnline) return response(null, { status: 404 });
      return response(new Uint8Array([1, 2, 3, 4]));
    });
    const controller = await RuntimeCacheController.create(
      window.localStorage,
      fetchMock as unknown as typeof fetch,
      cacheStorage.storage,
    );
    await controller.loadAll(() => undefined);
    thumbnailOnline = false;
    fetchMock.mockClear();

    const result = await controller.checkHealth(() => undefined);

    expect(result.phase).toBe('complete');
    expect(result.failures).toEqual([]);
    expect(
      fetchMock.mock.calls.some(([input]) => new URL(String(input), window.location.origin).pathname === thumbnailPath),
    ).toBe(false);
  });

  it('reports a missing asset precisely and completes a later repair when it returns', async () => {
    let thumbnailOnline = true;
    const thumbnailPath = '/sdkjs/common/Images/fonts_thumbnail@1.75x.png';
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return response(
          JSON.stringify({
            version: 2,
            generatedAt: '2026-07-29T00:00:00.000Z',
            assets: runtimeAssets,
          }),
          { headers: { 'X-OnlyOffice-Asset-Version': 'missing-font-v1' } },
        );
      }
      if (url.pathname === '/onlyoffice-browser-font-assets.json') {
        return response(
          JSON.stringify({
            defaultFonts: [],
            builtInFonts: [],
            fontFamilies: [],
            assets: [
              {
                path: thumbnailPath.slice(1),
                bytes: 4,
                revision: '9f64a747e1b97f13',
              },
            ],
          }),
        );
      }
      if (url.pathname === thumbnailPath && !thumbnailOnline) return response(null, { status: 404 });
      if (init?.method === 'HEAD') return response(null, { headers: { 'content-length': '4' } });
      return response(new Uint8Array([1, 2, 3, 4]));
    });
    const controller = await RuntimeCacheController.create(
      window.localStorage,
      fetchMock as unknown as typeof fetch,
      undefined,
    );
    await controller.loadAll(() => undefined);
    thumbnailOnline = false;

    const failed = await controller.checkHealth(() => undefined);

    expect(failed).toMatchObject({
      phase: 'error',
      failedFiles: 1,
      failures: [{ path: thumbnailPath.slice(1), reason: expect.stringContaining('HTTP 404') }],
    });
    expect(controller.isComplete()).toBe(false);

    thumbnailOnline = true;
    const repaired = await controller.checkHealth(() => undefined);
    expect(repaired.phase).toBe('complete');
    expect(repaired.failures).toEqual([]);
    expect(controller.isComplete()).toBe(true);
  });
});
