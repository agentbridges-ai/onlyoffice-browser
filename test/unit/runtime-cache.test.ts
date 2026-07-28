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
});
