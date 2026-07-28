import { beforeEach, describe, expect, it, vi } from 'vitest';
import { RuntimeCacheController, type RuntimeCacheProgress } from '../../src/lib/runtime-cache';

const runtimeAssets = [
  { path: 'sdkjs/common/core.js', bytes: 4, pack: 'core' },
  { path: 'sdkjs/word/word.js', bytes: 4, pack: 'word' },
  { path: 'sdkjs/cell/cell.js', bytes: 4, pack: 'cell' },
  { path: 'sdkjs/slide/slide.js', bytes: 4, pack: 'slide' },
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
            allFonts: 'sdkjs/common/AllFonts.js',
            fontSelection: 'server/FileConverter/bin/font_selection.bin',
            fontSourceMap: 'onlyoffice-browser-font-source-map.json',
            fontThumbnails: ['sdkjs/common/Images/fonts_thumbnail.png'],
            fonts: ['fonts/000.ttf'],
            assets: [
              { path: 'sdkjs/common/AllFonts.js', bytes: 4 },
              { path: 'server/FileConverter/bin/font_selection.bin', bytes: 4 },
              { path: 'onlyoffice-browser-font-source-map.json', bytes: 4 },
              { path: 'sdkjs/common/Images/fonts_thumbnail.png', bytes: 4 },
              { path: 'fonts/000.ttf', bytes: 4 },
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

    expect(initial.totalFiles).toBe(10);
    expect(initial.categories.map((category) => [category.category, category.totalFiles])).toEqual([
      ['fonts', 6],
      ['core', 1],
      ['word', 1],
      ['cell', 1],
      ['slide', 1],
    ]);
    expect(initial.totalBytes).toBe(36);

    let finalProgress: RuntimeCacheProgress | undefined;
    await controller.loadAll((progress) => {
      finalProgress = progress;
    });
    expect(finalProgress?.phase).toBe('complete');
    expect(finalProgress?.completedFiles).toBe(10);
    expect(fetchMock.mock.calls.some(([, init]) => init?.cache === 'force-cache')).toBe(true);

    const restored = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    expect(restored.isComplete()).toBe(true);
    const restoredRequests = fetchMock.mock.calls.slice(-1);
    expect(restoredRequests).toHaveLength(1);
    expect(restoredRequests[0]?.[1]).toMatchObject({ method: 'HEAD', cache: 'no-cache' });
  });

  it('invalidates stored progress when the deployed asset version changes', async () => {
    window.localStorage.setItem(
      'onlyoffice-browser:shared-runtime-cache',
      JSON.stringify({
        version: 'old',
        completed: runtimeAssets.map((asset) => asset.path),
        assets: runtimeAssets.map((asset) => ({ ...asset, category: asset.pack })),
      }),
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
          { headers: { 'X-OnlyOffice-Asset-Version': 'new' } },
        );
      }
      return response(JSON.stringify({ fonts: [] }));
    });

    const controller = await RuntimeCacheController.create(window.localStorage, fetchMock as unknown as typeof fetch);
    expect(controller.isComplete()).toBe(false);
    expect(controller.getProgress().completedFiles).toBe(0);
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
    }));
    window.localStorage.setItem(
      'onlyoffice-browser:shared-runtime-cache',
      JSON.stringify({
        version: 'asset-v1',
        completed: assets.map((asset) => asset.path),
        assets,
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
});
