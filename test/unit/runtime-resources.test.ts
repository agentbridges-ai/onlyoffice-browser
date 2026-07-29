import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createOfficeRuntimeResourceManager,
  type OfficeRuntimeResourceSnapshot,
} from '../../src/lib/runtime-resources';

const revision = '9f64a747e1b97f13';

function response(body: BodyInit | null, init: ResponseInit = {}): Response {
  return new Response(body, { status: 200, ...init });
}

function resourceFetch() {
  return vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input), window.location.origin);
    if (url.pathname === '/onlyoffice-runtime-assets.json' && init?.method === 'HEAD') {
      return response(null, { headers: { 'X-OnlyOffice-Asset-Version': 'resource-v1' } });
    }
    if (url.pathname === '/onlyoffice-runtime-assets.json') {
      return response(
        JSON.stringify({
          version: 2,
          generatedAt: '2026-07-28T00:00:00.000Z',
          assets: [{ path: 'sdkjs/word/word.js', bytes: 4, pack: 'word', revision }],
        }),
      );
    }
    if (url.pathname === '/onlyoffice-browser-font-assets.json') {
      return response(
        JSON.stringify({
          defaultFonts: ['fonts/dengxian.ttf'],
          builtInFonts: [],
          fontFamilies: [
            { name: 'DengXian', paths: ['fonts/dengxian.ttf'] },
            { name: 'Microsoft YaHei', paths: ['fonts/yahei.ttf'] },
          ],
          assets: [
            {
              path: 'fonts/dengxian.ttf',
              bytes: 4,
              revision,
              families: ['DengXian'],
            },
            {
              path: 'fonts/yahei.ttf',
              bytes: 4,
              revision,
              families: ['Microsoft YaHei'],
            },
          ],
        }),
      );
    }
    return response(new Uint8Array([1, 2, 3, 4]));
  });
}

describe('OfficeRuntimeResourceManager', () => {
  beforeEach(() => {
    window.localStorage.clear();
  });

  it('publishes verified optional fonts and keeps the required default non-removable', async () => {
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: resourceFetch() as unknown as typeof fetch,
      cacheStorage: undefined,
    });
    const snapshots: OfficeRuntimeResourceSnapshot[] = [];
    manager.subscribe((snapshot) => snapshots.push(snapshot));

    expect(manager.getSnapshot().fonts).toEqual([
      expect.objectContaining({ name: 'DengXian', removable: false }),
      expect.objectContaining({ name: 'Microsoft YaHei', downloaded: false, removable: true }),
    ]);
    expect(manager.getSnapshot()).toMatchObject({
      packageVersion: '0.4.2',
      assetVersion: 'resource-v1',
      readiness: 'needs-download',
      packs: expect.arrayContaining([expect.objectContaining({ id: 'word', ready: false })]),
    });
    expect(manager.getSnapshot()).toBe(manager.getSnapshot());

    await manager.downloadFontFamily('microsoft yahei');

    expect(manager.getVerifiedFontPaths()).toContain('fonts/yahei.ttf');
    expect(manager.getSnapshot()).toMatchObject({
      operation: null,
      error: null,
      fonts: [
        expect.objectContaining({ name: 'DengXian', removable: false }),
        expect.objectContaining({ name: 'Microsoft YaHei', downloaded: true, removable: true }),
      ],
    });
    expect(snapshots.some((snapshot) => snapshot.operation === 'download-font')).toBe(true);
    expect(snapshots.at(-1)).toBe(manager.getSnapshot());

    await manager.uninstallFontFamily('microsoft yahei');
    expect(manager.getVerifiedFontPaths()).not.toContain('fonts/yahei.ttf');
  });

  it('prepares only the packs needed by the active document', async () => {
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });

    await manager.prepareForDocumentType('word');

    expect(manager.getSnapshot().packs.find((pack) => pack.id === 'word')).toMatchObject({ ready: true });
    expect(manager.getSnapshot().operation).toBeNull();
  });

  it('coalesces identical operations and serializes different mutations', async () => {
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });

    const first = manager.downloadFontFamily('microsoft yahei');
    const duplicate = manager.downloadFontFamily('microsoft yahei');
    const remove = manager.uninstallFontFamily('microsoft yahei');

    expect(duplicate).toBe(first);
    await first;
    await remove;
    expect(manager.getSnapshot().fonts.find((font) => font.name === 'Microsoft YaHei')).toMatchObject({
      downloaded: false,
    });
  });

  it('loads and verifies resources from an embedding-provided canonical asset origin', async () => {
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
      assetBaseUrl: 'https://onlyoffice.example.test/runtime/',
    });

    await manager.downloadFontFamily('microsoft yahei');

    expect(
      fetchMock.mock.calls.some(
        ([input]) => String(input) === 'https://onlyoffice.example.test/onlyoffice-runtime-assets.json',
      ),
    ).toBe(true);
    expect(fetchMock.mock.calls.find(([input]) => String(input).includes('/fonts/yahei.ttf'))?.[1]).toMatchObject({
      mode: 'cors',
      credentials: 'omit',
    });
  });

  it('surfaces integrity failures without losing the last usable catalog', async () => {
    const fetchMock = resourceFetch();
    fetchMock.mockImplementation(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname.endsWith('/word.js') && init?.cache !== undefined) {
        return response(new Uint8Array([9, 9, 9, 9]));
      }
      return resourceFetch()(input, init);
    });
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });

    const result = await manager.loadAll();

    expect(result.phase).toBe('error');
    expect(manager.getSnapshot().progress.failedFiles).toBe(1);
    expect(manager.getSnapshot().fonts).toHaveLength(2);
  });
});
