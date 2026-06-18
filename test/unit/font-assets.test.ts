import { afterEach, describe, expect, it, vi } from 'vitest';

import {
  assertGeneratedFontAssetsAvailable,
  fetchGeneratedFontAssetsManifest,
  fetchRuntimeBinaryAsset,
} from '../../src/lib/font-assets';

const MANIFEST = {
  version: 1,
  allFonts: 'sdkjs/common/AllFonts.js',
  fontSelection: 'server/FileConverter/bin/font_selection.bin',
  fontThumbnails: ['sdkjs/common/Images/fonts_thumbnail.png'],
  fonts: ['fonts/000.ttf'],
};

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('generated font assets runtime checks', () => {
  it('requires the generated font asset manifest', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response('', { status: 404 })));

    await expect(fetchGeneratedFontAssetsManifest()).rejects.toThrow('OnlyOffice font assets are required');
  });

  it('rejects incomplete generated font asset manifests', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(new Response(JSON.stringify({ version: 1, fonts: [] }), { status: 200 })),
    );

    await expect(fetchGeneratedFontAssetsManifest()).rejects.toThrow('missing allFonts');
  });

  it('probes the canonical generated font asset paths before editor startup', async () => {
    const fetchMock = vi.fn((url: string | URL | Request, init?: RequestInit) => {
      const path = new URL(String(url), window.location.href).pathname;
      if (path === '/onlyoffice-browser-font-assets.json') {
        return Promise.resolve(new Response(JSON.stringify(MANIFEST), { status: 200 }));
      }
      if (init?.headers && (init.headers as Record<string, string>).Range === 'bytes=0-0') {
        return Promise.resolve(new Response(null, { status: 200 }));
      }
      return Promise.resolve(new Response('', { status: 404 }));
    });
    vi.stubGlobal('fetch', fetchMock);

    await expect(assertGeneratedFontAssetsAvailable()).resolves.toMatchObject(MANIFEST);
    expect(fetchMock).toHaveBeenCalledWith(expect.stringContaining('/sdkjs/common/AllFonts.js'), {
      cache: 'no-cache',
      headers: {
        Range: 'bytes=0-0',
      },
    });
    expect(fetchMock).toHaveBeenCalledWith(expect.stringContaining('/fonts/000.ttf'), {
      cache: 'no-cache',
      headers: {
        Range: 'bytes=0-0',
      },
    });
  });

  it('loads generated binary assets through the canonical runtime path', async () => {
    vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(new Uint8Array([1, 2, 3]), { status: 200 })));

    await expect(fetchRuntimeBinaryAsset('fonts/000.ttf')).resolves.toEqual(new Uint8Array([1, 2, 3]));
  });
});
