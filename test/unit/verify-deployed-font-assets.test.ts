import { describe, expect, it, vi } from 'vitest';
// @ts-expect-error Deployment scripts are intentionally shipped as plain Node ESM.
import { verifyDeployedFontAssets } from '../../scripts/verify-deployed-font-assets.mjs';

const manifestUrl = 'https://onlyoffice.example.test/onlyoffice-browser-font-assets.json';

function manifestResponse() {
  return new Response(
    JSON.stringify({
      assets: [
        { path: 'fonts/default.ttf', bytes: 4, revision: 'font-revision' },
        {
          path: 'sdkjs/common/Images/fonts_thumbnail@1.75x.png',
          bytes: 8,
          revision: 'thumbnail-revision',
        },
      ],
    }),
  );
}

describe('verifyDeployedFontAssets', () => {
  it('checks every manifest asset at its immutable revision', async () => {
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const pathname = new URL(String(input)).pathname;
      if (pathname.endsWith('onlyoffice-browser-font-assets.json')) return manifestResponse();
      return new Response(null, {
        status: 206,
        headers: {
          'content-range': `bytes 0-0/${pathname.endsWith('default.ttf') ? '4' : '8'}`,
        },
      });
    });

    await expect(
      verifyDeployedFontAssets({ manifestUrl, fetchImpl: fetchMock as typeof fetch, concurrency: 2 }),
    ).resolves.toEqual({ checked: 2 });
    expect(fetchMock).toHaveBeenCalledTimes(3);
    expect(String(fetchMock.mock.calls[2]?.[0])).toContain('__oobv=');
  });

  it('reports every missing or truncated object before deployment proceeds', async () => {
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const pathname = new URL(String(input)).pathname;
      if (pathname.endsWith('onlyoffice-browser-font-assets.json')) return manifestResponse();
      return pathname.endsWith('default.ttf')
        ? new Response(null, { status: 416 })
        : new Response(null, { status: 206, headers: { 'content-range': 'bytes 0-0/7' } });
    });

    await expect(verifyDeployedFontAssets({ manifestUrl, fetchImpl: fetchMock as typeof fetch })).rejects.toThrow(
      [
        'Deployed font assets are incomplete:',
        'fonts/default.ttf: HTTP 416',
        'sdkjs/common/Images/fonts_thumbnail@1.75x.png: expected 8 bytes, received 7',
      ].join('\n'),
    );
  });
});
