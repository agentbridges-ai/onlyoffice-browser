import { afterEach, describe, expect, it, vi } from 'vitest';

// Executable deployment helper; intentionally shipped as ESM rather than compiled TypeScript.
// @ts-expect-error JavaScript build script has no declaration output.
import { fetchWithRetry } from '../../scripts/hydrate-cloudflare-matrix-fonts.mjs';

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('Cloudflare font release hydration', () => {
  it('retries when the response body terminates after successful headers', async () => {
    const expected = new Uint8Array([1, 2, 3]).buffer;
    const fetchMock = vi
      .fn()
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: vi.fn().mockRejectedValue(new TypeError('terminated')),
      })
      .mockResolvedValueOnce({
        ok: true,
        arrayBuffer: vi.fn().mockResolvedValue(expected),
      });
    vi.stubGlobal('fetch', fetchMock);

    const bytes = await fetchWithRetry(
      'https://onlyoffice.example.test/font.ttf',
      (response: Response) => response.arrayBuffer(),
      [0, 0],
    );

    expect(new Uint8Array(bytes)).toEqual(new Uint8Array([1, 2, 3]));
    expect(fetchMock).toHaveBeenCalledTimes(2);
  });

  it('reports the final body error after all attempts are exhausted', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue({
        ok: true,
        arrayBuffer: vi.fn().mockRejectedValue(new TypeError('terminated')),
      }),
    );

    await expect(
      fetchWithRetry(
        'https://onlyoffice.example.test/font.ttf',
        (response: Response) => response.arrayBuffer(),
        [0, 0],
      ),
    ).rejects.toThrow('Unable to download https://onlyoffice.example.test/font.ttf: TypeError: terminated');
  });
});
