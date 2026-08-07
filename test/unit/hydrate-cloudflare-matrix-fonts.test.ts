import { afterEach, describe, expect, it, vi } from 'vitest';

// Executable deployment helper; intentionally shipped as ESM rather than compiled TypeScript.
// @ts-expect-error JavaScript build script has no declaration output.
import { downloadRangedAsset, fetchWithRetry, readRelease } from '../../scripts/hydrate-cloudflare-matrix-fonts.mjs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import crypto from 'node:crypto';

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('Cloudflare font release hydration', () => {
  it('prefers the v5 channel and follows its release manifest URL', async () => {
    const releaseId = 'v0.5.15-test';
    const fontManifest = {
      path: 'onlyoffice-browser-font-assets.json',
      sha256: 'a'.repeat(64),
      bytes: 3,
    };
    const fetchMock = vi.fn(async (url: string) => {
      if (url === 'https://onlyoffice.getpi.work/channels/stable-v5.json') {
        return new Response(
          JSON.stringify({
            version: 1,
            releaseId,
            manifestUrl: `/releases/${releaseId}/manifest.json`,
          }),
        );
      }
      if (url === `https://onlyoffice.getpi.work/releases/${releaseId}/manifest.json`) {
        return new Response(
          JSON.stringify({
            version: 5,
            releaseId,
            assets: [
              fontManifest,
              { path: 'fonts/aptos.ttf', sha256: 'b'.repeat(64), bytes: 4, profile: 'fonts-basic' },
            ],
          }),
        );
      }
      throw new Error(`unexpected URL: ${url}`);
    });
    vi.stubGlobal('fetch', fetchMock);

    await expect(readRelease()).resolves.toEqual({
      releaseId,
      assets: [fontManifest, { path: 'fonts/aptos.ttf', sha256: 'b'.repeat(64), bytes: 4, profile: 'fonts-basic' }],
    });
    expect(fetchMock.mock.calls.map(([url]) => url)).toEqual([
      'https://onlyoffice.getpi.work/channels/stable-v5.json',
      `https://onlyoffice.getpi.work/releases/${releaseId}/manifest.json`,
    ]);
  });

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

  it('downloads large fonts through resumable validated Range windows', async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-font-range-'));
    const body = Buffer.from('range-font-payload');
    const sha256 = crypto.createHash('sha256').update(body).digest('hex');
    const asset = { path: 'fonts/large.ttf', bytes: body.byteLength, sha256 };
    let firstRange = true;
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      const range =
        init?.headers instanceof Headers ? init.headers.get('Range') : (init?.headers as Record<string, string>)?.Range;
      const match = /^bytes=(\d+)-(\d+)$/.exec(range || '');
      if (!match) throw new Error('missing Range header');
      const start = Number(match[1]);
      const end = Number(match[2]);
      if (firstRange) {
        firstRange = false;
        return {
          ok: true,
          status: 206,
          headers: new Headers({ 'content-range': `bytes ${start}-${end}/${body.byteLength}` }),
          arrayBuffer: vi.fn().mockRejectedValue(new TypeError('terminated')),
        } as unknown as Response;
      }
      const value = body.subarray(start, end + 1);
      return new Response(value, {
        status: 206,
        headers: { 'content-range': `bytes ${start}-${end}/${body.byteLength}` },
      });
    });
    vi.stubGlobal('fetch', fetchMock);

    await expect(
      downloadRangedAsset('release-range', asset, { root, rangeBytes: 5, retryDelays: [0, 0] }),
    ).resolves.toBe(true);
    expect(fs.readFileSync(path.join(root, asset.path))).toEqual(body);
    expect(fs.existsSync(`${path.join(root, asset.path)}.download`)).toBe(false);
    expect(fs.existsSync(`${path.join(root, asset.path)}.download.json`)).toBe(false);
    expect(fetchMock).toHaveBeenCalledTimes(5);
    fs.rmSync(root, { recursive: true, force: true });
  });

  it('keeps a verified prefix and resumes after a later Range timeout', async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-font-resume-'));
    const body = Buffer.from('resume-font-payload');
    const sha256 = crypto.createHash('sha256').update(body).digest('hex');
    const asset = { path: 'fonts/resume.ttf', bytes: body.byteLength, sha256 };
    let failSecondWindow = true;
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      const headers = init?.headers as Record<string, string>;
      const match = /^bytes=(\d+)-(\d+)$/.exec(headers?.Range || '');
      if (!match) throw new Error('missing Range header');
      const start = Number(match[1]);
      const end = Number(match[2]);
      if (failSecondWindow && start >= 5) throw new TypeError('terminated');
      const value = body.subarray(start, end + 1);
      return new Response(value, {
        status: 206,
        headers: { 'content-range': `bytes ${start}-${end}/${body.byteLength}` },
      });
    });
    vi.stubGlobal('fetch', fetchMock);

    await expect(
      downloadRangedAsset('release-resume', asset, { root, rangeBytes: 5, retryDelays: [0] }),
    ).rejects.toThrow('Unable to download');
    expect(fs.readFileSync(path.join(root, `${asset.path}.download`))).toEqual(body.subarray(0, 5));

    failSecondWindow = false;
    await expect(downloadRangedAsset('release-resume', asset, { root, rangeBytes: 5, retryDelays: [0] })).resolves.toBe(
      true,
    );
    expect(fs.readFileSync(path.join(root, asset.path))).toEqual(body);
    expect(
      fetchMock.mock.calls.some(([, init]) => (init?.headers as Record<string, string>)?.Range === 'bytes=0-4'),
    ).toBe(true);
    expect(
      fetchMock.mock.calls.some(([, init]) => (init?.headers as Record<string, string>)?.Range === 'bytes=5-9'),
    ).toBe(true);
    fs.rmSync(root, { recursive: true, force: true });
  });
});
