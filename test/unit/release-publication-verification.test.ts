import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
// Executable Node release scripts intentionally have no declaration output.
// @ts-expect-error JavaScript build script has no declaration output.
import { buildRelease, computeStorageSetSha256 } from '../../scripts/build-release-manifest.mjs';
// @ts-expect-error JavaScript verification script has no declaration output.
import * as publicationVerifier from '../../scripts/verify-release-publication.mjs';
// @ts-expect-error JavaScript verification script has no declaration output.
import { verifyReleaseHttp } from '../../scripts/verify-release-http.mjs';

const { hashFile, loadReleasePublication, verifyLocalRelease, verifyObjects } = publicationVerifier;

const temporaryDirectories: string[] = [];

function temp(): string {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-publication-'));
  temporaryDirectories.push(directory);
  return directory;
}

function write(root: string, relative: string, value: string | Uint8Array): void {
  const file = path.join(root, relative);
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, value);
}

function fixture() {
  const root = temp();
  const output = temp();
  write(root, 'office-host.html', '<html>host bootstrap</html>');
  write(root, 'resource-broker.html', '<html>broker bootstrap</html>');
  write(root, 'index.html', '<html>standalone shell</html>');
  write(root, 'sw.js', 'self.addEventListener("fetch", () => {});');
  write(root, 'assets/officeHost-test.js', 'export const host = true;');
  write(root, 'assets/main-test.js', 'export const shell = true;');
  write(root, 'wasm/x2t/x2t.wasm', Buffer.from('0123456789abcdef'));
  write(
    root,
    'onlyoffice-runtime-assets.json',
    JSON.stringify({
      version: 2,
      assets: [{ path: 'wasm/x2t/x2t.wasm', pack: 'core', bytes: 16, revision: 'test' }],
    }),
  );
  const manifest = buildRelease({
    root,
    output,
    packageVersion: '0.5.7',
    x2tVersion: '9.3.0+2',
    x2tCommit: '1bb9b45a399f87ca162eea0c86abd4660f295469',
  });
  return { root, output, manifest };
}

function selectedHttpObject(publication: ReturnType<typeof loadReleasePublication>) {
  const object = publication.objects.find(
    (candidate: { bytes: number; kind: string }) =>
      ['whole', 'fastcdc', 'package-segment'].includes(candidate.kind) && candidate.bytes >= 8,
  );
  if (!object) throw new Error('fixture has no HTTP object');
  return object;
}

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

describe('Release Manifest v5 publication verification', () => {
  it('verifies v5/v4 pointers and every local immutable SHA-256 object', async () => {
    const { output, manifest } = fixture();
    const publication = loadReleasePublication(output, {
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'forbid',
    });

    expect(publication.releaseId).toBe(manifest.releaseId);
    expect(publication.manifest.version).toBe(5);
    expect(publication.compatibilityManifest.version).toBe(4);
    expect(publication.fastCdcAssets).toEqual([]);
    expect(publication.objects.map((object: { kind: string }) => object.kind)).toEqual(
      expect.arrayContaining(['whole', 'package', 'package-segment', 'manifest-v5', 'manifest-v4']),
    );

    const result = await verifyLocalRelease(publication, { concurrency: 3 });
    expect(result.objects).toBe(publication.objects.length);
    expect(result.bytes).toBeGreaterThan(0);
  });

  it('does not treat a successful transport read as verification when bytes or digest differ', async () => {
    const { output } = fixture();
    const publication = loadReleasePublication(output, {
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'forbid',
    });
    const first = publication.objects[0];
    const inspect = vi.fn(async (object: { key: string }) =>
      object.key === first.key
        ? { bytes: first.bytes, sha256: '0'.repeat(64) }
        : hashFile(path.join(output, ...object.key.split('/'))),
    );

    await expect(
      verifyObjects(publication.objects, inspect, {
        concurrency: 2,
        label: 'Remote immutable object',
      }),
    ).rejects.toThrow(/Remote immutable object verification failed/);
    expect(inspect).toHaveBeenCalled();
  });

  it('fails closed if a small whole-file CAS asset is unnecessarily FastCDC chunked', () => {
    const { output, manifest } = fixture();
    const releaseDirectory = path.join(output, 'releases', manifest.releaseId);
    const manifestPath = path.join(releaseDirectory, 'manifest.json');
    const value = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
    const asset = value.assets.find((candidate: { bytes: number }) => candidate.bytes > 0);
    asset.representations.fastcdc = {
      algorithm: 'fastcdc-v2020',
      minBytes: 64 * 1024,
      averageBytes: 256 * 1024,
      maxBytes: 1024 * 1024,
      normalization: 1,
      seed: 0,
      chunks: [{ offset: 0, bytes: asset.bytes, sha256: asset.sha256 }],
    };
    value.contentProtocol.storageSetSha256 = computeStorageSetSha256(value.package, value.assets);
    const manifestText = `${JSON.stringify(value, null, 2)}\n`;
    fs.writeFileSync(manifestPath, manifestText);
    const pointerPath = path.join(output, 'channels/stable-v5.json');
    const pointer = JSON.parse(fs.readFileSync(pointerPath, 'utf8'));
    pointer.manifestSha256 = crypto.createHash('sha256').update(manifestText).digest('hex');
    fs.writeFileSync(pointerPath, `${JSON.stringify(pointer, null, 2)}\n`);

    expect(() =>
      loadReleasePublication(output, {
        expectedPackageVersion: '0.5.7',
        fastCdcEvidenceMode: 'automatic',
      }),
    ).toThrow(/Small asset .* must remain a whole-file CAS object/);
  });
});

describe('release-specific production HTTP verification', () => {
  it('checks Host identity, Broker CSP, object 200/206/416 semantics, and both stable pointers', async () => {
    const { output } = fixture();
    const publication = loadReleasePublication(output, {
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'forbid',
    });
    const object = selectedHttpObject(publication);
    const objectFile = path.join(output, ...object.key.split('/'));
    const objectBytes = fs.readFileSync(objectFile);
    const csp = [
      "default-src 'none'",
      "script-src 'self'",
      "connect-src 'self'",
      "worker-src 'self'",
      "base-uri 'none'",
      "form-action 'none'",
      'frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work',
      'https://aries.getpi.work https://taurus.getpi.work https://gemini.getpi.work',
      'https://cancer.getpi.work https://leo.getpi.work https://virgo.getpi.work',
      'https://libra.getpi.work https://scorpio.getpi.work https://sagittarius.getpi.work',
      'https://capricorn.getpi.work https://aquarius.getpi.work https://pisces.getpi.work',
    ].join('; ');
    const fetchImpl = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      const url = new URL(typeof input === 'string' ? input : input instanceof URL ? input : input.url);
      if (url.pathname.endsWith('/office-host.html')) {
        return new Response('<html>host bootstrap</html>', {
          status: 200,
          headers: {
            'Content-Type': 'text/html; charset=utf-8',
            'Cache-Control': 'public, max-age=31536000, immutable, no-transform',
            'X-OnlyOffice-Asset-Version': publication.releaseId,
            'Origin-Agent-Cluster': '?1',
          },
        });
      }
      if (url.pathname.endsWith('/resource-broker.html')) {
        return new Response('<html>broker bootstrap</html>', {
          status: 200,
          headers: { 'Content-Security-Policy': csp },
        });
      }
      if (url.pathname.includes('/objects/')) {
        const range = new Headers(init?.headers).get('range');
        const shared = {
          'Accept-Ranges': 'bytes',
          'X-Content-SHA256': object.sha256,
          'X-OnlyOffice-Asset-Version': publication.releaseId,
          'Cache-Control': 'public, max-age=31536000, immutable',
          'Access-Control-Allow-Origin': '*',
          'Cross-Origin-Resource-Policy': 'cross-origin',
        };
        if (init?.method === 'HEAD') {
          return new Response(null, {
            status: 200,
            headers: { ...shared, 'Content-Length': String(object.bytes) },
          });
        }
        if (range === `bytes=${object.bytes}-`) {
          return new Response(null, {
            status: 416,
            headers: { ...shared, 'Content-Range': `bytes */${object.bytes}` },
          });
        }
        return new Response(objectBytes.subarray(0, 8), {
          status: 206,
          headers: {
            ...shared,
            'Content-Length': '8',
            'Content-Range': `bytes 0-7/${object.bytes}`,
          },
        });
      }
      if (url.pathname.startsWith('/channels/')) {
        const pointer = url.pathname === '/channels/stable-v5.json' ? publication.stableV5 : publication.stableV4;
        return Response.json(pointer, { headers: { 'Cache-Control': 'no-store' } });
      }
      if (url.pathname === publication.stableV5.manifestUrl || url.pathname === publication.stableV4.manifestUrl) {
        return new Response(fs.readFileSync(path.join(output, ...url.pathname.slice(1).split('/'))));
      }
      return new Response('not found', { status: 404 });
    });

    await expect(
      verifyReleaseHttp(publication, {
        canonicalOrigin: 'https://onlyoffice.getpi.work',
        editorOrigin: 'https://office-editor-github-actions-smoke.getpi.work',
        verifyPointers: true,
        fetchImpl,
      }),
    ).resolves.toMatchObject({
      releaseId: publication.releaseId,
      object: object.key,
      pointers: true,
    });
    expect(fetchImpl).toHaveBeenCalled();
  });
});
