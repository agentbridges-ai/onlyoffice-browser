import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { EventEmitter } from 'node:events';
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

const {
  hashFile,
  loadReleasePublication,
  loadV5ManifestPublication,
  planIncrementalRemoteVerification,
  verifyLocalRelease,
  verifyObjects,
} = publicationVerifier;

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
    sourceCommit: 'a'.repeat(40),
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
      expectedSourceCommit: 'a'.repeat(40),
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
    expect(() =>
      loadReleasePublication(output, {
        expectedPackageVersion: '0.5.7',
        expectedSourceCommit: 'b'.repeat(40),
        fastCdcEvidenceMode: 'forbid',
      }),
    ).toThrow(/sourceCommit/);
  });

  it('derives a complete remote audit plan from one immutable v5 manifest', () => {
    const { output, manifest } = fixture();
    const manifestFile = path.join(output, 'releases', manifest.releaseId, 'manifest.json');
    const expectedManifestSha256 = crypto.createHash('sha256').update(fs.readFileSync(manifestFile)).digest('hex');
    const publication = loadV5ManifestPublication(manifestFile, {
      releaseId: manifest.releaseId,
      expectedManifestSha256,
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'automatic',
    });

    expect(publication.releaseId).toBe(manifest.releaseId);
    expect(publication.objects.map((object: { kind: string }) => object.kind)).toEqual(
      expect.arrayContaining(['whole', 'package', 'package-segment', 'manifest-v5']),
    );
    expect(publication.objects.map((object: { kind: string }) => object.kind)).not.toContain('manifest-v4');
    expect(() =>
      loadV5ManifestPublication(manifestFile, {
        releaseId: manifest.releaseId,
        expectedManifestSha256: '0'.repeat(64),
      }),
    ).toThrow(/expected pointer digest/);
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

  it('hashes only new CAS objects while checking reused objects by inventory and size', () => {
    const { output } = fixture();
    const publication = loadReleasePublication(output, {
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'forbid',
    });
    const beforeInventory = new Map(
      publication.objects.map((object: { key: string; bytes: number }) => [object.key, { bytes: object.bytes }]),
    );
    const changed = publication.objects[0];
    beforeInventory.delete(changed.key);
    const afterInventory = new Map(
      publication.objects.map((object: { key: string; bytes: number }) => [object.key, { bytes: object.bytes }]),
    );

    const plan = planIncrementalRemoteVerification(publication.objects, beforeInventory, afterInventory);
    expect(plan.toHash).toEqual([changed]);
    expect(plan.reusedObjects).toBe(publication.objects.length - 1);
    expect(plan.reusedBytes).toBe(
      publication.objects.reduce((total: number, object: { bytes: number }) => total + object.bytes, 0) - changed.bytes,
    );
  });

  it('verifies a partial Range materialization without reading intentionally reused local objects', () => {
    const { output, manifest } = fixture();
    const publication = loadReleasePublication(output, {
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'forbid',
    });
    const inventoryPath = path.join(temp(), 'r2-before.json');
    const fakeRcloneInventoryPath = path.join(temp(), 'r2-lsf.txt');
    fs.writeFileSync(
      inventoryPath,
      JSON.stringify(
        publication.objects.map((object: { key: string; bytes: number }) => ({
          Path: object.key,
          Size: object.bytes,
        })),
      ),
    );
    fs.writeFileSync(
      fakeRcloneInventoryPath,
      publication.objects.map((object: { key: string; bytes: number }) => `${object.key}\t${object.bytes}`).join('\n') +
        '\n',
    );
    fs.rmSync(path.join(output, 'blobs'), { recursive: true, force: true });
    fs.rmSync(path.join(output, 'packages'), { recursive: true, force: true });
    fs.rmSync(path.join(output, 'segments'), { recursive: true, force: true });
    const fakeRclone = path.join(temp(), 'fake-rclone.mjs');
    fs.writeFileSync(
      fakeRclone,
      '#!/usr/bin/env node\n' +
        "import fs from 'node:fs';\n" +
        "if (process.argv[2] !== 'lsf') process.exit(2);\n" +
        'process.stdout.write(fs.readFileSync(process.env.FAKE_RCLONE_INVENTORY));\n',
    );
    fs.chmodSync(fakeRclone, 0o755);
    const result = spawnSync(
      process.execPath,
      [
        'scripts/verify-release-publication.mjs',
        '--release-root',
        output,
        '--expected-package-version',
        '0.5.7',
        '--fastcdc-evidence-mode',
        'forbid',
        '--remote',
        'r2:test',
        '--remote-verification-mode',
        'incremental',
        '--skip-local-verification',
        '--remote-inventory',
        inventoryPath,
        '--rclone-bin',
        fakeRclone,
      ],
      {
        cwd: path.resolve('.'),
        env: {
          ...process.env,
          FAKE_RCLONE_INVENTORY: fakeRcloneInventoryPath,
        },
        encoding: 'utf8',
      },
    );
    expect(result.status, result.stderr || result.stdout).toBe(0);
    expect(result.stdout).toContain(`Verified ${publication.objects.length} remote immutable objects`);
  });

  it('streams the R2 inventory instead of buffering lsjson output', async () => {
    const child = new EventEmitter() as EventEmitter & {
      stdout: EventEmitter;
      stderr: EventEmitter;
    };
    child.stdout = new EventEmitter();
    child.stderr = new EventEmitter();
    const spawnImpl = vi.fn((_binary: string, args: string[]) => {
      expect(args).toEqual([
        'lsf',
        'r2:test',
        '--recursive',
        '--files-only',
        '--format',
        'ps',
        '--separator',
        '\t',
        '--disable',
        'ListR',
      ]);
      queueMicrotask(() => {
        child.stdout.emit('data', Buffer.from('blobs/sha256/a\t12\nreleases/r1/manifest.json\t34'));
        child.emit('close', 0, null);
      });
      return child;
    });

    const inventory = await publicationVerifier.inspectRcloneInventory('r2:test', { spawnImpl });
    expect(inventory).toEqual(
      new Map([
        ['blobs/sha256/a', { bytes: 12 }],
        ['releases/r1/manifest.json', { bytes: 34 }],
      ]),
    );
  });

  it('fails closed when the post-upload inventory is missing or has a wrong size', () => {
    const { output } = fixture();
    const publication = loadReleasePublication(output, {
      expectedPackageVersion: '0.5.7',
      fastCdcEvidenceMode: 'forbid',
    });
    const beforeInventory = new Map();
    const afterInventory = new Map(
      publication.objects.map((object: { key: string; bytes: number }) => [object.key, { bytes: object.bytes }]),
    );
    afterInventory.set(publication.objects[0].key, { bytes: publication.objects[0].bytes + 1 });

    expect(() => planIncrementalRemoteVerification(publication.objects, beforeInventory, afterInventory)).toThrow(
      /wrong size/,
    );
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
    const { root, output } = fixture();
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
    const workerVersionId = 'dc8dcd28-271b-4367-9840-6c244f84cb40';
    let includeRuntimeWorkerVersion = true;
    const respond = (body?: BodyInit | null, init?: ResponseInit, includeWorkerVersion = true) => {
      const response = new Response(body, init);
      if (includeWorkerVersion) response.headers.set('X-OnlyOffice-Worker-Version', workerVersionId);
      return response;
    };
    const fetchImpl = vi.fn(async (input: string | URL | Request, init?: RequestInit) => {
      expect(new Headers(init?.headers).get('cloudflare-workers-version-overrides')).toBe(
        `onlyoffice-browser-runtime="${workerVersionId}"`,
      );
      const url = new URL(typeof input === 'string' ? input : input instanceof URL ? input : input.url);
      if (url.pathname.endsWith('/office-host.html')) {
        return respond(
          '<html>host bootstrap</html>',
          {
            status: 200,
            headers: {
              'Content-Type': 'text/html; charset=utf-8',
              'Cache-Control': 'public, max-age=31536000, immutable, no-transform',
              'X-OnlyOffice-Asset-Version': publication.releaseId,
              'Origin-Agent-Cluster': '?1',
            },
          },
          false,
        );
      }
      if (url.pathname.endsWith('/resource-broker.html')) {
        return respond(
          '<html>broker bootstrap</html>',
          {
            status: 200,
            headers: { 'Content-Security-Policy': csp },
          },
          false,
        );
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
          return respond(
            null,
            {
              status: 200,
              headers: { ...shared, 'Content-Length': String(object.bytes) },
            },
            false,
          );
        }
        if (range === `bytes=${object.bytes}-`) {
          return respond(
            null,
            {
              status: 416,
              headers: { ...shared, 'Content-Range': `bytes */${object.bytes}` },
            },
            false,
          );
        }
        return respond(
          objectBytes.subarray(0, 8),
          {
            status: 206,
            headers: {
              ...shared,
              'Content-Length': '8',
              'Content-Range': `bytes 0-7/${object.bytes}`,
            },
          },
          false,
        );
      }
      if (url.pathname.startsWith('/channels/')) {
        const pointer = url.pathname === '/channels/stable-v5.json' ? publication.stableV5 : publication.stableV4;
        return respond(
          JSON.stringify(pointer),
          {
            headers: { 'Cache-Control': 'no-store', 'Content-Type': 'application/json' },
          },
          false,
        );
      }
      if (url.pathname === publication.stableV5.manifestUrl || url.pathname === publication.stableV4.manifestUrl) {
        return respond(fs.readFileSync(path.join(output, ...url.pathname.slice(1).split('/'))), undefined, false);
      }
      if (url.pathname === '/onlyoffice-runtime-assets.json') {
        return respond(
          fs.readFileSync(path.join(root, 'onlyoffice-runtime-assets.json')),
          {
            headers: {
              'Cache-Control': 'public, max-age=0, must-revalidate',
              'X-OnlyOffice-Asset-Version': publication.releaseId,
            },
          },
          includeRuntimeWorkerVersion,
        );
      }
      if (url.pathname === '/index.html') {
        return respond(
          fs.readFileSync(path.join(root, 'index.html')),
          {
            headers: {
              'Cache-Control': 'public, max-age=0, must-revalidate',
              'Content-Type': 'text/html; charset=utf-8',
              'X-OnlyOffice-Asset-Version': publication.releaseId,
            },
          },
          false,
        );
      }
      if (url.pathname === '/sw.js') {
        return respond(
          fs.readFileSync(path.join(root, 'sw.js')),
          {
            headers: {
              'Cache-Control': 'public, max-age=0, must-revalidate',
              'Content-Type': 'text/javascript; charset=utf-8',
              'X-OnlyOffice-Asset-Version': publication.releaseId,
            },
          },
          false,
        );
      }
      return respond('not found', { status: 404 });
    });

    await expect(
      verifyReleaseHttp(publication, {
        canonicalOrigin: 'https://onlyoffice.getpi.work',
        editorOrigin: 'https://office-editor-github-actions-smoke.getpi.work',
        verifyPointers: true,
        verifyStableRoot: true,
        fetchImpl,
        workerVersionId,
      }),
    ).resolves.toMatchObject({
      releaseId: publication.releaseId,
      object: object.key,
      pointers: true,
      stableRoot: true,
    });
    expect(fetchImpl).toHaveBeenCalled();

    includeRuntimeWorkerVersion = false;
    await expect(
      verifyReleaseHttp(publication, {
        canonicalOrigin: 'https://onlyoffice.getpi.work',
        editorOrigin: 'https://office-editor-github-actions-smoke.getpi.work',
        verifyStableRoot: true,
        fetchImpl,
        workerVersionId,
      }),
    ).rejects.toThrow('Worker version x-onlyoffice-worker-version mismatch: missing');
  });
});
