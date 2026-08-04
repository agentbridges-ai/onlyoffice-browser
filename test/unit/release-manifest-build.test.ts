import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
// The release builder is intentionally executable Node ESM rather than part of
// the browser TypeScript bundle.
// @ts-expect-error JavaScript build script has no declaration output.
import { buildRelease, chunkReleaseAssets, computeStorageSetSha256 } from '../../scripts/build-release-manifest.mjs';
import { parseReleaseManifest } from '../../src/lib/release-resources';

const temporaryDirectories: string[] = [];

function temp(): string {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-release-'));
  temporaryDirectories.push(directory);
  return directory;
}

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

describe('immutable release builder', () => {
  it('creates deterministic profile chunks with a 32 MiB ceiling', () => {
    const assets = [
      { path: 'a', bytes: 20, profile: 'word', chunk: '' },
      { path: 'b', bytes: 20, profile: 'word', chunk: '' },
      { path: 'c', bytes: 40, profile: 'word', chunk: '' },
    ] as never[];
    const chunks = chunkReleaseAssets(assets, 24, 32);
    expect(
      chunks.map((chunk: { bytes: number; paths: string[] }) => ({ bytes: chunk.bytes, paths: chunk.paths })),
    ).toEqual([
      { bytes: 20, paths: ['a'] },
      { bytes: 20, paths: ['b'] },
      { bytes: 40, paths: ['c'] },
    ]);
  });

  it('emits full hashes, immutable blobs, and a stable pointer', () => {
    const root = temp();
    const output = temp();
    const files = {
      'office-host.html': '<html>host</html>',
      'index.html': '<html>shell</html>',
      'sw.js': 'worker',
      'assets/officeHost-abcdefgh.js': 'host bundle',
      'assets/main-abcdefgh.js': 'shell bundle',
      'npm/public-api.js': 'not a deployed runtime asset',
      '.vite/manifest.json': 'not requested by the runtime',
      'wasm/x2t/x2t.wasm': 'x2t bytes',
      'sdkjs/word/word.js': 'word',
      'fonts/basic.ttf': 'basic font',
      'fonts/compat.ttf': 'compatibility font',
      'fonts/legacy-unlisted.ttf': 'legacy compatibility font',
      'server/FileConverter/bin/AllFonts.js': 'server font index',
      'wasm/x2t/x2t.wasm.br': 'precompressed duplicate',
    };
    for (const [relative, contents] of Object.entries(files)) {
      fs.mkdirSync(path.dirname(path.join(root, relative)), { recursive: true });
      fs.writeFileSync(path.join(root, relative), contents);
    }
    fs.writeFileSync(
      path.join(root, 'onlyoffice-runtime-assets.json'),
      JSON.stringify({
        version: 2,
        assets: [
          { path: 'sdkjs/word/word.js', pack: 'word', bytes: 4, revision: 'legacy' },
          { path: 'fonts/basic.ttf', pack: 'core', bytes: 10, revision: 'legacy' },
          { path: 'fonts/compat.ttf', pack: 'core', bytes: 18, revision: 'legacy' },
          { path: 'fonts/legacy-unlisted.ttf', pack: 'core', bytes: 25, revision: 'legacy' },
          { path: 'wasm/x2t/x2t.wasm.br', pack: 'core', bytes: 23, revision: 'legacy' },
        ],
      }),
    );
    fs.writeFileSync(
      path.join(root, 'onlyoffice-browser-font-assets.json'),
      JSON.stringify({
        version: 1,
        assets: [{ path: 'fonts/basic.ttf' }, { path: 'fonts/compat.ttf' }],
        defaultFonts: ['fonts/basic.ttf'],
      }),
    );
    const first = buildRelease({
      root,
      output,
      packageVersion: '0.4.0',
      sourceCommit: 'a'.repeat(40),
      x2tVersion: '9.3.0+1',
      x2tCommit: 'abc123',
    });
    const secondOutput = temp();
    const second = buildRelease({
      root,
      output: secondOutput,
      packageVersion: '0.4.0',
      sourceCommit: 'a'.repeat(40),
      x2tVersion: '9.3.0+1',
      x2tCommit: 'abc123',
    });
    expect(second.releaseId).toBe(first.releaseId);
    expect(first.version).toBe(5);
    expect(first.sourceCommit).toBe('a'.repeat(40));
    expect(first.protocolHostBuildId).toBe('office-host-0.4.0-r1');
    const releaseIdentity = Object.fromEntries(
      Object.entries(first).filter(([key]) => key !== 'version' && key !== 'releaseId'),
    );
    expect(first.releaseId).toBe(
      `v0.4.0-${crypto.createHash('sha256').update(JSON.stringify(releaseIdentity)).digest('hex').slice(0, 16)}`,
    );
    expect(() =>
      buildRelease({
        root,
        output: temp(),
        packageVersion: '0.4.0',
        sourceCommit: 'not-a-commit',
        x2tVersion: '9.3.0+1',
        x2tCommit: 'abc123',
      }),
    ).toThrow(/sourceCommit/);
    expect(first.contentProtocol).toMatchObject({
      version: 1,
      digest: 'sha256',
      cacheKeyFormat: 'canonical-sha256-v1',
      fastcdcPolicyId: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0',
    });
    expect(first.contentProtocol.storageSetSha256).toBe(computeStorageSetSha256(first.package, first.assets));
    expect(() => parseReleaseManifest(first)).not.toThrow();
    expect(first.package).toMatchObject({
      format: 'onlyoffice-pack-v1',
      path: 'office-resources.oobpack',
    });
    expect(second.package.sha256).toBe(first.package.sha256);
    expect(first.package.segments.reduce((sum: number, segment: { bytes: number }) => sum + segment.bytes, 0)).toBe(
      first.package.bytes,
    );
    expect(first.assets.every((item: { sha256: string }) => /^[a-f0-9]{64}$/.test(item.sha256))).toBe(true);
    expect(first.profiles.word).toContain('sdkjs/word/word.js');
    expect(first.profiles['fonts-basic']).toEqual(
      expect.arrayContaining(['fonts/basic.ttf', 'server/FileConverter/bin/AllFonts.js']),
    );
    expect(first.profiles['fonts-office-compat']).toEqual(['fonts/compat.ttf', 'fonts/legacy-unlisted.ttf']);
    expect(first.profiles.base).not.toEqual(
      expect.arrayContaining(['fonts/basic.ttf', 'fonts/compat.ttf', 'server/FileConverter/bin/AllFonts.js']),
    );
    expect(first.assets.map((item: { path: string }) => item.path)).not.toContain('npm/public-api.js');
    expect(first.assets.map((item: { path: string }) => item.path)).not.toContain('.vite/manifest.json');
    expect(first.assets.map((item: { path: string }) => item.path)).not.toContain('wasm/x2t/x2t.wasm.br');
    expect(fs.existsSync(path.join(output, 'releases', first.releaseId, 'manifest.json'))).toBe(true);
    expect(fs.existsSync(path.join(output, 'releases', first.releaseId, 'manifest-v4.json'))).toBe(true);
    const packPath = path.join(output, 'packages/sha256', `${first.package.sha256}.oobpack`);
    expect(fs.existsSync(packPath)).toBe(true);
    expect(fs.readFileSync(packPath).subarray(0, 8).toString()).toBe('OOBPACK1');
    expect(fs.statSync(packPath).size).toBe(first.package.bytes);
    for (const segment of first.package.segments as Array<{ id: string; sha256: string; bytes: number }>) {
      expect(segment.id).toBe(segment.sha256);
      expect(segment.sha256).toMatch(/^[a-f0-9]{64}$/);
      expect(fs.statSync(path.join(output, 'segments/sha256', segment.sha256)).size).toBe(segment.bytes);
    }
    expect(
      first.assets.every(
        (item: { packageOffset?: number; bytes: number }) =>
          Number.isSafeInteger(item.packageOffset) &&
          Number(item.packageOffset) >= first.package.headerBytes &&
          Number(item.packageOffset) + item.bytes <= first.package.bytes,
      ),
    ).toBe(true);
    expect(JSON.parse(fs.readFileSync(path.join(output, 'channels/stable.json'), 'utf8'))).toEqual({
      version: 1,
      releaseId: first.releaseId,
      manifestUrl: `/releases/${first.releaseId}/manifest-v4.json`,
      manifestSha256: expect.stringMatching(/^[a-f0-9]{64}$/),
    });
    expect(JSON.parse(fs.readFileSync(path.join(output, 'channels/stable-v5.json'), 'utf8'))).toEqual({
      version: 1,
      releaseId: first.releaseId,
      manifestUrl: `/releases/${first.releaseId}/manifest.json`,
      manifestSha256: expect.stringMatching(/^[a-f0-9]{64}$/),
    });
    for (const item of first.assets as Array<{ sha256: string }>) {
      expect(fs.existsSync(path.join(output, 'blobs/sha256', item.sha256))).toBe(true);
    }
    expect(
      first.assets.every(
        (item: { bytes: number; sha256: string; representations: { whole: { bytes: number; sha256: string } } }) =>
          item.representations.whole.bytes === item.bytes && item.representations.whole.sha256 === item.sha256,
      ),
    ).toBe(true);

    fs.writeFileSync(path.join(root, 'sdkjs/word/word.js'), 'word changed');
    const changed = buildRelease({
      root,
      output: temp(),
      packageVersion: '0.4.1',
      sourceCommit: 'b'.repeat(40),
      x2tVersion: '9.3.0+1',
      x2tCommit: 'abc123',
    });
    const originalSegments = new Set(
      (first.package.segments as Array<{ sha256: string; bytes: number }>).slice(1).map((segment) => segment.sha256),
    );
    const reusedSegments = (changed.package.segments as Array<{ sha256: string; bytes: number }>).filter((segment) =>
      originalSegments.has(segment.sha256),
    );
    expect(reusedSegments.length).toBeGreaterThan(0);
    expect(reusedSegments.reduce((sum, segment: { bytes: number }) => sum + segment.bytes, 0)).toBeGreaterThan(0);
  });

  it('adds stable per-file FastCDC objects automatically for large cache writes', () => {
    const root = temp();
    const output = temp();
    const largeChunks = Array.from({ length: 16 }, (_, index) => Buffer.alloc(1024 * 1024, index));
    const largeFile = Buffer.concat(largeChunks);
    const files = {
      'office-host.html': '<html>host</html>',
      'index.html': '<html>shell</html>',
      'sw.js': 'worker',
      'assets/officeHost-test.js': 'host bundle',
      'assets/main-test.js': 'shell bundle',
      'wasm/x2t/x2t.wasm': largeFile,
    };
    for (const [relative, contents] of Object.entries(files)) {
      fs.mkdirSync(path.dirname(path.join(root, relative)), { recursive: true });
      fs.writeFileSync(path.join(root, relative), contents);
    }
    fs.writeFileSync(
      path.join(root, 'onlyoffice-runtime-assets.json'),
      JSON.stringify({
        version: 2,
        assets: [{ path: 'wasm/x2t/x2t.wasm', pack: 'core', bytes: largeFile.byteLength, revision: 'test' }],
      }),
    );
    let offset = 0;
    const chunks = largeChunks.map((bytes) => {
      const chunk = {
        offset,
        bytes: bytes.byteLength,
        sha256: crypto.createHash('sha256').update(bytes).digest('hex'),
      };
      offset += bytes.byteLength;
      return chunk;
    });
    const manifest = buildRelease({
      root,
      output,
      packageVersion: '0.6.0',
      sourceCommit: 'c'.repeat(40),
      x2tVersion: '9.3.0+2',
      x2tCommit: 'abc123',
      fastCdcIndexer: () => ({
        bytes: largeFile.byteLength,
        configurations: [
          {
            minimumBytes: 64 * 1024,
            averageBytes: 256 * 1024,
            maximumBytes: 1024 * 1024,
            chunks,
          },
        ],
      }),
    });
    const x2t = manifest.assets.find((asset: { path: string }) => asset.path === 'wasm/x2t/x2t.wasm');
    expect(x2t?.representations.fastcdc).toMatchObject({
      algorithm: 'fastcdc-v2020',
      chunks,
    });
    for (const chunk of chunks) {
      expect(fs.statSync(path.join(output, 'blobs', 'sha256', chunk.sha256)).size).toBe(chunk.bytes);
    }
    expect(() => parseReleaseManifest(manifest)).not.toThrow();
  });
});
