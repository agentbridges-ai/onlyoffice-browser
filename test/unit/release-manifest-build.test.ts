import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
// The release builder is intentionally executable Node ESM rather than part of
// the browser TypeScript bundle.
// @ts-expect-error JavaScript build script has no declaration output.
import { buildRelease, chunkReleaseAssets } from '../../scripts/build-release-manifest.mjs';

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
      x2tVersion: '9.3.0+1',
      x2tCommit: 'abc123',
    });
    const secondOutput = temp();
    const second = buildRelease({
      root,
      output: secondOutput,
      packageVersion: '0.4.0',
      x2tVersion: '9.3.0+1',
      x2tCommit: 'abc123',
    });
    expect(second.releaseId).toBe(first.releaseId);
    expect(first.version).toBe(4);
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
    });
    for (const item of first.assets as Array<{ sha256: string }>) {
      expect(fs.existsSync(path.join(output, 'blobs/sha256', item.sha256))).toBe(true);
    }

    fs.writeFileSync(path.join(root, 'sdkjs/word/word.js'), 'word changed');
    const changed = buildRelease({
      root,
      output: temp(),
      packageVersion: '0.4.1',
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
});
