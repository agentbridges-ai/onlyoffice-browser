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
    expect(JSON.parse(fs.readFileSync(path.join(output, 'channels/stable.json'), 'utf8'))).toEqual({
      version: 1,
      releaseId: first.releaseId,
    });
    for (const item of first.assets as Array<{ sha256: string }>) {
      expect(fs.existsSync(path.join(output, 'blobs/sha256', item.sha256))).toBe(true);
    }
  });
});
