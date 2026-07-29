import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

type RuntimeAssetModule = {
  parseRuntimeAssetArgs(argv: string[]): {
    input: string;
    output: string;
    splitOutput: string;
    pruneRoot: boolean;
    types: string[];
    dictionaries: string[];
    keepHelp: boolean;
    dryRun: boolean;
  };
  validateRuntimeAssetOptions(options: any): any;
  getRuntimeAssetPack(
    relativePath: string,
    options?: { types?: string[]; dictionaries?: string[]; keepHelp?: boolean },
  ): string | null;
  buildRuntimeAssets(options: any): {
    version: number;
    selected: number;
    excluded: number;
    totalBytes: number;
    packs: Record<string, number>;
    assets: Array<{ path: string; pack: string; bytes: number; revision: string }>;
  };
};

const modulePromise = import(
  pathToFileURL(path.resolve('scripts/build-onlyoffice-runtime-assets.mjs')).href
) as Promise<RuntimeAssetModule>;

const tempRoots: string[] = [];

function makeTempRoot() {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-runtime-assets-test-'));
  tempRoots.push(root);
  return root;
}

function touch(root: string, relativePath: string) {
  const target = path.join(root, relativePath);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  fs.writeFileSync(target, relativePath);
}

function exists(root: string, relativePath: string) {
  return fs.existsSync(path.join(root, relativePath));
}

afterEach(() => {
  for (const root of tempRoots.splice(0)) {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

describe('build-onlyoffice-runtime-assets', () => {
  it('ships the upstream-generated high-density document format sprite', () => {
    const spritePath = path.resolve('public/web-apps/apps/common/main/resources/img/doc-formats/formats@2.5x.svg');
    const sprite = fs.readFileSync(spritePath, 'utf8');

    expect(sprite).toContain('<symbol');
    expect(sprite).toContain('id="docx"');
    expect(sprite).toContain('id="xlsx"');
    expect(sprite).toContain('id="pptx"');
  });

  it('classifies core and document-type runtime assets', async () => {
    const mod = await modulePromise;

    expect(mod.getRuntimeAssetPack('web-apps/apps/api/documents/api.js')).toBe('core');
    expect(mod.getRuntimeAssetPack('web-apps/vendor/fetch/fetch.umd.js')).toBe('core');
    expect(mod.getRuntimeAssetPack('sdkjs/common/zlib/engine/zlib.wasm')).toBe('core');
    expect(mod.getRuntimeAssetPack('wasm/x2t/x2t.wasm')).toBe('core');
    expect(mod.getRuntimeAssetPack('dictionaries/en_US/en_US.dic')).toBe('core');
    expect(mod.getRuntimeAssetPack('web-apps/apps/documenteditor/main/app.js')).toBe('word');
    expect(mod.getRuntimeAssetPack('sdkjs/word/sdk-all.js')).toBe('word');
    expect(mod.getRuntimeAssetPack('web-apps/apps/spreadsheeteditor/embed/app-all.js')).toBe('cell');
    expect(mod.getRuntimeAssetPack('sdkjs/cell/sdk-all.js')).toBe('cell');
    expect(mod.getRuntimeAssetPack('web-apps/apps/presentationeditor/main/app.js')).toBe('slide');
    expect(mod.getRuntimeAssetPack('sdkjs/slide/sdk-all.js')).toBe('slide');
    expect(mod.getRuntimeAssetPack('sw.js')).toBeNull();
  });

  it('excludes low-frequency assets from the compact default profile', async () => {
    const mod = await modulePromise;

    expect(mod.getRuntimeAssetPack('sdkjs/pdf/src/engine/drawingfile.wasm')).toBeNull();
    expect(mod.getRuntimeAssetPack('sdkjs/visio/sdk-all.js')).toBeNull();
    expect(mod.getRuntimeAssetPack('dictionaries/fr_FR/fr_FR.dic')).toBeNull();
    expect(mod.getRuntimeAssetPack('fonts/000.ttf')).toBeNull();
    expect(mod.getRuntimeAssetPack('server/FileConverter/bin/font_selection.bin')).toBeNull();
    expect(
      mod.getRuntimeAssetPack('web-apps/apps/spreadsheeteditor/main/resources/help/en/images/lookup_function.gif'),
    ).toBeNull();
  });

  it('allows selected dictionaries and help resources when explicitly requested', async () => {
    const mod = await modulePromise;

    expect(mod.getRuntimeAssetPack('dictionaries/fr_FR/fr_FR.dic', { dictionaries: ['en_US', 'fr_FR'] })).toBe('core');
    expect(
      mod.getRuntimeAssetPack('web-apps/apps/documenteditor/main/resources/help/en/images/example.gif', {
        keepHelp: true,
      }),
    ).toBe('word');
  });

  it('validates options and rejects unknown document types', async () => {
    const mod = await modulePromise;
    const input = makeTempRoot();
    touch(input, 'web-apps/apps/api/documents/api.js');
    const parsed = mod.parseRuntimeAssetArgs([
      '--input',
      input,
      '--output',
      path.join(input, 'out'),
      '--types',
      'word',
    ]);

    expect(mod.validateRuntimeAssetOptions(parsed).types).toEqual(['word']);
    expect(() =>
      mod.validateRuntimeAssetOptions(
        mod.parseRuntimeAssetArgs(['--input', input, '--output', path.join(input, 'out'), '--types', 'pdf']),
      ),
    ).toThrow('Invalid --types value');
  });

  it('prunes root assets and creates canonical split packs', async () => {
    const mod = await modulePromise;
    const input = makeTempRoot();
    touch(input, 'web-apps/apps/api/documents/api.js');
    touch(input, 'web-apps/apps/documenteditor/main/app.js');
    touch(input, 'web-apps/apps/spreadsheeteditor/main/app.js');
    touch(input, 'sw.js');
    const emptySprite = 'web-apps/apps/spreadsheeteditor/main/resources/img/iconshuge.png';
    fs.mkdirSync(path.dirname(path.join(input, emptySprite)), { recursive: true });
    fs.writeFileSync(path.join(input, emptySprite), '');
    touch(input, 'web-apps/apps/spreadsheeteditor/main/resources/help/en/images/large.gif');
    touch(input, 'sdkjs/word/sdk-all.js');
    touch(input, 'sdkjs/cell/sdk-all.js');
    touch(input, 'sdkjs/pdf/src/engine/drawingfile.wasm');
    touch(input, 'dictionaries/en_US/en_US.dic');
    touch(input, 'dictionaries/fr_FR/fr_FR.dic');

    const splitOutput = path.join(input, 'asset-packs');
    const manifest = mod.buildRuntimeAssets(
      mod.validateRuntimeAssetOptions(
        mod.parseRuntimeAssetArgs([
          '--input',
          input,
          '--prune-root',
          '--split-output',
          splitOutput,
          '--types',
          'word,cell',
        ]),
      ),
    );

    expect(manifest.packs.core).toBe(2);
    expect(manifest.packs.word).toBe(2);
    expect(manifest.packs.cell).toBe(2);
    expect(manifest.version).toBe(2);
    expect(manifest.totalBytes).toBe(manifest.assets.reduce((total, asset) => total + asset.bytes, 0));
    expect(manifest.assets.every((asset) => asset.bytes > 0)).toBe(true);
    expect(manifest.assets.some((asset) => asset.path === emptySprite)).toBe(false);
    expect(manifest.assets.every((asset) => /^[a-f0-9]{16}$/.test(asset.revision))).toBe(true);
    expect(manifest.assets).toEqual([...manifest.assets].sort((left, right) => left.path.localeCompare(right.path)));
    expect(exists(input, 'web-apps/apps/api/documents/api.js')).toBe(true);
    expect(exists(input, 'sw.js')).toBe(true);
    expect(manifest.assets.some((asset) => asset.path === 'sw.js')).toBe(false);
    expect(exists(input, 'sdkjs/pdf/src/engine/drawingfile.wasm')).toBe(false);
    expect(exists(input, 'dictionaries/fr_FR/fr_FR.dic')).toBe(false);
    expect(exists(input, emptySprite)).toBe(false);
    expect(exists(input, 'web-apps/apps/spreadsheeteditor/main/resources/help/en/images/large.gif')).toBe(false);
    expect(exists(splitOutput, 'core/web-apps/apps/api/documents/api.js')).toBe(true);
    expect(exists(splitOutput, 'word/sdkjs/word/sdk-all.js')).toBe(true);
    expect(exists(splitOutput, 'cell/sdkjs/cell/sdk-all.js')).toBe(true);
  });

  it('hashes the runtime only after all build-time content patches', () => {
    const buildScript = fs.readFileSync(path.resolve('bin/build.sh'), 'utf8');
    const patchIndex = buildScript.indexOf('node scripts/patch-onlyoffice-print-fallback.mjs dist');
    const revisionIndex = buildScript.indexOf('node scripts/inject-build-revision.mjs dist/sw.js');
    const manifestIndex = buildScript.indexOf('node scripts/build-onlyoffice-runtime-assets.mjs');
    const releaseIndex = buildScript.indexOf('node scripts/build-release-manifest.mjs');

    expect(patchIndex).toBeGreaterThanOrEqual(0);
    expect(revisionIndex).toBeGreaterThan(patchIndex);
    expect(manifestIndex).toBeGreaterThan(revisionIndex);
    expect(releaseIndex).toBeGreaterThan(manifestIndex);
  });
});
