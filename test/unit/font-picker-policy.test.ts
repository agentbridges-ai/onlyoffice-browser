import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

interface PolicyModule {
  applyFontPickerPolicy(rootDirectory: string): { visibleNames: string[] };
}

const { applyFontPickerPolicy } = (await import(
  pathToFileURL(path.resolve('scripts/apply-font-picker-policy.mjs')).href
)) as PolicyModule;

const tempDirectories: string[] = [];

function fixture(): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-font-picker-policy-'));
  tempDirectories.push(root);
  fs.mkdirSync(path.join(root, 'sdkjs/common'), { recursive: true });
  const infos = ['Aptos', 'Arial', 'DengXian', 'Vemana2000', 'Wingdings'].map((name) => [
    name,
    0,
    0,
    -1,
    -1,
    -1,
    -1,
    -1,
    -1,
  ]);
  fs.writeFileSync(
    path.join(root, 'sdkjs/common/AllFonts.js'),
    [
      'window["__fonts_files"] = ["000.ttf"];',
      `window["__fonts_infos"] = ${JSON.stringify(infos)};`,
      'window["__fonts_ranges"] = [32,126,0];',
      `window["__fonts_visible_names"] = ${JSON.stringify(infos.map((info) => info[0]))};`,
    ].join('\n'),
  );
  fs.writeFileSync(
    path.join(root, 'onlyoffice-browser-font-source-map.json'),
    JSON.stringify({
      keptFamilies: infos.map((info) => info[0]),
      visibleFamilies: infos.map((info) => info[0]),
    }),
  );
  fs.writeFileSync(
    path.join(root, 'onlyoffice-browser-font-assets.json'),
    JSON.stringify({
      version: 1,
      fontSet: 'full',
      fontFamilies: infos.map((info) => ({ name: info[0], paths: ['fonts/000.ttf'] })),
      totalBytes: 2,
      assets: [
        { path: 'sdkjs/common/AllFonts.js', bytes: 1, revision: 'old' },
        { path: 'onlyoffice-browser-font-source-map.json', bytes: 1, revision: 'old' },
      ],
    }),
  );
  return root;
}

afterEach(() => {
  for (const root of tempDirectories.splice(0)) {
    fs.rmSync(root, { recursive: true, force: true });
  }
});

describe('font picker policy', () => {
  it('hides support faces while keeping their runtime mappings installed', () => {
    const root = fixture();
    const first = applyFontPickerPolicy(root);
    const second = applyFontPickerPolicy(root);

    expect(first.visibleNames).toEqual(['Aptos', 'Arial', 'DengXian']);
    expect(second).toEqual(first);
    const allFonts = fs.readFileSync(path.join(root, 'sdkjs/common/AllFonts.js'), 'utf8');
    expect(allFonts).toContain('"Vemana2000"');
    expect(allFonts).toContain('"Wingdings"');
    expect(allFonts).toContain('window["__fonts_visible_names"] = ["Aptos","Arial","DengXian"];');
    const sourceMap = JSON.parse(fs.readFileSync(path.join(root, 'onlyoffice-browser-font-source-map.json'), 'utf8'));
    expect(sourceMap.visibleFamilies).toEqual(first.visibleNames);
    expect(sourceMap.keptFamilies).toContain('Wingdings');
    const manifest = JSON.parse(fs.readFileSync(path.join(root, 'onlyoffice-browser-font-assets.json'), 'utf8'));
    expect(manifest.fontFamilies.map((family: { name: string }) => family.name)).toEqual(first.visibleNames);
    expect(manifest.assets).not.toContainEqual(expect.objectContaining({ revision: 'old' }));
    expect(manifest.totalBytes).toBe(
      manifest.assets.reduce((total: number, asset: { bytes: number }) => total + asset.bytes, 0),
    );
  });
});
