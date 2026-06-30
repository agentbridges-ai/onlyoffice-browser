import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

interface VerifyModule {
  FONT_ASSETS_DIR_ENV: string;
  parseVerifyFontAssetsArgs(argv: string[], env?: Record<string, string>): { input: string; help: boolean };
  verifyOnlyOfficeFontAssets(input: string): { root: string; fontSet: string; fonts: number; thumbnails: number };
}

const verifier = (await import(
  pathToFileURL(path.resolve('scripts/verify-onlyoffice-font-assets.mjs')).href
)) as VerifyModule;
const { FONT_ASSETS_DIR_ENV, parseVerifyFontAssetsArgs, verifyOnlyOfficeFontAssets } = verifier;

const tempDirs: string[] = [];

function makeTempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-font-assets-verify-test-'));
  tempDirs.push(dir);
  return dir;
}

function makePngHeader(width: number, height: number): Buffer {
  const buffer = Buffer.alloc(24);
  Buffer.from('89504e470d0a1a0a', 'hex').copy(buffer, 0);
  buffer.writeUInt32BE(13, 8);
  buffer.write('IHDR', 12, 'ascii');
  buffer.writeUInt32BE(width, 16);
  buffer.writeUInt32BE(height, 20);
  return buffer;
}

function writeAllFonts(
  root: string,
  options: {
    files?: string[];
    infos?: unknown[];
    ranges?: number[];
    visibleNames?: string[];
  } = {},
): void {
  const files = options.files ?? ['000.ttf'];
  const infos = options.infos ?? [
    ['Arial', 0, 0, -1, -1, -1, -1, -1, -1],
    ['Droid Sans Fallback', 0, 0, -1, -1, -1, -1, -1, -1],
  ];
  const ranges = options.ranges ?? [32, 126, 0, 19968, 40869, 1];
  const visibleNames = options.visibleNames ?? ['Arial', 'Droid Sans Fallback'];
  fs.writeFileSync(
    path.join(root, 'sdkjs/common/AllFonts.js'),
    [
      `window["__fonts_files"] = ${JSON.stringify(files)};`,
      `window["__fonts_infos"] = ${JSON.stringify(infos)};`,
      `window["__fonts_ranges"] = ${JSON.stringify(ranges)};`,
      `window["__fonts_visible_names"] = ${JSON.stringify(visibleNames)};`,
    ].join('\n'),
  );
}

function writeFixture(root: string): void {
  fs.mkdirSync(path.join(root, 'sdkjs/common/Images'), { recursive: true });
  fs.mkdirSync(path.join(root, 'server/FileConverter/bin'), { recursive: true });
  fs.mkdirSync(path.join(root, 'fonts'), { recursive: true });
  writeAllFonts(root);
  fs.writeFileSync(path.join(root, 'sdkjs/common/Images/fonts_thumbnail.png'), makePngHeader(300, 56));
  fs.writeFileSync(path.join(root, 'server/FileConverter/bin/font_selection.bin'), '');
  fs.writeFileSync(path.join(root, 'fonts/000.ttf'), '');
  fs.writeFileSync(
    path.join(root, 'onlyoffice-browser-font-assets.json'),
    JSON.stringify({
      version: 1,
      fontSet: 'zh-core',
      allFonts: 'sdkjs/common/AllFonts.js',
      fontSelection: 'server/FileConverter/bin/font_selection.bin',
      fontThumbnails: ['sdkjs/common/Images/fonts_thumbnail.png'],
      fonts: ['fonts/000.ttf'],
    }),
  );
}

afterEach(() => {
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe('verify-onlyoffice-font-assets', () => {
  it('parses input from cli or env', () => {
    expect(parseVerifyFontAssetsArgs(['--input', '/assets'], {}).input).toBe('/assets');
    expect(parseVerifyFontAssetsArgs([], { [FONT_ASSETS_DIR_ENV]: '/env-assets' }).input).toBe('/env-assets');
  });

  it('verifies a complete generated font asset directory', () => {
    const root = makeTempDir();
    writeFixture(root);

    expect(verifyOnlyOfficeFontAssets(root)).toEqual({
      root,
      fontSet: 'zh-core',
      fonts: 1,
      thumbnails: 1,
    });
  });

  it('rejects missing generated font files referenced by the manifest', () => {
    const root = makeTempDir();
    writeFixture(root);
    fs.rmSync(path.join(root, 'fonts/000.ttf'));

    expect(() => verifyOnlyOfficeFontAssets(root)).toThrow('Generated font asset is missing: fonts/000.ttf');
  });

  it('rejects AllFonts font file references that are missing from the generated fonts directory', () => {
    const root = makeTempDir();
    writeFixture(root);
    writeAllFonts(root, { files: ['001.ttf'] });

    expect(() => verifyOnlyOfficeFontAssets(root)).toThrow(
      'Generated AllFonts.js references missing font file: fonts/001.ttf',
    );
  });

  it('rejects CJK ranges that point at a non-CJK font family', () => {
    const root = makeTempDir();
    writeFixture(root);
    writeAllFonts(root, { ranges: [32, 126, 0, 19968, 40869, 0] });

    expect(() => verifyOnlyOfficeFontAssets(root)).toThrow(
      'Generated AllFonts.js maps CJK range 19968-40869 to non-CJK font Arial',
    );
  });

  it('rejects font thumbnails whose row count no longer matches AllFonts entries', () => {
    const root = makeTempDir();
    writeFixture(root);
    fs.writeFileSync(path.join(root, 'sdkjs/common/Images/fonts_thumbnail.png'), makePngHeader(300, 28));

    expect(() => verifyOnlyOfficeFontAssets(root)).toThrow(
      'Generated font thumbnail row height is too small for __fonts_infos',
    );
  });
});
