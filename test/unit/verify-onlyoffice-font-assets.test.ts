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

function writeFixture(root: string): void {
  fs.mkdirSync(path.join(root, 'sdkjs/common/Images'), { recursive: true });
  fs.mkdirSync(path.join(root, 'server/FileConverter/bin'), { recursive: true });
  fs.mkdirSync(path.join(root, 'fonts'), { recursive: true });
  fs.writeFileSync(path.join(root, 'sdkjs/common/AllFonts.js'), '');
  fs.writeFileSync(path.join(root, 'sdkjs/common/Images/fonts_thumbnail.png'), '');
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
});
