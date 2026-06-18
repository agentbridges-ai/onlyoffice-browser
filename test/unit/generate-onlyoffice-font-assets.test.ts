import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

interface GeneratorOptions {
  input: string;
  output: string;
  image: string;
  fontSet: string;
  keepFonts: string[];
  help: boolean;
}

interface GeneratorModule {
  DEFAULT_FONT_GENERATOR_IMAGE: string;
  FONT_GENERATOR_IMAGE_ENV: string;
  DEFAULT_FONT_SET: string;
  FONT_SET_ENV: string;
  collectFontFiles(inputDir: string): string[];
  isSupportedFontFileName(fileName: string): boolean;
  parseGenerateFontAssetsArgs(argv: string[], env?: Record<string, string>): GeneratorOptions;
  validateGenerateFontAssetsOptions(options: GeneratorOptions): {
    input: string;
    output: string;
    image: string;
    fontSet: string;
    keepFonts: string[];
    fontFiles: string[];
  };
}

const generator = (await import(
  pathToFileURL(path.resolve('scripts/generate-onlyoffice-font-assets.mjs')).href
)) as GeneratorModule;
const {
  DEFAULT_FONT_GENERATOR_IMAGE,
  DEFAULT_FONT_SET,
  FONT_GENERATOR_IMAGE_ENV,
  FONT_SET_ENV,
  collectFontFiles,
  isSupportedFontFileName,
  parseGenerateFontAssetsArgs,
  validateGenerateFontAssetsOptions,
} = generator;

const tempDirs: string[] = [];

function makeTempDir(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-font-assets-test-'));
  tempDirs.push(dir);
  return dir;
}

afterEach(() => {
  for (const dir of tempDirs.splice(0)) {
    fs.rmSync(dir, { recursive: true, force: true });
  }
});

describe('generate-onlyoffice-font-assets options', () => {
  it('parses required input and output with default image', () => {
    const options = parseGenerateFontAssetsArgs(['--input', '/fonts', '--output', '/assets'], {});

    expect(options).toEqual({
      input: '/fonts',
      output: '/assets',
      image: DEFAULT_FONT_GENERATOR_IMAGE,
      fontSet: DEFAULT_FONT_SET,
      keepFonts: [],
      help: false,
    });
  });

  it('allows image override through env or cli', () => {
    expect(
      parseGenerateFontAssetsArgs(['--input', '/fonts', '--output', '/assets'], {
        [FONT_GENERATOR_IMAGE_ENV]: 'onlyoffice/documentserver:custom',
      }).image,
    ).toBe('onlyoffice/documentserver:custom');

    expect(
      parseGenerateFontAssetsArgs(['--input', '/fonts', '--output', '/assets', '--image', 'custom:tag']).image,
    ).toBe('custom:tag');
  });

  it('allows font set override and extra kept font families', () => {
    expect(
      parseGenerateFontAssetsArgs(['--input', '/fonts', '--output', '/assets'], {
        [FONT_SET_ENV]: 'full',
      }).fontSet,
    ).toBe('full');

    expect(
      parseGenerateFontAssetsArgs([
        '--input',
        '/fonts',
        '--output',
        '/assets',
        '--font-set',
        'zh-core',
        '--keep-font',
        'Wingdings',
        '--keep-font',
        'Symbol',
      ]),
    ).toMatchObject({
      fontSet: 'zh-core',
      keepFonts: ['Wingdings', 'Symbol'],
    });
  });

  it('rejects invalid font set names', () => {
    const input = makeTempDir();
    const output = makeTempDir();
    fs.writeFileSync(path.join(input, 'font.ttf'), '');

    expect(() =>
      validateGenerateFontAssetsOptions(
        parseGenerateFontAssetsArgs(['--input', input, '--output', output, '--font-set', 'everything']),
      ),
    ).toThrow('Invalid --font-set');
  });

  it('requires input and output', () => {
    expect(() => validateGenerateFontAssetsOptions(parseGenerateFontAssetsArgs(['--output', '/assets']))).toThrow(
      'Missing required --input',
    );
    expect(() => validateGenerateFontAssetsOptions(parseGenerateFontAssetsArgs(['--input', '/fonts']))).toThrow(
      'Missing required --output',
    );
  });

  it('rejects an empty input font directory', () => {
    const input = makeTempDir();
    const output = makeTempDir();

    expect(() =>
      validateGenerateFontAssetsOptions(parseGenerateFontAssetsArgs(['--input', input, '--output', output])),
    ).toThrow('No supported font files found');
  });

  it('accepts supported office font extensions recursively', () => {
    const input = makeTempDir();
    const output = makeTempDir();
    fs.mkdirSync(path.join(input, 'nested'));
    fs.writeFileSync(path.join(input, 'nested', 'msyh.ttc'), '');
    fs.writeFileSync(path.join(input, 'notes.txt'), '');

    const validated = validateGenerateFontAssetsOptions(
      parseGenerateFontAssetsArgs(['--input', input, '--output', output]),
    );

    expect(validated.fontFiles).toEqual([path.join(input, 'nested', 'msyh.ttc')]);
    expect(collectFontFiles(input)).toEqual([path.join(input, 'nested', 'msyh.ttc')]);
    expect(isSupportedFontFileName('cambria.otc')).toBe(true);
    expect(isSupportedFontFileName('embedded.tte')).toBe(true);
    expect(isSupportedFontFileName('notes.txt')).toBe(false);
  });
});
