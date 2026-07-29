import { describe, expect, it } from 'vitest';
import {
  buildAllFontsMetadataFallbackBootstrap,
  resolveFontMetadataFallbackConfig,
} from '../../src/lib/font-metadata-fallback';

const manifest = {
  allFonts: 'sdkjs/common/AllFonts.js',
  defaultFonts: ['fonts/019.ttf', 'fonts/020.ttf'],
  builtInFonts: ['fonts/010.ttf'],
  fontFamilies: [
    { name: 'DengXian', paths: ['fonts/019.ttf', 'fonts/020.ttf'] },
    { name: 'Microsoft YaHei', paths: ['fonts/005.ttc', 'fonts/026.ttc'] },
    { name: 'SimHei', paths: ['fonts/037.ttf'] },
  ],
};

describe('font metadata fallback', () => {
  it('classifies complete downloaded families while retaining required defaults', () => {
    expect(resolveFontMetadataFallbackConfig(manifest, ['fonts/005.ttc', 'fonts/026.ttc'])).toEqual({
      fallbackFamilyName: 'DengXian',
      unavailableFamilyNames: ['SimHei'],
      visibleFamilyNames: ['DengXian', 'Microsoft YaHei'],
    });
  });

  it('keeps a partially downloaded family unavailable', () => {
    expect(resolveFontMetadataFallbackConfig(manifest, ['fonts/005.ttc']).unavailableFamilyNames).toContain(
      'Microsoft YaHei',
    );
  });

  it('remaps AllFonts tuples before OnlyOffice constructs its font indexes', () => {
    const config = resolveFontMetadataFallbackConfig(manifest, []);
    const scope: Record<string, unknown> = {
      __fonts_infos: [
        ['DengXian', 19, 0, 19, 0, 20, 0, 20, 0],
        ['Microsoft YaHei', 5, 0, 5, 0, 26, 0, 26, 0],
        ['SimHei', 37, 0, -1, -1, -1, -1, -1, -1],
      ],
      __fonts_visible_names: ['DengXian', 'Microsoft YaHei', 'SimHei'],
    };
    const run = new Function('window', buildAllFontsMetadataFallbackBootstrap(config));

    run(scope);

    expect(scope.__fonts_infos).toEqual([
      ['DengXian', 19, 0, 19, 0, 20, 0, 20, 0],
      ['Microsoft YaHei', 19, 0, 19, 0, 20, 0, 20, 0],
      ['SimHei', 19, 0, 19, 0, 20, 0, 20, 0],
    ]);
    expect(scope.__fonts_visible_names).toEqual(['DengXian']);
    expect(scope.__onlyOfficeBrowserFontMetadataFallback).toEqual(config);
  });
});
