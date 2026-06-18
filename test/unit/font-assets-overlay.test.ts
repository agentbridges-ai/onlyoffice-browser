import path from 'node:path';
import { describe, expect, it } from 'vitest';
import { GENERATED_FONT_ASSETS_MANIFEST, resolveGeneratedFontAssetPath } from '../../vite.config';

const ROOT = path.resolve('/tmp/generated-font-assets');

describe('generated font assets overlay', () => {
  it.each([
    ['/sdkjs/common/AllFonts.js', 'sdkjs/common/AllFonts.js'],
    ['/sdkjs/common/Images/fonts_thumbnail.png', 'sdkjs/common/Images/fonts_thumbnail.png'],
    ['/sdkjs/common/Images/fonts_thumbnail@2x.png', 'sdkjs/common/Images/fonts_thumbnail@2x.png'],
    ['/sdkjs/common/Images/fonts_thumbnail_ea@1.5x.png', 'sdkjs/common/Images/fonts_thumbnail_ea@1.5x.png'],
    ['/fonts/calibri.ttf', 'fonts/calibri.ttf'],
    ['/fonts/000', 'fonts/000'],
    ['/fonts/cjk/msyh.ttc', 'fonts/cjk/msyh.ttc'],
    ['/server/FileConverter/bin/font_selection.bin', 'server/FileConverter/bin/font_selection.bin'],
    [`/${GENERATED_FONT_ASSETS_MANIFEST}`, GENERATED_FONT_ASSETS_MANIFEST],
  ])('maps %s to generated asset %s', (requestUrl, relativePath) => {
    expect(resolveGeneratedFontAssetPath(ROOT, requestUrl)).toBe(path.join(ROOT, relativePath));
  });

  it.each([
    '/sdkjs/common/Images/logo.png',
    '/sdkjs/common/Images/fonts_thumbnail.svg',
    '/web-apps/apps/api/documents/api.js',
    '/wasm/x2t/x2t.wasm',
    '/server/FileConverter/bin/other.bin',
    '/fonts/fonts/000',
    '/fonts//fonts/000',
  ])('does not overlay unrelated runtime asset %s', (requestUrl) => {
    expect(resolveGeneratedFontAssetPath(ROOT, requestUrl)).toBeNull();
  });
});
