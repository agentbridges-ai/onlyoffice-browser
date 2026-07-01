import fs from 'node:fs';
import path from 'node:path';
import { describe, expect, it } from 'vitest';

describe('npm package manifest', () => {
  it('does not publish generated or raw font files', () => {
    const pkg = JSON.parse(fs.readFileSync(path.resolve('package.json'), 'utf8')) as { files?: string[] };

    expect(pkg.files).toEqual(expect.any(Array));
    expect(pkg.files).toContain('scripts/generate-onlyoffice-font-assets.mjs');
    expect(pkg.files).toContain('docs/fonts.md');
    expect(pkg.files).toContain('docs/fonts.zh.md');

    for (const entry of pkg.files ?? []) {
      expect(entry).not.toMatch(/(^|\/)fonts\/?$/);
      expect(entry).not.toMatch(/^\.onlyoffice-font-assets\/?$/);
      expect(entry).not.toMatch(/^\.onlyoffice-runtime-assets\/?$/);
      expect(entry).not.toMatch(/\.(?:ttf|ttc|otf|otc|woff2?|tte)$/i);
    }
  });
});
