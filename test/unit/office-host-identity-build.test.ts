import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import packageJson from '../../package.json';

describe('Office Host build identity', () => {
  it('derives package and Host versions from npm build metadata', () => {
    const source = readFileSync(resolve(process.cwd(), 'src/office-host.ts'), 'utf8');
    const viteConfig = readFileSync(resolve(process.cwd(), 'vite.config.ts'), 'utf8');
    expect(source).toContain("import { ONLYOFFICE_BROWSER_VERSION } from './version'");
    expect(source).toContain('const OFFICE_BROWSER_PACKAGE_VERSION = ONLYOFFICE_BROWSER_VERSION');
    expect(source).toContain('`office-host-${ONLYOFFICE_BROWSER_VERSION}-r1`');
    expect(source).not.toMatch(/OFFICE_BROWSER_PACKAGE_VERSION\s*=\s*['"]0\./);
    expect(viteConfig).toContain('entryFileNames: `assets/[name]-v${packageJson.version}-[hash].js`');
    expect(viteConfig).toContain('chunkFileNames: `assets/[name]-v${packageJson.version}-[hash].js`');
    expect(viteConfig).toContain('assetFileNames: `assets/[name]-v${packageJson.version}-[hash][extname]`');
    expect(packageJson.version).toBe('0.5.14');
  });

  it('returns saved bytes to the embedding parent instead of downloading inside the isolated iframe', () => {
    const source = readFileSync(resolve(process.cwd(), 'src/office-host.ts'), 'utf8');
    expect(source).toContain("saveBehavior: 'callback'");
    expect(source).not.toContain('saveBehavior: message.options.saveBehavior');
    expect(source).toContain('onSave: (file) => postSavedFile(file, activeSaveRequestId)');
  });

  it('keeps the fixed origin service worker and verified package cache across editor reuse', () => {
    const source = readFileSync(resolve(process.cwd(), 'src/office-host.ts'), 'utf8');
    expect(source).toContain('isReusableOfficeEditorHostname(window.location.hostname)');
    expect(source).toContain('window.caches?.delete(PRINT_PDF_CACHE_NAME)');
    expect(source).toContain('registrations.map((registration) => registration.unregister())');
  });
});
