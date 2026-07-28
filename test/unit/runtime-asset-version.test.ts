import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';

type VersionModule = {
  calculateRuntimeAssetVersion(root: string, fontManifestPath?: string): string;
};

const modulePromise = import(
  pathToFileURL(path.resolve('scripts/runtime-asset-version.mjs')).href
) as Promise<VersionModule>;
const roots: string[] = [];

function runtimeRoot(generatedAt: string): string {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-runtime-version-'));
  roots.push(root);
  fs.mkdirSync(path.join(root, 'sdkjs/common'), { recursive: true });
  fs.writeFileSync(path.join(root, 'sdkjs/common/core.js'), 'same-content');
  fs.writeFileSync(
    path.join(root, 'onlyoffice-runtime-assets.json'),
    JSON.stringify({
      version: 2,
      generatedAt,
      selected: 1,
      assets: [{ path: 'sdkjs/common/core.js', pack: 'core', bytes: 12 }],
    }),
  );
  return root;
}

afterEach(() => {
  for (const root of roots.splice(0)) fs.rmSync(root, { recursive: true, force: true });
});

describe('runtime asset version', () => {
  it('stays stable when only generated metadata changes', async () => {
    const mod = await modulePromise;
    expect(mod.calculateRuntimeAssetVersion(runtimeRoot('first'))).toBe(
      mod.calculateRuntimeAssetVersion(runtimeRoot('second')),
    );
  });

  it('changes when a shared asset changes without relying on its size', async () => {
    const mod = await modulePromise;
    const first = runtimeRoot('same');
    const second = runtimeRoot('same');
    fs.writeFileSync(path.join(second, 'sdkjs/common/core.js'), 'same-contEnt');
    expect(mod.calculateRuntimeAssetVersion(first)).not.toBe(mod.calculateRuntimeAssetVersion(second));
  });

  it('ignores origin-bound service worker revisions', async () => {
    const mod = await modulePromise;
    const first = runtimeRoot('same');
    const second = runtimeRoot('same');
    for (const [root, content] of [
      [first, 'first-worker'],
      [second, 'other-worker'],
    ] as const) {
      fs.writeFileSync(path.join(root, 'sw.js'), content);
      const manifestPath = path.join(root, 'onlyoffice-runtime-assets.json');
      const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
      manifest.selected += 1;
      manifest.assets.push({ path: 'sw.js', pack: 'core', bytes: content.length });
      fs.writeFileSync(manifestPath, JSON.stringify(manifest));
    }
    expect(mod.calculateRuntimeAssetVersion(first)).toBe(mod.calculateRuntimeAssetVersion(second));
  });

  it('changes when the generated font overlay manifest changes', async () => {
    const mod = await modulePromise;
    const root = runtimeRoot('same');
    const first = path.join(root, 'fonts-first.json');
    const second = path.join(root, 'fonts-second.json');
    fs.writeFileSync(first, JSON.stringify({ version: 1, assets: [{ path: 'fonts/000.ttf', bytes: 10 }] }));
    fs.writeFileSync(second, JSON.stringify({ version: 1, assets: [{ path: 'fonts/000.ttf', bytes: 11 }] }));
    expect(mod.calculateRuntimeAssetVersion(root, first)).not.toBe(mod.calculateRuntimeAssetVersion(root, second));
  });
});
