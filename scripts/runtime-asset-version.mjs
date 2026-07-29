#!/usr/bin/env node
import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export function calculateRuntimeAssetVersion(root, fontManifestPath = '') {
  const absoluteRoot = path.resolve(root);
  const manifestPath = path.join(absoluteRoot, 'onlyoffice-runtime-assets.json');
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
  if (!Array.isArray(manifest.assets) || manifest.assets.length !== manifest.selected) {
    throw new Error('Runtime asset manifest is missing its complete asset inventory');
  }

  const digest = crypto.createHash('sha256');
  const sharedAssets = manifest.assets.filter((asset) => {
    const basename = path.posix.basename(asset.path).toLowerCase();
    return (
      asset.path !== 'reset.html' &&
      asset.path !== 'sw.js' &&
      asset.path !== 'document_editor_service_worker.js' &&
      !basename.includes('worker')
    );
  });
  for (const asset of [...sharedAssets].sort((left, right) => left.path.localeCompare(right.path))) {
    const filePath = path.resolve(absoluteRoot, asset.path);
    const relative = path.relative(absoluteRoot, filePath);
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) {
      throw new Error(`Runtime asset path escapes its root: ${asset.path}`);
    }
    const bytes = fs.readFileSync(filePath);
    const revision = crypto.createHash('sha256').update(bytes).digest('hex').slice(0, 16);
    if (asset.revision !== revision) {
      throw new Error(`Runtime asset revision does not match final bytes: ${asset.path}`);
    }
    digest.update(asset.path);
    digest.update('\0');
    digest.update(asset.revision);
    digest.update('\0');
  }
  if (fontManifestPath) {
    digest.update('font-manifest\0');
    digest.update(
      crypto
        .createHash('sha256')
        .update(fs.readFileSync(path.resolve(fontManifestPath)))
        .digest(),
    );
  }
  return digest.digest('hex').slice(0, 16);
}

function main() {
  const rootIndex = process.argv.indexOf('--root');
  const root = rootIndex >= 0 ? process.argv[rootIndex + 1] : '';
  const fontManifestIndex = process.argv.indexOf('--font-manifest');
  const fontManifestPath = fontManifestIndex >= 0 ? process.argv[fontManifestIndex + 1] : '';
  if (!root) {
    throw new Error(
      'Usage: node scripts/runtime-asset-version.mjs --root <runtime-directory> [--font-manifest <path>]',
    );
  }
  process.stdout.write(`${calculateRuntimeAssetVersion(root, fontManifestPath)}\n`);
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  try {
    main();
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
