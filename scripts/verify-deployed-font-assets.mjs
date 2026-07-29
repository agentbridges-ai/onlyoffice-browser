#!/usr/bin/env node

import path from 'node:path';
import { fileURLToPath } from 'node:url';

function isSafeAssetPath(value) {
  return typeof value === 'string' && value.length > 0 && !value.startsWith('/') && !value.includes('..');
}

function parseManifest(value) {
  if (!value || !Array.isArray(value.assets) || value.assets.length === 0) {
    throw new Error('Deployed font manifest is missing its asset inventory');
  }
  for (const asset of value.assets) {
    if (
      !asset ||
      !isSafeAssetPath(asset.path) ||
      !Number.isSafeInteger(asset.bytes) ||
      asset.bytes < 0 ||
      typeof asset.revision !== 'string' ||
      asset.revision.length === 0
    ) {
      throw new Error('Deployed font manifest contains an invalid asset');
    }
  }
  return value.assets;
}

export async function verifyDeployedFontAssets({ manifestUrl, fetchImpl = fetch, concurrency = 8 }) {
  const response = await fetchImpl(manifestUrl, {
    cache: 'no-cache',
    headers: { accept: 'application/json' },
  });
  if (!response.ok) throw new Error(`Deployed font manifest request failed (HTTP ${response.status})`);
  const assets = parseManifest(await response.json());
  const failures = [];
  let cursor = 0;
  const worker = async () => {
    while (cursor < assets.length) {
      const asset = assets[cursor++];
      const url = new URL(asset.path, manifestUrl);
      url.searchParams.set('__oobv', asset.revision);
      try {
        const assetResponse = await fetchImpl(url, {
          headers: { range: 'bytes=0-0' },
          cache: 'no-cache',
        });
        const contentRange = assetResponse.headers.get('content-range');
        const rangeMatch = contentRange?.match(/\/(\d+)$/);
        const deployedBytes = rangeMatch ? Number(rangeMatch[1]) : Number(assetResponse.headers.get('content-length'));
        await assetResponse.body?.cancel();
        if (!assetResponse.ok) {
          failures.push(`${asset.path}: HTTP ${assetResponse.status}`);
        } else if (!Number.isSafeInteger(deployedBytes) || deployedBytes !== asset.bytes) {
          failures.push(`${asset.path}: expected ${asset.bytes} bytes, received ${deployedBytes || 'unknown'}`);
        }
      } catch (error) {
        failures.push(`${asset.path}: ${error instanceof Error ? error.message : String(error)}`);
      }
    }
  };
  await Promise.all(Array.from({ length: Math.min(Math.max(1, concurrency), assets.length) }, worker));
  if (failures.length > 0) {
    throw new Error(`Deployed font assets are incomplete:\n${failures.sort().join('\n')}`);
  }
  return { checked: assets.length };
}

function readOption(argv, name) {
  const index = argv.indexOf(name);
  return index >= 0 ? argv[index + 1] : '';
}

async function main() {
  const manifestUrl = readOption(process.argv.slice(2), '--manifest-url');
  if (!manifestUrl) {
    throw new Error(
      'Usage: node scripts/verify-deployed-font-assets.mjs --manifest-url <https://host/onlyoffice-browser-font-assets.json>',
    );
  }
  const result = await verifyDeployedFontAssets({ manifestUrl });
  process.stdout.write(`Verified ${result.checked} deployed font assets.\n`);
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  });
}
