#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { applyFontPickerPolicy } from './apply-font-picker-policy.mjs';

const sourceOrigin = (process.env.ONLYOFFICE_MATRIX_FONT_SOURCE || 'https://onlyoffice.getpi.work').replace(/\/+$/, '');
const pinnedReleaseId = process.env.ONLYOFFICE_MATRIX_FONT_RELEASE_ID || '';
const pinnedManifestSha256 = process.env.ONLYOFFICE_MATRIX_FONT_MANIFEST_SHA256 || '';
const cacheRoot = path.resolve(process.env.ONLYOFFICE_MATRIX_FONT_CACHE || '.onlyoffice-cloudflare-matrix-cache');
const distRoot = path.resolve(process.env.ONLYOFFICE_MATRIX_DIST || 'dist');
const localFontRoot = process.env.ONLYOFFICE_MATRIX_FONT_ASSETS_DIR
  ? path.resolve(process.env.ONLYOFFICE_MATRIX_FONT_ASSETS_DIR)
  : '';
const fontProfiles = new Set(['fonts-basic', 'fonts-office-compat']);
const DEFAULT_FONT_DOWNLOAD_CONCURRENCY = 3;
const RANGED_FONT_ASSET_THRESHOLD_BYTES = 8 * 1024 * 1024;
const DEFAULT_FONT_RANGE_BYTES = 4 * 1024 * 1024;

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

export async function fetchWithRetry(
  url,
  consume = (response) => response,
  retryDelays = [1_000, 3_000, 10_000, 30_000],
  requestInit = {},
) {
  let lastError;
  for (let attempt = 0; attempt <= retryDelays.length; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 60_000);
    try {
      const response = await fetch(url, { ...requestInit, cache: 'no-store', signal: controller.signal });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return await consume(response);
    } catch (error) {
      lastError = error;
      if (attempt < retryDelays.length) {
        await new Promise((resolve) => setTimeout(resolve, retryDelays[attempt] ?? 0));
      }
    } finally {
      clearTimeout(timeout);
    }
  }
  throw new Error(`Unable to download ${url}: ${lastError}`);
}

function fontDownloadConcurrency() {
  const parsed = Number.parseInt(process.env.ONLYOFFICE_MATRIX_FONT_CONCURRENCY || '', 10);
  if (!Number.isSafeInteger(parsed) || parsed < 1) return DEFAULT_FONT_DOWNLOAD_CONCURRENCY;
  return Math.min(parsed, 6);
}

async function readRelease() {
  let releaseId = pinnedReleaseId;
  if (!releaseId) {
    const channel = await fetchWithRetry(`${sourceOrigin}/channels/stable.json`, (response) => response.json());
    if (channel.version !== 1 || typeof channel.releaseId !== 'string') {
      throw new Error('Cloudflare matrix font source returned an invalid stable channel');
    }
    releaseId = channel.releaseId;
  }
  const manifest = await fetchWithRetry(
    `${sourceOrigin}/releases/${encodeURIComponent(releaseId)}/manifest.json`,
    (response) => response.json(),
  );
  if (
    (manifest.version !== 3 && manifest.version !== 4) ||
    manifest.releaseId !== releaseId ||
    !Array.isArray(manifest.assets)
  ) {
    throw new Error('Cloudflare matrix font source returned an invalid release manifest');
  }
  const assets = manifest.assets.filter(
    (asset) => fontProfiles.has(asset.profile) || asset.path === 'onlyoffice-browser-font-assets.json',
  );
  if (
    !assets.length ||
    !assets.some((asset) => asset.path === 'onlyoffice-browser-font-assets.json') ||
    assets.some(
      (asset) =>
        typeof asset.path !== 'string' ||
        !/^[a-f0-9]{64}$/.test(asset.sha256) ||
        !Number.isSafeInteger(asset.bytes) ||
        asset.bytes < 0,
    )
  ) {
    throw new Error('Cloudflare matrix font release is incomplete');
  }
  for (const asset of assets) safeAssetPath(cacheRoot, asset.path);
  const fontManifest = assets.find((asset) => asset.path === 'onlyoffice-browser-font-assets.json');
  if (pinnedManifestSha256 && fontManifest.sha256 !== pinnedManifestSha256) {
    throw new Error(
      `Pinned Cloudflare font manifest mismatch: expected ${pinnedManifestSha256}, received ${fontManifest.sha256}`,
    );
  }
  return { releaseId, assets };
}

function safeAssetPath(root, assetPath) {
  const resolvedRoot = path.resolve(root);
  const target = path.resolve(resolvedRoot, assetPath);
  const relative = path.relative(resolvedRoot, target);
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) {
    throw new Error(`Font asset path escapes its cache root: ${assetPath}`);
  }
  return target;
}

function validCachedAsset(asset) {
  const target = safeAssetPath(cacheRoot, asset.path);
  if (!fs.existsSync(target) || fs.statSync(target).size !== asset.bytes) return false;
  return sha256(fs.readFileSync(target)) === asset.sha256;
}

function parseContentRange(value, expectedStart, expectedEnd, expectedTotal) {
  const match = /^bytes (\d+)-(\d+)\/(\d+)$/.exec(value || '');
  if (!match) return false;
  return Number(match[1]) === expectedStart && Number(match[2]) === expectedEnd && Number(match[3]) === expectedTotal;
}

export async function downloadRangedAsset(releaseId, asset, options = {}) {
  const root = options.root || cacheRoot;
  const rangeBytes = options.rangeBytes || DEFAULT_FONT_RANGE_BYTES;
  if (!Number.isSafeInteger(rangeBytes) || rangeBytes <= 0) throw new TypeError('Invalid font range size');
  const target = safeAssetPath(root, asset.path);
  const temporary = `${target}.download`;
  const metadataPath = `${temporary}.json`;
  const metadata = {
    releaseId,
    path: asset.path,
    bytes: asset.bytes,
    sha256: asset.sha256,
  };
  let offset = 0;
  let digest = crypto.createHash('sha256');
  try {
    if (fs.existsSync(temporary) && fs.existsSync(metadataPath)) {
      try {
        const saved = JSON.parse(fs.readFileSync(metadataPath, 'utf8'));
        const size = fs.statSync(temporary).size;
        if (
          saved.releaseId === metadata.releaseId &&
          saved.path === metadata.path &&
          saved.bytes === metadata.bytes &&
          saved.sha256 === metadata.sha256 &&
          Number.isSafeInteger(size) &&
          size >= 0 &&
          size <= asset.bytes
        ) {
          const prefix = fs.readFileSync(temporary);
          digest.update(prefix);
          offset = prefix.byteLength;
        } else {
          fs.rmSync(temporary, { force: true });
          fs.rmSync(metadataPath, { force: true });
        }
      } catch {
        fs.rmSync(temporary, { force: true });
        fs.rmSync(metadataPath, { force: true });
      }
    }

    fs.mkdirSync(path.dirname(target), { recursive: true });
    if (offset === 0) fs.writeFileSync(temporary, Buffer.alloc(0));
    fs.writeFileSync(metadataPath, `${JSON.stringify(metadata)}\n`);
    const file = fs.openSync(temporary, 'a');
    try {
      while (offset < asset.bytes) {
        const end = Math.min(asset.bytes - 1, offset + rangeBytes - 1);
        const bytes = Buffer.from(
          await fetchWithRetry(
            `${sourceOrigin}/r/${encodeURIComponent(releaseId)}/${asset.path}`,
            async (response) => {
              if (
                response.status !== 206 ||
                !parseContentRange(response.headers.get('content-range'), offset, end, asset.bytes)
              ) {
                throw new Error(
                  `Range response identity mismatch for ${asset.path}: ${response.status} ${response.headers.get('content-range') || 'missing range'}`,
                );
              }
              const body = Buffer.from(await response.arrayBuffer());
              if (body.byteLength !== end - offset + 1) {
                throw new Error(`Range response length mismatch for ${asset.path}`);
              }
              return body;
            },
            options.retryDelays,
            { headers: { Range: `bytes=${offset}-${end}` } },
          ),
        );
        fs.writeSync(file, bytes);
        digest.update(bytes);
        offset += bytes.byteLength;
        fs.writeFileSync(metadataPath, `${JSON.stringify({ ...metadata, offset })}\n`);
      }
    } finally {
      fs.closeSync(file);
    }
    const completeDigest = digest.digest('hex');
    if (completeDigest !== asset.sha256) {
      fs.rmSync(temporary, { force: true });
      fs.rmSync(metadataPath, { force: true });
      throw new Error(`Font asset integrity mismatch: ${asset.path}`);
    }
    fs.renameSync(temporary, target);
    fs.rmSync(metadataPath, { force: true });
    return true;
  } catch (error) {
    // Keep a verified prefix and its identity sidecar so a later CI retry can
    // resume at the last completed Range instead of downloading the font from
    // byte zero again. A digest mismatch is not resumable and must restart.
    if (fs.existsSync(temporary) && fs.statSync(temporary).size > asset.bytes) {
      fs.rmSync(temporary, { force: true });
      fs.rmSync(metadataPath, { force: true });
    }
    throw error;
  }
}

async function downloadAsset(releaseId, asset) {
  const target = safeAssetPath(cacheRoot, asset.path);
  if (validCachedAsset(asset)) return false;
  if (asset.bytes >= RANGED_FONT_ASSET_THRESHOLD_BYTES) return downloadRangedAsset(releaseId, asset);
  const bytes = Buffer.from(
    await fetchWithRetry(`${sourceOrigin}/r/${encodeURIComponent(releaseId)}/${asset.path}`, (response) =>
      response.arrayBuffer(),
    ),
  );
  if (bytes.byteLength !== asset.bytes || sha256(bytes) !== asset.sha256) {
    throw new Error(`Font asset integrity mismatch: ${asset.path}`);
  }
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const temporary = `${target}.download`;
  fs.writeFileSync(temporary, bytes);
  fs.renameSync(temporary, target);
  return true;
}

async function hydrateCache(releaseId, assets) {
  let cursor = 0;
  let downloaded = 0;
  const concurrency = Math.min(fontDownloadConcurrency(), assets.length);
  await Promise.all(
    Array.from({ length: concurrency }, async () => {
      while (cursor < assets.length) {
        const asset = assets[cursor++];
        if (await downloadAsset(releaseId, asset)) downloaded += 1;
        process.stdout.write(
          `Verified Cloudflare font ${cursor}/${assets.length}${downloaded ? ` (${downloaded} downloaded)` : ''}\n`,
        );
      }
    }),
  );
  fs.writeFileSync(
    path.join(cacheRoot, 'matrix-font-release.json'),
    `${JSON.stringify(
      {
        version: 1,
        sourceOrigin,
        releaseId,
        assets: assets.map(({ path: assetPath, bytes, sha256: digest }) => ({
          path: assetPath,
          bytes,
          sha256: digest,
        })),
      },
      null,
      2,
    )}\n`,
  );
}

function copyToDist(assets) {
  for (const asset of assets) {
    const source = safeAssetPath(cacheRoot, asset.path);
    if (!validCachedAsset(asset)) throw new Error(`Font cache verification failed: ${asset.path}`);
    const destination = safeAssetPath(distRoot, asset.path);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.copyFileSync(source, destination);
  }
}

export async function main() {
  if (localFontRoot) {
    const manifestPath = path.join(localFontRoot, 'onlyoffice-browser-font-assets.json');
    const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
    if (manifest.fontSet !== 'full' || !Array.isArray(manifest.fontFamilies) || manifest.fontFamilies.length < 45) {
      throw new Error('Local Cloudflare matrix requires the verified curated full font set with at least 45 families');
    }
    const paths = ['onlyoffice-browser-font-assets.json', ...(manifest.assets || []).map((asset) => asset.path)];
    const assets = [...new Set(paths)].map((assetPath) => {
      const source = path.join(localFontRoot, assetPath);
      if (!fs.existsSync(source) || !fs.statSync(source).isFile()) {
        throw new Error(`Local Cloudflare matrix font asset is missing: ${assetPath}`);
      }
      return {
        path: assetPath,
        bytes: fs.statSync(source).size,
        sha256: sha256(fs.readFileSync(source)),
      };
    });
    for (const asset of assets) {
      const destination = path.join(distRoot, asset.path);
      fs.mkdirSync(path.dirname(destination), { recursive: true });
      fs.copyFileSync(path.join(localFontRoot, asset.path), destination);
    }
    const policy = applyFontPickerPolicy(distRoot);
    process.stdout.write(
      `Hydrated ${assets.length} verified local full-font objects (${assets.reduce((sum, asset) => sum + asset.bytes, 0)} bytes, ${policy.visibleNames.length} visible families)\n`,
    );
  } else {
    const { releaseId, assets } = await readRelease();
    fs.mkdirSync(cacheRoot, { recursive: true });
    await hydrateCache(releaseId, assets);
    copyToDist(assets);
    const policy = applyFontPickerPolicy(distRoot);
    process.stdout.write(
      `Hydrated ${assets.length} verified font objects (${assets.reduce((sum, asset) => sum + asset.bytes, 0)} bytes, ${policy.visibleNames.length} visible families) from ${releaseId}\n`,
    );
  }
}

function isDirectRun() {
  if (!process.argv[1]) return false;
  try {
    return fs.realpathSync(process.argv[1]) === fs.realpathSync(fileURLToPath(import.meta.url));
  } catch {
    return import.meta.url === pathToFileURL(process.argv[1]).href;
  }
}

if (isDirectRun()) {
  await main();
}
