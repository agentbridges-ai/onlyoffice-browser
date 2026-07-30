#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const sourceOrigin = (process.env.ONLYOFFICE_MATRIX_FONT_SOURCE || 'https://onlyoffice.getpi.work').replace(/\/+$/, '');
const pinnedReleaseId = process.env.ONLYOFFICE_MATRIX_FONT_RELEASE_ID || '';
const pinnedManifestSha256 = process.env.ONLYOFFICE_MATRIX_FONT_MANIFEST_SHA256 || '';
const cacheRoot = path.resolve(process.env.ONLYOFFICE_MATRIX_FONT_CACHE || '.onlyoffice-cloudflare-matrix-cache');
const distRoot = path.resolve(process.env.ONLYOFFICE_MATRIX_DIST || 'dist');
const localFontRoot = process.env.ONLYOFFICE_MATRIX_FONT_ASSETS_DIR
  ? path.resolve(process.env.ONLYOFFICE_MATRIX_FONT_ASSETS_DIR)
  : '';
const fontProfiles = new Set(['fonts-basic', 'fonts-office-compat']);

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

export async function fetchWithRetry(url, consume = (response) => response, retryDelays = [1_000, 3_000]) {
  let lastError;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 60_000);
    try {
      const response = await fetch(url, { cache: 'no-store', signal: controller.signal });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      return await consume(response);
    } catch (error) {
      lastError = error;
      if (attempt < 2) {
        await new Promise((resolve) => setTimeout(resolve, retryDelays[attempt] ?? 0));
      }
    } finally {
      clearTimeout(timeout);
    }
  }
  throw new Error(`Unable to download ${url}: ${lastError}`);
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
        typeof asset.path !== 'string' || !/^[a-f0-9]{64}$/.test(asset.sha256) || !Number.isSafeInteger(asset.bytes),
    )
  ) {
    throw new Error('Cloudflare matrix font release is incomplete');
  }
  const fontManifest = assets.find((asset) => asset.path === 'onlyoffice-browser-font-assets.json');
  if (pinnedManifestSha256 && fontManifest.sha256 !== pinnedManifestSha256) {
    throw new Error(
      `Pinned Cloudflare font manifest mismatch: expected ${pinnedManifestSha256}, received ${fontManifest.sha256}`,
    );
  }
  return { releaseId, assets };
}

function validCachedAsset(asset) {
  const target = path.join(cacheRoot, asset.path);
  if (!fs.existsSync(target) || fs.statSync(target).size !== asset.bytes) return false;
  return sha256(fs.readFileSync(target)) === asset.sha256;
}

async function downloadAsset(releaseId, asset) {
  if (validCachedAsset(asset)) return false;
  const bytes = Buffer.from(
    await fetchWithRetry(`${sourceOrigin}/r/${encodeURIComponent(releaseId)}/${asset.path}`, (response) =>
      response.arrayBuffer(),
    ),
  );
  if (bytes.byteLength !== asset.bytes || sha256(bytes) !== asset.sha256) {
    throw new Error(`Font asset integrity mismatch: ${asset.path}`);
  }
  const target = path.join(cacheRoot, asset.path);
  fs.mkdirSync(path.dirname(target), { recursive: true });
  const temporary = `${target}.download`;
  fs.writeFileSync(temporary, bytes);
  fs.renameSync(temporary, target);
  return true;
}

async function hydrateCache(releaseId, assets) {
  let cursor = 0;
  let downloaded = 0;
  await Promise.all(
    Array.from({ length: 6 }, async () => {
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
    const source = path.join(cacheRoot, asset.path);
    if (!validCachedAsset(asset)) throw new Error(`Font cache verification failed: ${asset.path}`);
    const destination = path.join(distRoot, asset.path);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.copyFileSync(source, destination);
  }
}

export async function main() {
  if (localFontRoot) {
    const manifestPath = path.join(localFontRoot, 'onlyoffice-browser-font-assets.json');
    const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
    if (manifest.fontSet !== 'full' || !Array.isArray(manifest.fontFamilies) || manifest.fontFamilies.length < 70) {
      throw new Error('Local Cloudflare matrix requires a verified full font set with at least 70 families');
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
    process.stdout.write(
      `Hydrated ${assets.length} verified local full-font objects (${assets.reduce((sum, asset) => sum + asset.bytes, 0)} bytes, ${manifest.fontFamilies.length} families)\n`,
    );
  } else {
    const { releaseId, assets } = await readRelease();
    fs.mkdirSync(cacheRoot, { recursive: true });
    await hydrateCache(releaseId, assets);
    copyToDist(assets);
    process.stdout.write(
      `Hydrated ${assets.length} verified font objects (${assets.reduce((sum, asset) => sum + asset.bytes, 0)} bytes) from ${releaseId}\n`,
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
