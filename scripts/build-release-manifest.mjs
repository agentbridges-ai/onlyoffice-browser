#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const TARGET_CHUNK_BYTES = 24 * 1024 * 1024;
const MAX_CHUNK_BYTES = 32 * 1024 * 1024;
const PROFILES = ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'];

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function walkFiles(root, directory = root) {
  const files = [];
  for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
    const absolute = path.join(directory, entry.name);
    if (entry.isDirectory()) files.push(...walkFiles(root, absolute));
    else if (entry.isFile()) files.push(path.relative(root, absolute).split(path.sep).join('/'));
  }
  return files;
}

function mimeFor(file) {
  const extension = path.extname(file).toLowerCase();
  return (
    {
      '.bin': 'application/octet-stream',
      '.css': 'text/css; charset=utf-8',
      '.eot': 'application/vnd.ms-fontobject',
      '.html': 'text/html; charset=utf-8',
      '.js': 'text/javascript; charset=utf-8',
      '.json': 'application/json; charset=utf-8',
      '.otc': 'font/collection',
      '.otf': 'font/otf',
      '.png': 'image/png',
      '.svg': 'image/svg+xml',
      '.ttc': 'font/collection',
      '.tte': 'font/ttf',
      '.ttf': 'font/ttf',
      '.wasm': 'application/wasm',
      '.woff': 'font/woff',
      '.woff2': 'font/woff2',
    }[extension] || 'application/octet-stream'
  );
}

function profileFor(asset, fontManifest) {
  if (asset.pack === 'word') return 'word';
  if (asset.pack === 'cell') return 'cell';
  if (asset.pack === 'slide') return 'slide';
  if (asset.pack === 'fonts') {
    const basic = new Set([
      ...(fontManifest.defaultFonts || []),
      ...(fontManifest.builtInFonts || []),
      fontManifest.allFonts,
      fontManifest.fontSelection,
      fontManifest.fontSourceMap,
      ...(fontManifest.fontThumbnails || []),
    ]);
    return basic.has(asset.path) ? 'fonts-basic' : 'fonts-office-compat';
  }
  return 'base';
}

export function chunkReleaseAssets(assets, targetBytes = TARGET_CHUNK_BYTES, maxBytes = MAX_CHUNK_BYTES) {
  const chunks = [];
  for (const profile of PROFILES) {
    const selected = assets
      .filter((asset) => asset.profile === profile)
      .sort((left, right) => left.path.localeCompare(right.path));
    let current = [];
    let currentBytes = 0;
    const flush = () => {
      if (!current.length) return;
      const id = `${profile}-${String(chunks.filter((chunk) => chunk.profile === profile).length + 1).padStart(3, '0')}`;
      chunks.push({ id, profile, bytes: currentBytes, paths: current.map((asset) => asset.path) });
      for (const asset of current) asset.chunk = id;
      current = [];
      currentBytes = 0;
    };
    for (const asset of selected) {
      if (asset.bytes > maxBytes) {
        flush();
        current = [asset];
        currentBytes = asset.bytes;
        flush();
        continue;
      }
      if (current.length && (currentBytes + asset.bytes > maxBytes || currentBytes >= targetBytes)) flush();
      current.push(asset);
      currentBytes += asset.bytes;
    }
    flush();
  }
  return chunks;
}

function linkOrCopy(source, destination) {
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  try {
    fs.linkSync(source, destination);
  } catch {
    fs.copyFileSync(source, destination);
  }
}

export function buildRelease({ root, output, packageVersion, x2tVersion, x2tCommit }) {
  const runtimePath = path.join(root, 'onlyoffice-runtime-assets.json');
  const fontPath = path.join(root, 'onlyoffice-browser-font-assets.json');
  const runtimeBytes = fs.readFileSync(runtimePath);
  const hasFontManifest = fs.existsSync(fontPath);
  const fontBytes = hasFontManifest ? fs.readFileSync(fontPath) : Buffer.from('{"assets":[]}\n');
  const runtime = JSON.parse(runtimeBytes);
  const fonts = JSON.parse(fontBytes);
  const inventory = new Map();
  for (const asset of [...(runtime.assets || []), ...(fonts.assets || [])]) {
    if (asset?.path && !inventory.has(asset.path)) inventory.set(asset.path, asset);
  }
  for (const manifestPath of [
    'onlyoffice-runtime-assets.json',
    ...(hasFontManifest ? ['onlyoffice-browser-font-assets.json'] : []),
  ]) {
    const absolute = path.join(root, manifestPath);
    inventory.set(manifestPath, { path: manifestPath, bytes: fs.statSync(absolute).size, pack: 'core' });
  }
  for (const deployPath of walkFiles(root)) {
    if (deployPath.endsWith('.br') || deployPath.endsWith('.map')) continue;
    if (!inventory.has(deployPath)) {
      inventory.set(deployPath, {
        path: deployPath,
        bytes: fs.statSync(path.join(root, deployPath)).size,
        pack: 'core',
      });
    }
  }
  const assets = [...inventory.values()]
    .map((asset) => {
      const absolute = path.join(root, asset.path);
      const bytes = fs.readFileSync(absolute);
      return {
        path: asset.path,
        bytes: bytes.byteLength,
        mime: mimeFor(asset.path),
        sha256: sha256(bytes),
        profile: profileFor(asset, fonts),
        chunk: '',
      };
    })
    .sort((left, right) => left.path.localeCompare(right.path));
  const chunks = chunkReleaseAssets(assets);
  const hostAssets = assets.filter(
    (asset) => asset.path === 'office-host.html' || /^assets\/officeHost-[^/]+\.js$/.test(asset.path),
  );
  const shellAssets = assets.filter(
    (asset) =>
      asset.path === 'index.html' ||
      asset.path === 'sw.js' ||
      /^assets\/(?:main|base)-[^/]+\.(?:js|css)$/.test(asset.path),
  );
  const x2tAsset =
    assets.find((asset) => /(?:^|\/)x2t\.wasm$/.test(asset.path)) ||
    assets.find((asset) => asset.path.endsWith('.wasm'));
  if (!hostAssets.length || !shellAssets.length || !x2tAsset) {
    throw new Error('Release identity requires Host, shell, and x2t artifacts');
  }
  const profiles = Object.fromEntries(
    PROFILES.map((profile) => [
      profile,
      assets.filter((asset) => asset.profile === profile).map((asset) => asset.path),
    ]),
  );
  const identity = {
    packageVersion,
    hostBuildId: sha256(hostAssets.map((asset) => `${asset.path}\0${asset.sha256}\n`).join('')),
    shellRevision: sha256(shellAssets.map((asset) => `${asset.path}\0${asset.sha256}\n`).join('')),
    runtimeManifestSha256: sha256(runtimeBytes),
    fontManifestSha256: sha256(fontBytes),
    x2t: {
      version: x2tVersion,
      commit: x2tCommit,
      sha256: x2tAsset.sha256,
    },
    profiles,
    chunks,
    assets,
    fontFamilies: fonts.fontFamilies || [],
  };
  const releaseId = `v${packageVersion}-${sha256(JSON.stringify(identity)).slice(0, 16)}`;
  const manifest = { version: 3, releaseId, ...identity };
  fs.rmSync(output, { recursive: true, force: true });
  for (const asset of assets) {
    linkOrCopy(path.join(root, asset.path), path.join(output, 'blobs', 'sha256', asset.sha256));
  }
  const releaseDir = path.join(output, 'releases', releaseId);
  fs.mkdirSync(releaseDir, { recursive: true });
  fs.writeFileSync(path.join(releaseDir, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  fs.mkdirSync(path.join(output, 'channels'), { recursive: true });
  fs.writeFileSync(
    path.join(output, 'channels', 'stable.json'),
    `${JSON.stringify({ version: 1, releaseId }, null, 2)}\n`,
  );
  return manifest;
}

function option(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const manifest = buildRelease({
    root: path.resolve(option('--root', 'dist')),
    output: path.resolve(option('--output', '.onlyoffice-release')),
    packageVersion: option('--package-version', process.env.npm_package_version || '0.0.0'),
    x2tVersion: option('--x2t-version', process.env.ONLYOFFICE_X2T_VERSION || '9.3.0+1'),
    x2tCommit: option('--x2t-commit', process.env.ONLYOFFICE_X2T_COMMIT || 'unknown'),
  });
  console.log(`Built immutable OnlyOffice release ${manifest.releaseId} with ${manifest.assets.length} assets`);
}
