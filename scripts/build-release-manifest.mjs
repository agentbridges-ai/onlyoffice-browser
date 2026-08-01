#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  FASTCDC_RELEASE_POLICY,
  buildFastCdcRepresentation,
  buildFastCdcRepresentationFromEvidence,
  parseFastCdcEvidence,
  readFastCdcEvidence,
} from './fastcdc-release-policy.mjs';

const TARGET_CHUNK_BYTES = 24 * 1024 * 1024;
const MAX_CHUNK_BYTES = 32 * 1024 * 1024;
const OFFICE_PACK_MAGIC = Buffer.from('OOBPACK1');
const OFFICE_PACK_SEGMENT_BYTES = 24 * 1024 * 1024;
const PROFILES = ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'];
const FASTCDC_POLICY_ID = 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0';

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function storageSetDescription(pack, assets) {
  return {
    version: 1,
    packageSegments: pack.segments.map((segment) => ({
      offset: segment.offset,
      bytes: segment.bytes,
      sha256: segment.sha256,
    })),
    assets: [...assets]
      .sort((left, right) => left.path.localeCompare(right.path))
      .map((asset) => ({
        path: asset.path,
        bytes: asset.bytes,
        sha256: asset.sha256,
        whole: {
          bytes: asset.representations.whole.bytes,
          sha256: asset.representations.whole.sha256,
        },
        ...(asset.representations.fastcdc
          ? {
              fastcdc: {
                algorithm: asset.representations.fastcdc.algorithm,
                minBytes: asset.representations.fastcdc.minBytes,
                averageBytes: asset.representations.fastcdc.averageBytes,
                maxBytes: asset.representations.fastcdc.maxBytes,
                normalization: asset.representations.fastcdc.normalization,
                seed: asset.representations.fastcdc.seed,
                chunks: asset.representations.fastcdc.chunks.map((chunk) => ({
                  offset: chunk.offset,
                  bytes: chunk.bytes,
                  sha256: chunk.sha256,
                })),
              },
            }
          : {}),
      })),
  };
}

export function computeStorageSetSha256(pack, assets) {
  return sha256(JSON.stringify(storageSetDescription(pack, assets)));
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
      '.oobpack': 'application/vnd.onlyoffice.browser-pack',
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

export function buildOfficePack(root, output, assets, segmentBytes = OFFICE_PACK_SEGMENT_BYTES) {
  const packedAssets = [...assets].sort(
    (left, right) => left.chunk.localeCompare(right.chunk) || left.path.localeCompare(right.path),
  );
  const entries = [];
  let relativeOffset = 0;
  for (const asset of packedAssets) {
    entries.push({
      path: asset.path,
      offset: relativeOffset,
      bytes: asset.bytes,
      mime: asset.mime,
      sha256: asset.sha256,
    });
    relativeOffset += asset.bytes;
  }

  const indexBytes = Buffer.from(JSON.stringify({ version: 1, entries }));
  const header = Buffer.alloc(OFFICE_PACK_MAGIC.byteLength + 4);
  OFFICE_PACK_MAGIC.copy(header, 0);
  header.writeUInt32BE(indexBytes.byteLength, OFFICE_PACK_MAGIC.byteLength);
  const headerBytes = header.byteLength + indexBytes.byteLength;
  for (let index = 0; index < packedAssets.length; index += 1) {
    packedAssets[index].packageOffset = headerBytes + entries[index].offset;
  }

  fs.mkdirSync(output, { recursive: true });
  const temporaryPath = path.join(output, '.office-resources.oobpack.tmp');
  const file = fs.openSync(temporaryPath, 'w');
  const completeDigest = crypto.createHash('sha256');
  let totalBytes = 0;
  const segments = [];

  const writeSegment = (parts) => {
    const offset = totalBytes;
    const digest = crypto.createHash('sha256');
    const temporarySegment = path.join(output, `.office-segment-${segments.length + 1}.tmp`);
    const segmentFile = fs.openSync(temporarySegment, 'w');
    let bytes = 0;
    try {
      for (const part of parts) {
        const value = Buffer.isBuffer(part) ? part : fs.readFileSync(path.join(root, part.path));
        fs.writeSync(file, value);
        fs.writeSync(segmentFile, value);
        completeDigest.update(value);
        digest.update(value);
        bytes += value.byteLength;
        totalBytes += value.byteLength;
      }
    } finally {
      fs.closeSync(segmentFile);
    }
    const segmentSha256 = digest.digest('hex');
    const segmentPath = path.join(output, 'segments', 'sha256', segmentSha256);
    fs.mkdirSync(path.dirname(segmentPath), { recursive: true });
    if (fs.existsSync(segmentPath)) fs.rmSync(temporarySegment);
    else fs.renameSync(temporarySegment, segmentPath);
    segments.push({
      id: segmentSha256,
      offset,
      bytes,
      sha256: segmentSha256,
    });
  };

  try {
    writeSegment([header, indexBytes]);
    let currentChunk = '';
    let currentAssets = [];
    const flushChunk = () => {
      if (!currentAssets.length) return;
      writeSegment(currentAssets);
      currentAssets = [];
    };
    for (const asset of packedAssets) {
      if (currentChunk && asset.chunk !== currentChunk) flushChunk();
      currentChunk = asset.chunk;
      currentAssets.push(asset);
    }
    flushChunk();
  } finally {
    fs.closeSync(file);
  }

  return {
    temporaryPath,
    descriptor: {
      format: 'onlyoffice-pack-v1',
      path: 'office-resources.oobpack',
      bytes: totalBytes,
      sha256: completeDigest.digest('hex'),
      headerBytes,
      segmentBytes,
      segments,
    },
  };
}

function excludeFromRelease(assetPath) {
  return (
    assetPath.endsWith('.br') ||
    assetPath.endsWith('.map') ||
    assetPath.startsWith('npm/') ||
    assetPath.startsWith('.vite/')
  );
}

function profileFor(asset, fontManifest) {
  if (asset.pack === 'word') return 'word';
  if (asset.pack === 'cell') return 'cell';
  if (asset.pack === 'slide') return 'slide';
  const fontAssets = new Set([
    ...(fontManifest.assets || []).map((item) => item?.path).filter(Boolean),
    'server/FileConverter/bin/AllFonts.js',
  ]);
  if (asset.pack === 'fonts' || asset.path.startsWith('fonts/') || fontAssets.has(asset.path)) {
    const basic = new Set([
      ...(fontManifest.defaultFonts || []),
      ...(fontManifest.builtInFonts || []),
      fontManifest.allFonts,
      fontManifest.fontSelection,
      fontManifest.fontSourceMap,
      ...(fontManifest.fontThumbnails || []),
      'server/FileConverter/bin/AllFonts.js',
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

export function buildRelease({
  root,
  output,
  packageVersion,
  x2tVersion,
  x2tCommit,
  fastCdcEvidence,
  fastCdcEvidencePath,
  fastCdcIndexer,
}) {
  const runtimePath = path.join(root, 'onlyoffice-runtime-assets.json');
  const fontPath = path.join(root, 'onlyoffice-browser-font-assets.json');
  const runtimeBytes = fs.readFileSync(runtimePath);
  const hasFontManifest = fs.existsSync(fontPath);
  const fontBytes = hasFontManifest ? fs.readFileSync(fontPath) : Buffer.from('{"assets":[]}\n');
  const runtime = JSON.parse(runtimeBytes);
  const fonts = JSON.parse(fontBytes);
  const inventory = new Map();
  for (const asset of [...(runtime.assets || []), ...(fonts.assets || [])]) {
    if (asset?.path && !excludeFromRelease(asset.path) && !inventory.has(asset.path)) {
      inventory.set(asset.path, asset);
    }
  }
  for (const manifestPath of [
    'onlyoffice-runtime-assets.json',
    ...(hasFontManifest ? ['onlyoffice-browser-font-assets.json'] : []),
  ]) {
    const absolute = path.join(root, manifestPath);
    inventory.set(manifestPath, { path: manifestPath, bytes: fs.statSync(absolute).size, pack: 'core' });
  }
  for (const deployPath of walkFiles(root)) {
    if (excludeFromRelease(deployPath)) {
      continue;
    }
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
  fs.rmSync(output, { recursive: true, force: true });
  const officePack = buildOfficePack(root, output, assets);
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
  const v5Assets = assets.map((asset) => ({
    ...asset,
    representations: {
      whole: {
        sha256: asset.sha256,
        bytes: asset.bytes,
      },
    },
  }));
  if (fastCdcEvidence && fastCdcEvidencePath) {
    throw new Error('Pass FastCDC evidence as an object or a path, not both');
  }
  const parsedFastCdcEvidence = fastCdcEvidence
    ? parseFastCdcEvidence(fastCdcEvidence)
    : fastCdcEvidencePath
      ? readFastCdcEvidence(fastCdcEvidencePath)
      : null;
  for (const asset of v5Assets) {
    if (asset.bytes < FASTCDC_RELEASE_POLICY.minimumAssetBytes) continue;
    if (parsedFastCdcEvidence) {
      const result = buildFastCdcRepresentationFromEvidence({
        assetPath: asset.path,
        inputPath: path.join(root, asset.path),
        output,
        expectedBytes: asset.bytes,
        expectedSha256: asset.sha256,
        evidence: parsedFastCdcEvidence,
        ...(fastCdcIndexer ? { indexer: fastCdcIndexer } : {}),
      });
      if (!result.selected) {
        throw new Error(`FastCDC policy unexpectedly rejected bounded cache object ${asset.path}`);
      }
      asset.representations.fastcdc = result.representation;
    } else {
      asset.representations.fastcdc = buildFastCdcRepresentation({
        inputPath: path.join(root, asset.path),
        output,
        expectedBytes: asset.bytes,
        expectedSha256: asset.sha256,
        ...(fastCdcIndexer ? { indexer: fastCdcIndexer } : {}),
      });
    }
  }
  const contentProtocol = {
    version: 1,
    digest: 'sha256',
    cacheKeyFormat: 'canonical-sha256-v1',
    storageSetSha256: computeStorageSetSha256(officePack.descriptor, v5Assets),
    fastcdcPolicyId: FASTCDC_POLICY_ID,
  };
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
    package: officePack.descriptor,
    contentProtocol,
    assets: v5Assets,
    fontFamilies: fonts.fontFamilies || [],
  };
  const releaseId = `v${packageVersion}-${sha256(JSON.stringify(identity)).slice(0, 16)}`;
  const manifest = { version: 5, releaseId, ...identity };
  const compatibilityManifest = {
    version: 4,
    releaseId,
    packageVersion,
    hostBuildId: identity.hostBuildId,
    shellRevision: identity.shellRevision,
    runtimeManifestSha256: identity.runtimeManifestSha256,
    fontManifestSha256: identity.fontManifestSha256,
    x2t: identity.x2t,
    profiles,
    chunks,
    package: officePack.descriptor,
    assets: assets.map((asset) => ({ ...asset })),
    fontFamilies: identity.fontFamilies,
  };
  for (const asset of assets) {
    linkOrCopy(path.join(root, asset.path), path.join(output, 'blobs', 'sha256', asset.sha256));
  }
  const packagePath = path.join(output, 'packages', 'sha256', `${officePack.descriptor.sha256}.oobpack`);
  fs.mkdirSync(path.dirname(packagePath), { recursive: true });
  fs.renameSync(officePack.temporaryPath, packagePath);
  const releaseDir = path.join(output, 'releases', releaseId);
  fs.mkdirSync(releaseDir, { recursive: true });
  const manifestText = `${JSON.stringify(manifest, null, 2)}\n`;
  const compatibilityManifestText = `${JSON.stringify(compatibilityManifest, null, 2)}\n`;
  fs.writeFileSync(path.join(releaseDir, 'manifest.json'), manifestText);
  fs.writeFileSync(path.join(releaseDir, 'manifest-v4.json'), compatibilityManifestText);
  fs.mkdirSync(path.join(output, 'channels'), { recursive: true });
  fs.writeFileSync(
    path.join(output, 'channels', 'stable.json'),
    `${JSON.stringify(
      {
        version: 1,
        releaseId,
        manifestUrl: `/releases/${releaseId}/manifest-v4.json`,
        manifestSha256: sha256(compatibilityManifestText),
      },
      null,
      2,
    )}\n`,
  );
  fs.writeFileSync(
    path.join(output, 'channels', 'stable-v5.json'),
    `${JSON.stringify(
      {
        version: 1,
        releaseId,
        manifestUrl: `/releases/${releaseId}/manifest.json`,
        manifestSha256: sha256(manifestText),
      },
      null,
      2,
    )}\n`,
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
    x2tVersion: option('--x2t-version', process.env.ONLYOFFICE_X2T_VERSION || '9.3.0+2'),
    x2tCommit: option('--x2t-commit', process.env.ONLYOFFICE_X2T_COMMIT || '1bb9b45a399f87ca162eea0c86abd4660f295469'),
    fastCdcEvidencePath: option('--fastcdc-evidence', process.env.ONLYOFFICE_FASTCDC_EVIDENCE),
  });
  console.log(
    `Built immutable OnlyOffice release ${manifest.releaseId} with one ${manifest.package.bytes}-byte Office Pack`,
  );
}
