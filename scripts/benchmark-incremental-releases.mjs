#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { Readable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { fileURLToPath } from 'node:url';

const DEFAULT_ORIGIN = 'https://onlyoffice.getpi.work';
const FASTCDC_AVERAGES = [256 * 1024, 1024 * 1024, 4 * 1024 * 1024];
const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const FASTCDC_MANIFEST = path.resolve(SCRIPT_DIR, '../tools/fastcdc-index/Cargo.toml');

function uniqueDescriptors(items) {
  const descriptors = new Map();
  for (const item of items || []) {
    const digest = item?.sha256;
    const bytes = item?.bytes;
    if (typeof digest !== 'string' || typeof bytes !== 'number') continue;
    const existing = descriptors.get(digest);
    if (existing !== undefined && existing !== bytes) {
      throw new Error(`digest ${digest} has conflicting byte lengths`);
    }
    descriptors.set(digest, bytes);
  }
  return descriptors;
}

export function compareContentAddressedItems(fromItems, toItems) {
  const from = uniqueDescriptors(fromItems);
  const to = uniqueDescriptors(toItems);
  let downloadBytes = 0;
  let downloadObjects = 0;
  let reusedBytes = 0;
  let reusedObjects = 0;
  for (const [digest, bytes] of to) {
    if (from.has(digest)) {
      reusedBytes += bytes;
      reusedObjects += 1;
    } else {
      downloadBytes += bytes;
      downloadObjects += 1;
    }
  }
  return {
    targetBytes: [...to.values()].reduce((total, bytes) => total + bytes, 0),
    targetObjects: to.size,
    downloadBytes,
    downloadObjects,
    reusedBytes,
    reusedObjects,
  };
}

export function compareAssetPaths(fromAssets, toAssets) {
  const from = new Map((fromAssets || []).map((asset) => [asset.path, asset]));
  const to = new Map((toAssets || []).map((asset) => [asset.path, asset]));
  let unchanged = 0;
  let changed = 0;
  let added = 0;
  let removed = 0;
  for (const [assetPath, asset] of to) {
    const previous = from.get(assetPath);
    if (!previous) added += 1;
    else if (previous.sha256 === asset.sha256) unchanged += 1;
    else changed += 1;
  }
  for (const assetPath of from.keys()) {
    if (!to.has(assetPath)) removed += 1;
  }
  return { unchanged, changed, added, removed };
}

export function compareReleasePair(fromManifest, toManifest, fastCdcIndexes) {
  const fixedSegments = compareContentAddressedItems(fromManifest.package?.segments, toManifest.package?.segments);
  const fileCas = compareContentAddressedItems(fromManifest.assets, toManifest.assets);
  const result = {
    fromReleaseId: fromManifest.releaseId,
    toReleaseId: toManifest.releaseId,
    logicalTargetBytes: toManifest.package.bytes,
    assetPaths: compareAssetPaths(fromManifest.assets, toManifest.assets),
    currentReleaseBoundSegments: {
      targetBytes: toManifest.package.bytes,
      targetObjects: toManifest.package.segments.length,
      downloadBytes: toManifest.package.bytes,
      downloadObjects: toManifest.package.segments.length,
      reusedBytes: 0,
      reusedObjects: 0,
    },
    fixedDigestSegments: fixedSegments,
    fileCas,
    fastCdc: [],
  };

  if (fastCdcIndexes) {
    for (const fromConfiguration of fastCdcIndexes.from.configurations) {
      const toConfiguration = fastCdcIndexes.to.configurations.find(
        (candidate) => candidate.averageBytes === fromConfiguration.averageBytes,
      );
      if (!toConfiguration) continue;
      result.fastCdc.push({
        minimumBytes: toConfiguration.minimumBytes,
        averageBytes: toConfiguration.averageBytes,
        maximumBytes: toConfiguration.maximumBytes,
        fromElapsedMs: fromConfiguration.elapsedMs,
        toElapsedMs: toConfiguration.elapsedMs,
        ...compareContentAddressedItems(fromConfiguration.chunks, toConfiguration.chunks),
      });
    }
  }
  return result;
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  const units = ['KiB', 'MiB', 'GiB'];
  let value = bytes;
  let unit = '';
  for (const candidate of units) {
    value /= 1024;
    unit = candidate;
    if (value < 1024) break;
  }
  return `${value.toFixed(value >= 100 ? 1 : 2)} ${unit}`;
}

function savedPercent(result) {
  if (!result.targetBytes) return '0.00%';
  return `${((result.reusedBytes / result.targetBytes) * 100).toFixed(2)}%`;
}

export function formatBenchmarkMarkdown(report) {
  const lines = ['# OnlyOffice incremental release benchmark', '', `Generated: ${report.generatedAt}`, ''];
  for (const pair of report.pairs) {
    lines.push(
      `## ${pair.fromReleaseId} → ${pair.toReleaseId}`,
      '',
      `Target package: ${formatBytes(pair.logicalTargetBytes)}; asset paths: ${pair.assetPaths.unchanged} unchanged, ${pair.assetPaths.changed} changed, ${pair.assetPaths.added} added, ${pair.assetPaths.removed} removed.`,
      '',
      '| Strategy | Update download | Update objects | Cold objects | Reused | Reuse |',
      '| --- | ---: | ---: | ---: | ---: | ---: |',
    );
    const rows = [
      ['Current release-bound 24 MiB segments', pair.currentReleaseBoundSegments],
      ['Digest-addressed fixed 24 MiB segments', pair.fixedDigestSegments],
      ['File-level CAS', pair.fileCas],
      ...pair.fastCdc.map((result) => [`FastCDC avg ${formatBytes(result.averageBytes)}`, result]),
    ];
    for (const [label, result] of rows) {
      lines.push(
        `| ${label} | ${formatBytes(result.downloadBytes)} | ${result.downloadObjects} | ${result.targetObjects} | ${formatBytes(result.reusedBytes)} | ${savedPercent(result)} |`,
      );
    }
    lines.push('');
  }
  return `${lines.join('\n')}\n`;
}

async function sha256File(filePath) {
  const digest = crypto.createHash('sha256');
  for await (const chunk of fs.createReadStream(filePath)) digest.update(chunk);
  return digest.digest('hex');
}

async function downloadPack(origin, manifest, workDir) {
  const destination = path.join(workDir, `${manifest.releaseId}.oobpack`);
  fs.mkdirSync(workDir, { recursive: true });
  let currentBytes = fs.existsSync(destination) ? fs.statSync(destination).size : 0;
  if (currentBytes > manifest.package.bytes) {
    fs.truncateSync(destination, 0);
    currentBytes = 0;
  }

  for (let attempt = 1; attempt <= 3 && currentBytes < manifest.package.bytes; attempt += 1) {
    const headers = currentBytes > 0 ? { Range: `bytes=${currentBytes}-` } : {};
    const response = await fetch(`${origin}/p/${encodeURIComponent(manifest.releaseId)}/office-resources.oobpack`, {
      headers,
    });
    if (!response.ok || !response.body) {
      throw new Error(`package request for ${manifest.releaseId} failed (${response.status})`);
    }
    const serverResumed = currentBytes > 0 && response.status === 206;
    const flags = serverResumed ? 'a' : 'w';
    if (!serverResumed) currentBytes = 0;
    await pipeline(Readable.fromWeb(response.body), fs.createWriteStream(destination, { flags }));
    currentBytes = fs.statSync(destination).size;
  }
  if (currentBytes !== manifest.package.bytes) {
    throw new Error(`${manifest.releaseId} package has ${currentBytes} bytes, expected ${manifest.package.bytes}`);
  }
  const digest = await sha256File(destination);
  if (digest !== manifest.package.sha256) {
    throw new Error(`${manifest.releaseId} package SHA-256 mismatch`);
  }
  return destination;
}

function buildFastCdcIndex(packPath) {
  const result = spawnSync(
    'cargo',
    [
      'run',
      '--quiet',
      '--release',
      '--manifest-path',
      FASTCDC_MANIFEST,
      '--',
      packPath,
      ...FASTCDC_AVERAGES.map(String),
    ],
    { encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 },
  );
  if (result.status !== 0) {
    throw new Error(result.stderr || `FastCDC indexer exited with ${result.status}`);
  }
  return JSON.parse(result.stdout);
}

function valuesAfter(arguments_, name) {
  const values = [];
  for (let index = 0; index < arguments_.length; index += 1) {
    if (arguments_[index] === name && arguments_[index + 1]) values.push(arguments_[index + 1]);
  }
  return values;
}

function valueAfter(arguments_, name, fallback) {
  const index = arguments_.indexOf(name);
  return index >= 0 && arguments_[index + 1] ? arguments_[index + 1] : fallback;
}

async function run() {
  const arguments_ = process.argv.slice(2);
  const releaseIds = valuesAfter(arguments_, '--release');
  if (releaseIds.length < 2) {
    throw new Error('pass at least two production release IDs with repeated --release options');
  }
  const origin = valueAfter(arguments_, '--origin', DEFAULT_ORIGIN).replace(/\/+$/, '');
  const output = valueAfter(arguments_, '--output', '');
  const workDir = valueAfter(arguments_, '--work-dir', '');
  const withFastCdc = arguments_.includes('--fastcdc');
  if (withFastCdc && !workDir) {
    throw new Error('--fastcdc requires --work-dir outside the repository for downloaded packages');
  }

  const manifests = [];
  for (const releaseId of releaseIds) {
    const response = await fetch(`${origin}/releases/${encodeURIComponent(releaseId)}/manifest.json`, {
      cache: 'no-store',
    });
    if (!response.ok) throw new Error(`manifest request for ${releaseId} failed (${response.status})`);
    const manifest = await response.json();
    if (manifest.releaseId !== releaseId || manifest.version !== 4) {
      throw new Error(`manifest identity mismatch for ${releaseId}`);
    }
    manifests.push(manifest);
  }

  const indexes = new Map();
  if (withFastCdc) {
    for (const manifest of manifests) {
      const packPath = await downloadPack(origin, manifest, path.resolve(workDir));
      indexes.set(manifest.releaseId, buildFastCdcIndex(packPath));
    }
  }

  const report = {
    version: 1,
    generatedAt: new Date().toISOString(),
    origin,
    releases: releaseIds,
    fastCdcAverages: withFastCdc ? FASTCDC_AVERAGES : [],
    pairs: [],
  };
  for (let index = 1; index < manifests.length; index += 1) {
    const from = manifests[index - 1];
    const to = manifests[index];
    report.pairs.push(
      compareReleasePair(
        from,
        to,
        withFastCdc ? { from: indexes.get(from.releaseId), to: indexes.get(to.releaseId) } : undefined,
      ),
    );
  }
  const markdown = formatBenchmarkMarkdown(report);
  if (output) {
    const outputPath = path.resolve(output);
    fs.mkdirSync(path.dirname(outputPath), { recursive: true });
    fs.writeFileSync(outputPath, markdown);
    fs.writeFileSync(`${outputPath}.json`, `${JSON.stringify(report, null, 2)}\n`);
  }
  process.stdout.write(markdown);
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  run().catch((error) => {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  });
}
