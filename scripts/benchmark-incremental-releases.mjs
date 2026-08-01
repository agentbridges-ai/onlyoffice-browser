#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DEFAULT_ORIGIN = 'https://onlyoffice.getpi.work';
const DEFAULT_MANIFEST_VERSION = 5;

function addDescriptor(descriptors, item, label = 'content object') {
  const digest = item?.sha256;
  const bytes = item?.bytes;
  if (typeof digest !== 'string' || !Number.isSafeInteger(bytes) || bytes < 0) {
    throw new Error(`${label} has an invalid identity`);
  }
  const existing = descriptors.get(digest);
  if (existing !== undefined && existing !== bytes) {
    throw new Error(`digest ${digest} has conflicting byte lengths`);
  }
  descriptors.set(digest, bytes);
}

function uniqueDescriptors(items) {
  const descriptors = new Map();
  for (const item of items || []) addDescriptor(descriptors, item);
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

function wholeDescriptor(asset) {
  const whole = asset?.representations?.whole;
  if (whole) {
    if (whole.sha256 !== asset.sha256 || whole.bytes !== asset.bytes) {
      throw new Error(`asset ${asset.path} has an invalid whole representation`);
    }
    return whole;
  }
  return { sha256: asset.sha256, bytes: asset.bytes };
}

function fastCdcDescriptors(asset) {
  const fastCdc = asset?.representations?.fastcdc;
  if (!fastCdc) return null;
  if (!Array.isArray(fastCdc.chunks) || fastCdc.chunks.length === 0) {
    throw new Error(`asset ${asset.path} has an empty FastCDC representation`);
  }
  let expectedOffset = 0;
  const chunks = fastCdc.chunks.map((chunk, index) => {
    if (
      chunk?.offset !== expectedOffset ||
      !Number.isSafeInteger(chunk.bytes) ||
      chunk.bytes <= 0 ||
      typeof chunk.sha256 !== 'string'
    ) {
      throw new Error(`asset ${asset.path} has an invalid FastCDC chunk at index ${index}`);
    }
    expectedOffset += chunk.bytes;
    return { sha256: chunk.sha256, bytes: chunk.bytes };
  });
  if (expectedOffset !== asset.bytes) {
    throw new Error(`asset ${asset.path} FastCDC chunks do not cover the complete file`);
  }
  uniqueDescriptors(chunks);
  return chunks;
}

/**
 * Models the v5 path-by-path update rule rather than applying content-defined
 * chunking to the complete Office Pack:
 *
 * - an unchanged path reuses its active mapping and downloads nothing;
 * - an ordinary changed/added path downloads its whole content-addressed blob;
 * - a changed/added path with a v5 FastCDC representation downloads only chunks
 *   absent from the previous manifest's representation for the same path.
 *
 * Object totals are de-duplicated by SHA-256 because the canonical store keeps
 * one physical copy even when several paths reference the same object.
 */
export function compareHybridPlanner(fromManifest, toManifest) {
  const fromByPath = new Map((fromManifest.assets || []).map((asset) => [asset.path, asset]));
  const toByPath = new Map((toManifest.assets || []).map((asset) => [asset.path, asset]));
  const targetObjects = new Map();
  const downloadObjects = new Map();
  const decisions = [];
  const paths = { unchanged: 0, whole: 0, fastCdc: 0, removed: 0 };
  let targetBytes = 0;

  for (const asset of [...toByPath.values()].sort((left, right) => left.path.localeCompare(right.path))) {
    if (!Number.isSafeInteger(asset.bytes) || asset.bytes < 0 || typeof asset.sha256 !== 'string') {
      throw new Error(`asset ${asset.path} has an invalid identity`);
    }
    targetBytes += asset.bytes;
    const previous = fromByPath.get(asset.path);
    const targetFastCdc = toManifest.version === 5 ? fastCdcDescriptors(asset) : null;
    const selectedObjects = targetFastCdc || [wholeDescriptor(asset)];
    for (const descriptor of selectedObjects) addDescriptor(targetObjects, descriptor, `asset ${asset.path}`);

    if (previous?.sha256 === asset.sha256) {
      paths.unchanged += 1;
      decisions.push({
        path: asset.path,
        change: 'unchanged',
        strategy: 'unchanged',
        targetBytes: asset.bytes,
        downloadBytes: 0,
        downloadObjects: 0,
      });
      continue;
    }

    if (!targetFastCdc) {
      const descriptor = wholeDescriptor(asset);
      const alreadyPlanned = downloadObjects.has(descriptor.sha256);
      addDescriptor(downloadObjects, descriptor, `asset ${asset.path}`);
      paths.whole += 1;
      decisions.push({
        path: asset.path,
        change: previous ? 'changed' : 'added',
        strategy: 'whole',
        targetBytes: asset.bytes,
        downloadBytes: alreadyPlanned ? 0 : descriptor.bytes,
        downloadObjects: alreadyPlanned ? 0 : 1,
      });
      continue;
    }

    const previousChunks = uniqueDescriptors(fromManifest.version === 5 ? fastCdcDescriptors(previous) || [] : []);
    let pathDownloadBytes = 0;
    let pathDownloadObjects = 0;
    for (const descriptor of targetFastCdc) {
      if (previousChunks.get(descriptor.sha256) === descriptor.bytes || downloadObjects.has(descriptor.sha256)) {
        continue;
      }
      addDescriptor(downloadObjects, descriptor, `asset ${asset.path}`);
      pathDownloadBytes += descriptor.bytes;
      pathDownloadObjects += 1;
    }
    paths.fastCdc += 1;
    decisions.push({
      path: asset.path,
      change: previous ? 'changed' : 'added',
      strategy: 'fastcdc',
      targetBytes: asset.bytes,
      downloadBytes: pathDownloadBytes,
      downloadObjects: pathDownloadObjects,
    });
  }

  for (const asset of [...fromByPath.values()].sort((left, right) => left.path.localeCompare(right.path))) {
    if (toByPath.has(asset.path)) continue;
    paths.removed += 1;
    decisions.push({
      path: asset.path,
      change: 'removed',
      strategy: 'removed',
      targetBytes: 0,
      downloadBytes: 0,
      downloadObjects: 0,
    });
  }

  const downloadBytes = [...downloadObjects.values()].reduce((total, bytes) => total + bytes, 0);
  return {
    targetBytes,
    targetObjects: targetObjects.size,
    downloadBytes,
    downloadObjects: downloadObjects.size,
    reusedBytes: Math.max(0, targetBytes - downloadBytes),
    reusedObjects: Math.max(0, targetObjects.size - downloadObjects.size),
    paths,
    decisions,
  };
}

export function compareReleasePair(fromManifest, toManifest) {
  const targetPackageBytes =
    toManifest.package?.bytes ||
    (toManifest.assets || []).reduce((total, asset) => total + (Number(asset.bytes) || 0), 0);
  const fixedSegments = compareContentAddressedItems(
    fromManifest.package?.segments || [],
    toManifest.package?.segments || [],
  );
  const fileCas = compareContentAddressedItems(fromManifest.assets, toManifest.assets);
  return {
    fromReleaseId: fromManifest.releaseId,
    toReleaseId: toManifest.releaseId,
    fromManifestVersion: fromManifest.version,
    toManifestVersion: toManifest.version,
    logicalTargetBytes: targetPackageBytes,
    assetPaths: compareAssetPaths(fromManifest.assets, toManifest.assets),
    currentReleaseBoundSegments: {
      targetBytes: targetPackageBytes,
      targetObjects: toManifest.package?.segments?.length || 0,
      downloadBytes: targetPackageBytes,
      downloadObjects: toManifest.package?.segments?.length || 0,
      reusedBytes: 0,
      reusedObjects: 0,
    },
    fixedDigestSegments: fixedSegments,
    fileCas,
    hybridPlanner: compareHybridPlanner(fromManifest, toManifest),
    // Kept as an empty compatibility field for consumers of report v1. The
    // whole-pack FastCDC experiment is historical evidence, not the v5 plan.
    fastCdc: [],
  };
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
    const hybrid = pair.hybridPlanner;
    lines.push(
      `## ${pair.fromReleaseId} → ${pair.toReleaseId}`,
      '',
      `Manifest: v${pair.fromManifestVersion ?? '?'} → v${pair.toManifestVersion ?? '?'}; target package: ${formatBytes(pair.logicalTargetBytes)}; asset paths: ${pair.assetPaths.unchanged} unchanged, ${pair.assetPaths.changed} changed, ${pair.assetPaths.added} added, ${pair.assetPaths.removed} removed.`,
      '',
      '| Strategy | Update download | Update objects | Target objects | Reused | Reuse |',
      '| --- | ---: | ---: | ---: | ---: | ---: |',
    );
    const rows = [
      ['Release-bound package segments', pair.currentReleaseBoundSegments],
      ['Digest-addressed fixed package segments', pair.fixedDigestSegments],
      ['File-level whole CAS', pair.fileCas],
      ['Hybrid path planner', hybrid],
    ];
    for (const [label, result] of rows) {
      lines.push(
        `| ${label} | ${formatBytes(result.downloadBytes)} | ${result.downloadObjects} | ${result.targetObjects} | ${formatBytes(result.reusedBytes)} | ${savedPercent(result)} |`,
      );
    }
    lines.push(
      '',
      `Hybrid paths: ${hybrid.paths.unchanged} unchanged (zero download), ${hybrid.paths.whole} whole blobs, ${hybrid.paths.fastCdc} per-file FastCDC, ${hybrid.paths.removed} removed.`,
      '',
    );
  }
  return `${lines.join('\n')}\n`;
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

export function releaseManifestPath(releaseId, version = DEFAULT_MANIFEST_VERSION) {
  if (version !== 4 && version !== 5) throw new Error('manifest version must be 4 or 5');
  const file = version === 4 ? 'manifest-v4.json' : 'manifest.json';
  return `/releases/${encodeURIComponent(releaseId)}/${file}`;
}

async function fetchReleaseManifest(origin, releaseId, version) {
  const paths = [
    releaseManifestPath(releaseId, version),
    ...(version === 4 ? [`/releases/${encodeURIComponent(releaseId)}/manifest.json`] : []),
  ];
  for (const [index, manifestPath] of paths.entries()) {
    const response = await fetch(`${origin}${manifestPath}`, { cache: 'no-store' });
    if (response.status === 404 && index < paths.length - 1) continue;
    if (!response.ok) throw new Error(`manifest request for ${releaseId} failed (${response.status})`);
    const manifest = await response.json();
    if (manifest.releaseId !== releaseId || manifest.version !== version) {
      throw new Error(`manifest identity mismatch for ${releaseId}`);
    }
    return manifest;
  }
  throw new Error(`manifest request for ${releaseId} failed`);
}

async function run() {
  const arguments_ = process.argv.slice(2);
  const releaseIds = valuesAfter(arguments_, '--release');
  if (releaseIds.length < 2) {
    throw new Error('pass at least two production release IDs with repeated --release options');
  }
  if (arguments_.includes('--fastcdc') || arguments_.includes('--work-dir')) {
    throw new Error('whole-pack FastCDC options were removed; v5 manifests carry per-file FastCDC metadata');
  }
  const origin = valueAfter(arguments_, '--origin', DEFAULT_ORIGIN).replace(/\/+$/, '');
  const output = valueAfter(arguments_, '--output', '');
  const manifestVersion = Number(valueAfter(arguments_, '--manifest-version', String(DEFAULT_MANIFEST_VERSION)));
  if (manifestVersion !== 4 && manifestVersion !== 5) {
    throw new Error('--manifest-version must be 4 or 5');
  }

  const manifests = [];
  for (const releaseId of releaseIds) {
    manifests.push(await fetchReleaseManifest(origin, releaseId, manifestVersion));
  }

  const report = {
    version: 2,
    generatedAt: new Date().toISOString(),
    origin,
    releases: releaseIds,
    manifestVersion,
    pairs: [],
  };
  for (let index = 1; index < manifests.length; index += 1) {
    report.pairs.push(compareReleasePair(manifests[index - 1], manifests[index]));
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
