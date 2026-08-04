#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;

function fail(message) {
  throw new Error(message);
}

function readJson(file, label) {
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (error) {
    throw new Error(`${label} is not valid JSON: ${error instanceof Error ? error.message : String(error)}`);
  }
}

function requireReleaseId(value, label) {
  if (typeof value !== 'string' || !RELEASE_ID_PATTERN.test(value)) fail(`${label} is invalid`);
  return value;
}

function requireDigest(value, label) {
  if (typeof value !== 'string' || !SHA256_PATTERN.test(value)) fail(`${label} is invalid`);
  return value;
}

function requireSafeKey(value, label) {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.startsWith('/') ||
    value.includes('\\') ||
    value.split('/').some((part) => !part || part === '.' || part === '..')
  ) {
    fail(`${label} is not a safe R2 key`);
  }
  return value;
}

function addExact(keys, value, label) {
  keys.add(requireSafeKey(value, label));
}

function addManifestObjects(keys, manifest) {
  if (manifest?.version !== 5) fail('latest cleanup requires a Release Manifest v5');
  const releaseId = requireReleaseId(manifest.releaseId, 'manifest.releaseId');
  if (!manifest.package || !Array.isArray(manifest.package.segments) || !manifest.package.sha256) {
    fail('latest manifest package descriptor is invalid');
  }
  addExact(
    keys,
    `packages/sha256/${requireDigest(manifest.package.sha256, 'manifest.package.sha256')}.oobpack`,
    'package',
  );
  for (const [index, segment] of manifest.package.segments.entries()) {
    const digest = requireDigest(segment?.sha256, `package segment ${index}.sha256`);
    addExact(keys, `segments/sha256/${digest}`, `package segment ${index}`);
  }
  if (!Array.isArray(manifest.assets) || manifest.assets.length === 0) fail('latest manifest assets are missing');
  for (const [index, asset] of manifest.assets.entries()) {
    const whole = requireDigest(asset?.sha256, `asset ${index}.sha256`);
    addExact(keys, `blobs/sha256/${whole}`, `asset ${index}`);
    const fastcdc = asset?.representations?.fastcdc;
    if (fastcdc) {
      if (!Array.isArray(fastcdc.chunks) || fastcdc.chunks.length === 0) {
        fail(`asset ${asset.path || index} FastCDC chunks are missing`);
      }
      for (const [chunkIndex, chunk] of fastcdc.chunks.entries()) {
        addExact(
          keys,
          `blobs/sha256/${requireDigest(chunk?.sha256, `asset ${index} chunk ${chunkIndex}.sha256`)}`,
          `asset ${index} chunk ${chunkIndex}`,
        );
      }
    }
  }
  return releaseId;
}

function inventoryEntries(inventory) {
  if (!Array.isArray(inventory)) fail('R2 inventory must be a JSON array');
  const seen = new Set();
  return inventory.flatMap((entry, index) => {
    const rawKey = entry?.Path;
    // Some S3-compatible list implementations emit an empty root marker. The
    // marker is not a deletable object; keep dot paths in strict validation so
    // a real `.`/`./` object can never be silently omitted from cleanup.
    if (rawKey === '') return [];
    const normalizedKey = typeof rawKey === 'string' ? rawKey.replace(/^\.\//, '') : rawKey;
    const key = requireSafeKey(normalizedKey, `inventory entry ${index}.Path`);
    if (seen.has(key)) fail(`R2 inventory contains duplicate key ${key}`);
    seen.add(key);
    if (!Number.isSafeInteger(entry?.Size) || entry.Size < 0) fail(`inventory entry ${index}.Size is invalid`);
    return [{ key, bytes: entry.Size }];
  });
}

function sumBytes(entries) {
  return entries.reduce((total, entry) => total + entry.bytes, 0);
}

export function buildLatestOnlyCleanupPlan({ pointer, legacyPointer, manifest, manifestBytes, inventory }) {
  if (pointer?.version !== 1) fail('stable-v5 pointer version is invalid');
  const releaseId = requireReleaseId(pointer.releaseId, 'stable-v5.releaseId');
  if (pointer.manifestUrl !== `/releases/${releaseId}/manifest.json`) {
    fail('stable-v5 pointer is not bound to the v5 release manifest');
  }
  if (
    legacyPointer &&
    (legacyPointer.version !== 1 ||
      legacyPointer.releaseId !== releaseId ||
      legacyPointer.manifestUrl !== `/releases/${releaseId}/manifest-v4.json`)
  ) {
    fail('legacy stable pointer is not bound to the latest v5 release');
  }
  const manifestReleaseId = addManifestObjects(new Set(), manifest);
  if (manifestReleaseId !== releaseId) fail('stable-v5 pointer and manifest release IDs differ');
  const expectedManifestSha256 = crypto
    .createHash('sha256')
    .update(manifestBytes || Buffer.from(`${JSON.stringify(manifest)}\n`))
    .digest('hex');
  if (pointer.manifestSha256 !== expectedManifestSha256) {
    fail('stable-v5 pointer digest does not match the supplied manifest bytes');
  }
  const keepExact = new Set(['channels/stable.json', 'channels/stable-v5.json']);
  addExact(keepExact, `releases/${releaseId}/manifest.json`, 'latest manifest');
  addExact(keepExact, `releases/${releaseId}/manifest-v4.json`, 'latest compatibility manifest');
  addManifestObjects(keepExact, manifest);
  const keepPrefixes = [
    `releases/${releaseId}/`,
    `promotions/${releaseId}/`,
    `rollbacks/${releaseId}/`,
    `promotion-intents/${releaseId}/`,
  ];
  const entries = inventoryEntries(inventory);
  const kept = entries.filter(
    (entry) => keepExact.has(entry.key) || keepPrefixes.some((prefix) => entry.key.startsWith(prefix)),
  );
  const deleted = entries.filter(
    (entry) => !keepExact.has(entry.key) && !keepPrefixes.some((prefix) => entry.key.startsWith(prefix)),
  );
  const inventoryKeys = new Set(entries.map((entry) => entry.key));
  const missing = [...keepExact].filter((key) => !inventoryKeys.has(key));
  if (missing.length > 0) fail(`latest release references missing R2 objects: ${missing.join(', ')}`);
  return {
    version: 1,
    mode: 'latest-only',
    releaseId,
    inventoryObjects: entries.length,
    inventoryBytes: sumBytes(entries),
    retainedObjects: kept.length,
    retainedBytes: sumBytes(kept),
    deleteObjects: deleted.length,
    deleteBytes: sumBytes(deleted),
    retainedPrefixes: keepPrefixes,
    retainedKeys: [...keepExact].sort(),
    deleteKeysSha256: crypto
      .createHash('sha256')
      .update(
        deleted
          .map((entry) => entry.key)
          .sort()
          .join('\n'),
      )
      .digest('hex'),
  };
}

function option(name) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const pointerPath = option('--pointer');
  const legacyPointerPath = option('--legacy-pointer');
  const manifestPath = option('--manifest');
  const inventoryPath = option('--inventory');
  const outputPath = option('--output');
  const deleteListPath = option('--delete-list');
  if (!pointerPath || !manifestPath || !inventoryPath || !outputPath || !deleteListPath) {
    fail('--pointer, --manifest, --inventory, --output and --delete-list are required');
  }
  const pointer = readJson(pointerPath, 'stable-v5 pointer');
  const legacyPointer = legacyPointerPath ? readJson(legacyPointerPath, 'stable pointer') : undefined;
  const manifestBytes = fs.readFileSync(manifestPath);
  const manifest = JSON.parse(manifestBytes.toString('utf8'));
  const inventory = inventoryEntries(readJson(inventoryPath, 'R2 inventory'));
  const plan = buildLatestOnlyCleanupPlan({ pointer, legacyPointer, manifest, manifestBytes, inventory });
  const keep = new Set(plan.retainedKeys);
  const keepPrefixes = plan.retainedPrefixes;
  const deleteKeys = inventory
    .filter((entry) => !keep.has(entry.key) && !keepPrefixes.some((prefix) => entry.key.startsWith(prefix)))
    .map((entry) => entry.key)
    .sort();
  fs.mkdirSync(path.dirname(path.resolve(outputPath)), { recursive: true });
  fs.writeFileSync(path.resolve(outputPath), `${JSON.stringify(plan, null, 2)}\n`);
  fs.writeFileSync(path.resolve(deleteListPath), deleteKeys.length ? `${deleteKeys.join('\n')}\n` : '');
  console.log(
    `Latest-only cleanup plan for ${plan.releaseId}: retain ${plan.retainedObjects} objects (${plan.retainedBytes} bytes), delete ${plan.deleteObjects} objects (${plan.deleteBytes} bytes)`,
  );
}
