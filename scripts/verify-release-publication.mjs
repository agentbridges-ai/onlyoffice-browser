#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { computeStorageSetSha256 } from './build-release-manifest.mjs';
import {
  FASTCDC_RELEASE_POLICY,
  evaluateFastCdcReleasePolicy,
  readFastCdcEvidence,
} from './fastcdc-release-policy.mjs';

const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const GIT_COMMIT_PATTERN = /^[a-f0-9]{40}$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const FASTCDC_POLICY_ID = 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0';
const FASTCDC_CONFIGURATION = Object.freeze({
  algorithm: 'fastcdc-v2020',
  minBytes: 64 * 1024,
  averageBytes: 256 * 1024,
  maxBytes: 1024 * 1024,
  normalization: 1,
  seed: 0,
});

function fail(message) {
  throw new Error(message);
}

function isPositiveSafeInteger(value) {
  return Number.isSafeInteger(value) && value > 0;
}

function isNonNegativeSafeInteger(value) {
  return Number.isSafeInteger(value) && value >= 0;
}

function assertSha256(value, label) {
  if (typeof value !== 'string' || !SHA256_PATTERN.test(value)) {
    fail(`${label} must be a lowercase SHA-256 digest`);
  }
  return value;
}

function assertReleaseId(value, label) {
  if (typeof value !== 'string' || !RELEASE_ID_PATTERN.test(value)) {
    fail(`${label} is invalid`);
  }
  return value;
}

function assertSafeRelativePath(value, label) {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.startsWith('/') ||
    value.includes('\\') ||
    value.includes('?') ||
    value.includes('#') ||
    [...value].some((character) => {
      const codePoint = character.codePointAt(0);
      return codePoint !== undefined && (codePoint <= 31 || codePoint === 127);
    }) ||
    value.split('/').some((segment) => !segment || segment === '.' || segment === '..')
  ) {
    fail(`${label} is not a safe relative path`);
  }
  return value;
}

function readJsonWithBytes(file, label) {
  const bytes = fs.readFileSync(file);
  try {
    return { bytes, value: JSON.parse(bytes.toString('utf8')) };
  } catch (error) {
    throw new Error(`${label} is not valid JSON`, { cause: error });
  }
}

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function stableJson(value) {
  return JSON.stringify(value);
}

function assertSame(left, right, label) {
  if (stableJson(left) !== stableJson(right)) fail(`${label} does not match between v5 and v4`);
}

function addObject(objects, object) {
  const existing = objects.get(object.key);
  if (existing) {
    if (existing.bytes !== object.bytes || existing.sha256 !== object.sha256) {
      fail(`Conflicting immutable object identity for ${object.key}`);
    }
    return;
  }
  objects.set(object.key, object);
}

function validatePointer(pointer, { label, releaseId, manifestPath, manifestBytes }) {
  if (
    !pointer ||
    typeof pointer !== 'object' ||
    pointer.version !== 1 ||
    pointer.releaseId !== releaseId ||
    pointer.manifestUrl !== `/${manifestPath}`
  ) {
    fail(`${label} does not point to ${manifestPath}`);
  }
  if (assertSha256(pointer.manifestSha256, `${label}.manifestSha256`) !== sha256(manifestBytes)) {
    fail(`${label} manifest digest does not match ${manifestPath}`);
  }
}

function validatePackage(manifest, objects) {
  const pack = manifest.package;
  if (
    !pack ||
    pack.format !== 'onlyoffice-pack-v1' ||
    pack.path !== 'office-resources.oobpack' ||
    !isPositiveSafeInteger(pack.bytes) ||
    !Array.isArray(pack.segments) ||
    pack.segments.length === 0
  ) {
    fail('v5 package descriptor is invalid');
  }
  assertSha256(pack.sha256, 'v5 package.sha256');
  let expectedOffset = 0;
  for (const [index, segment] of pack.segments.entries()) {
    const digest = assertSha256(segment?.sha256, `v5 package segment ${index}.sha256`);
    if (segment?.id !== digest || segment?.offset !== expectedOffset || !isPositiveSafeInteger(segment?.bytes)) {
      fail(`v5 package segment ${index} is invalid`);
    }
    addObject(objects, {
      kind: 'package-segment',
      key: `segments/sha256/${digest}`,
      bytes: segment.bytes,
      sha256: digest,
    });
    expectedOffset += segment.bytes;
  }
  if (expectedOffset !== pack.bytes) fail('v5 package segments do not cover the complete package');
  addObject(objects, {
    kind: 'package',
    key: `packages/sha256/${pack.sha256}.oobpack`,
    bytes: pack.bytes,
    sha256: pack.sha256,
  });
}

function validateFastCdc(asset, fastcdc, objects) {
  for (const [key, value] of Object.entries(FASTCDC_CONFIGURATION)) {
    if (fastcdc?.[key] !== value) fail(`FastCDC configuration mismatch for ${asset.path}`);
  }
  if (!Array.isArray(fastcdc.chunks) || fastcdc.chunks.length === 0) {
    fail(`FastCDC chunks are missing for ${asset.path}`);
  }
  let expectedOffset = 0;
  for (const [index, chunk] of fastcdc.chunks.entries()) {
    const digest = assertSha256(chunk?.sha256, `FastCDC ${asset.path} chunk ${index}.sha256`);
    if (chunk?.offset !== expectedOffset || !isPositiveSafeInteger(chunk?.bytes)) {
      fail(`FastCDC ${asset.path} chunk ${index} is invalid`);
    }
    expectedOffset += chunk.bytes;
    if (expectedOffset > asset.bytes) fail(`FastCDC chunks exceed ${asset.path}`);
    addObject(objects, {
      kind: 'fastcdc',
      key: `blobs/sha256/${digest}`,
      bytes: chunk.bytes,
      sha256: digest,
    });
  }
  if (expectedOffset !== asset.bytes) fail(`FastCDC chunks do not cover ${asset.path}`);
}

function validateAssets(manifest, objects) {
  if (!Array.isArray(manifest.assets) || manifest.assets.length === 0) {
    fail('v5 assets are missing');
  }
  const paths = new Set();
  const fastCdcAssets = [];
  for (const [index, asset] of manifest.assets.entries()) {
    const assetPath = assertSafeRelativePath(asset?.path, `v5 asset ${index}.path`);
    if (paths.has(assetPath)) fail(`Duplicate v5 asset path: ${assetPath}`);
    paths.add(assetPath);
    if (
      !isNonNegativeSafeInteger(asset?.bytes) ||
      typeof asset?.mime !== 'string' ||
      asset.mime.length === 0 ||
      !isNonNegativeSafeInteger(asset?.packageOffset) ||
      asset.packageOffset + asset.bytes > manifest.package.bytes
    ) {
      fail(`v5 asset ${assetPath} metadata is invalid`);
    }
    const digest = assertSha256(asset.sha256, `v5 asset ${assetPath}.sha256`);
    const whole = asset.representations?.whole;
    if (!whole || whole.sha256 !== digest || whole.bytes !== asset.bytes) {
      fail(`v5 whole representation does not match ${assetPath}`);
    }
    addObject(objects, {
      kind: 'whole',
      key: `blobs/sha256/${digest}`,
      bytes: asset.bytes,
      sha256: digest,
    });
    if (asset.bytes >= FASTCDC_RELEASE_POLICY.minimumAssetBytes && !asset.representations.fastcdc) {
      fail(`Large asset ${assetPath} is missing its bounded FastCDC representation`);
    }
    if (asset.bytes < FASTCDC_RELEASE_POLICY.minimumAssetBytes && asset.representations.fastcdc) {
      fail(`Small asset ${assetPath} must remain a whole-file CAS object`);
    }
    if (asset.representations.fastcdc) {
      validateFastCdc(asset, asset.representations.fastcdc, objects);
      fastCdcAssets.push(asset);
    }
  }
  return fastCdcAssets;
}

function validateFastCdcEvidence(manifest, fastCdcAssets, options) {
  const evidenceMode = options.fastCdcEvidenceMode || 'automatic';
  if (evidenceMode !== 'automatic' && evidenceMode !== 'forbid' && evidenceMode !== 'required') {
    fail(`Unknown FastCDC evidence mode: ${evidenceMode}`);
  }
  if (evidenceMode === 'automatic') {
    if (options.fastCdcEvidencePath) fail('FastCDC evidence cannot be passed in automatic mode');
    if (manifest.contentProtocol.fastcdcPolicyId !== FASTCDC_POLICY_ID) {
      fail('v5 FastCDC policy identity does not match the automatic bounded-write policy');
    }
    return null;
  }
  if (evidenceMode === 'forbid') {
    if (options.fastCdcEvidencePath) fail('FastCDC evidence cannot be passed in forbid mode');
    if (fastCdcAssets.length > 0) {
      fail('FastCDC objects are forbidden without authoritative two-transition evidence');
    }
    return null;
  }
  if (!options.fastCdcEvidencePath) {
    fail('FastCDC evidence mode required needs --fastcdc-evidence');
  }
  const evidence = readFastCdcEvidence(options.fastCdcEvidencePath);
  if (fastCdcAssets.length === 0) {
    fail('FastCDC evidence was required but the v5 manifest selected no FastCDC assets');
  }
  for (const asset of fastCdcAssets) {
    const decision = evaluateFastCdcReleasePolicy({
      assetPath: asset.path,
      assetBytes: asset.bytes,
      evidence,
    });
    if (!decision.selected || decision.samples < 2) {
      fail(`FastCDC asset ${asset.path} lacks qualifying two-transition evidence`);
    }
  }
  if (manifest.contentProtocol.fastcdcPolicyId !== FASTCDC_POLICY_ID) {
    fail('v5 FastCDC evidence does not match the declared policy');
  }
  return evidence;
}

function collectV5Objects(manifest, releaseId, options = {}) {
  if (manifest?.version !== 5 || manifest.releaseId !== releaseId) {
    fail('manifest.json must be Release Manifest v5 for the requested release');
  }
  if (typeof manifest.packageVersion !== 'string' || manifest.packageVersion.length === 0) {
    fail('v5 packageVersion identity is invalid');
  }
  if (options.expectedPackageVersion && manifest.packageVersion !== options.expectedPackageVersion) {
    fail(`Expected package version ${options.expectedPackageVersion}, received ${manifest.packageVersion}`);
  }
  if (options.expectedSourceCommit) {
    if (!GIT_COMMIT_PATTERN.test(options.expectedSourceCommit)) fail('expected source commit is invalid');
    if (manifest.sourceCommit !== options.expectedSourceCommit) {
      fail('v5 sourceCommit does not match the expected candidate commit');
    }
    const escapedPackageVersion = manifest.packageVersion.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    if (!new RegExp(`^office-host-${escapedPackageVersion}-r[1-9][0-9]*$`).test(manifest.protocolHostBuildId || '')) {
      fail('v5 protocolHostBuildId identity is invalid');
    }
  } else if (manifest.sourceCommit !== undefined && !GIT_COMMIT_PATTERN.test(manifest.sourceCommit)) {
    fail('v5 sourceCommit identity is invalid');
  }
  if (
    manifest.contentProtocol?.version !== 1 ||
    manifest.contentProtocol.digest !== 'sha256' ||
    manifest.contentProtocol.cacheKeyFormat !== 'canonical-sha256-v1' ||
    assertSha256(manifest.contentProtocol.storageSetSha256, 'contentProtocol.storageSetSha256') !==
      computeStorageSetSha256(manifest.package, manifest.assets) ||
    manifest.contentProtocol.fastcdcPolicyId !== FASTCDC_POLICY_ID
  ) {
    fail('v5 content protocol identity is invalid');
  }
  const objects = new Map();
  validatePackage(manifest, objects);
  const fastCdcAssets = validateAssets(manifest, objects);
  validateFastCdcEvidence(manifest, fastCdcAssets, options);
  return { objects, fastCdcAssets };
}

export function loadV5ManifestPublication(manifestFile, options = {}) {
  const file = path.resolve(manifestFile);
  const manifestV5 = readJsonWithBytes(file, 'v5 manifest');
  const releaseId = assertReleaseId(options.releaseId || manifestV5.value?.releaseId, 'releaseId');
  const { objects, fastCdcAssets } = collectV5Objects(manifestV5.value, releaseId, options);
  const manifestSha256 = sha256(manifestV5.bytes);
  if (options.expectedManifestSha256) {
    const expected = assertSha256(options.expectedManifestSha256, 'expected manifest digest');
    if (manifestSha256 !== expected) fail('v5 manifest digest does not match the expected pointer digest');
  }
  addObject(objects, {
    kind: 'manifest-v5',
    key: `releases/${releaseId}/manifest.json`,
    bytes: manifestV5.bytes.byteLength,
    sha256: manifestSha256,
  });
  return {
    root: path.dirname(file),
    releaseId,
    manifest: manifestV5.value,
    objects: [...objects.values()].sort((left, right) => left.key.localeCompare(right.key)),
    fastCdcAssets: fastCdcAssets.map((asset) => asset.path),
  };
}

export function loadReleasePublication(releaseRoot, options = {}) {
  const root = path.resolve(releaseRoot);
  const stableV5 = readJsonWithBytes(path.join(root, 'channels/stable-v5.json'), 'stable-v5.json');
  const releaseId = assertReleaseId(stableV5.value?.releaseId, 'stable-v5.releaseId');
  const v5ManifestPath = `releases/${releaseId}/manifest.json`;
  const compatibilityManifestPath = `releases/${releaseId}/manifest-v4.json`;
  const manifestV5 = readJsonWithBytes(path.join(root, v5ManifestPath), 'v5 manifest');
  const stableV4 = readJsonWithBytes(path.join(root, 'channels/stable.json'), 'stable.json');
  const manifestV4 = readJsonWithBytes(path.join(root, compatibilityManifestPath), 'v4 manifest');

  validatePointer(stableV5.value, {
    label: 'stable-v5.json',
    releaseId,
    manifestPath: v5ManifestPath,
    manifestBytes: manifestV5.bytes,
  });
  validatePointer(stableV4.value, {
    label: 'stable.json',
    releaseId,
    manifestPath: compatibilityManifestPath,
    manifestBytes: manifestV4.bytes,
  });

  const manifest = manifestV5.value;
  const compatibilityManifest = manifestV4.value;
  if (compatibilityManifest?.version !== 4 || compatibilityManifest.releaseId !== releaseId) {
    fail('manifest-v4.json must retain Release Manifest v4 compatibility');
  }
  if (compatibilityManifest.packageVersion !== manifest.packageVersion) {
    fail('v5/v4 packageVersion identity is invalid');
  }
  const { objects, fastCdcAssets } = collectV5Objects(manifest, releaseId, options);

  for (const field of [
    'hostBuildId',
    'shellRevision',
    'runtimeManifestSha256',
    'fontManifestSha256',
    'x2t',
    'profiles',
    'chunks',
    'package',
    'fontFamilies',
  ]) {
    assertSame(manifest[field], compatibilityManifest[field], field);
  }
  const v5CompatibilityAssets = manifest.assets.map(({ representations: _representations, ...asset }) => asset);
  assertSame(v5CompatibilityAssets, compatibilityManifest.assets, 'assets');

  addObject(objects, {
    kind: 'manifest-v5',
    key: v5ManifestPath,
    bytes: manifestV5.bytes.byteLength,
    sha256: stableV5.value.manifestSha256,
  });
  addObject(objects, {
    kind: 'manifest-v4',
    key: compatibilityManifestPath,
    bytes: manifestV4.bytes.byteLength,
    sha256: stableV4.value.manifestSha256,
  });

  return {
    root,
    releaseId,
    manifest,
    compatibilityManifest,
    stableV5: stableV5.value,
    stableV4: stableV4.value,
    objects: [...objects.values()].sort((left, right) => left.key.localeCompare(right.key)),
    fastCdcAssets: fastCdcAssets.map((asset) => asset.path),
  };
}

function localObjectPath(publication, object) {
  return path.join(publication.root, ...object.key.split('/'));
}

export async function hashFile(file) {
  const digest = crypto.createHash('sha256');
  let bytes = 0;
  for await (const chunk of fs.createReadStream(file)) {
    digest.update(chunk);
    bytes += chunk.byteLength;
  }
  return { bytes, sha256: digest.digest('hex') };
}

export async function verifyObjects(objects, inspect, { concurrency = 4, label = 'object' } = {}) {
  if (!Number.isSafeInteger(concurrency) || concurrency <= 0 || concurrency > 32) {
    throw new TypeError('Verification concurrency must be between 1 and 32');
  }
  let cursor = 0;
  let verifiedBytes = 0;
  async function worker() {
    while (true) {
      const index = cursor;
      cursor += 1;
      if (index >= objects.length) return;
      const object = objects[index];
      const actual = await inspect(object);
      if (actual.bytes !== object.bytes || actual.sha256 !== object.sha256) {
        fail(
          `${label} verification failed for ${object.key}: expected ${object.bytes}/${object.sha256}, received ${actual.bytes}/${actual.sha256}`,
        );
      }
      verifiedBytes += actual.bytes;
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, objects.length) }, () => worker()));
  return { objects: objects.length, bytes: verifiedBytes };
}

export async function verifyLocalRelease(publication, options = {}) {
  return verifyObjects(publication.objects, (object) => hashFile(localObjectPath(publication, object)), {
    ...options,
    label: 'Local immutable object',
  });
}

export function inspectRcloneObject(remote, object, { rcloneBinary = 'rclone', spawnImpl = spawn } = {}) {
  const target = `${remote.replace(/\/+$/, '')}/${object.key}`;
  return new Promise((resolve, reject) => {
    const child = spawnImpl(rcloneBinary, ['cat', target], {
      stdio: ['ignore', 'pipe', 'pipe'],
    });
    const digest = crypto.createHash('sha256');
    let bytes = 0;
    let stderr = '';
    child.stdout.on('data', (chunk) => {
      digest.update(chunk);
      bytes += chunk.byteLength;
    });
    child.stderr.on('data', (chunk) => {
      if (stderr.length < 16 * 1024) stderr += chunk.toString('utf8');
    });
    child.on('error', reject);
    child.on('close', (code, signal) => {
      if (code !== 0) {
        reject(
          new Error(`rclone cat failed for ${object.key} (${signal || code}): ${stderr.trim() || 'no diagnostic'}`),
        );
        return;
      }
      resolve({ bytes, sha256: digest.digest('hex') });
    });
  });
}

/**
 * Read the remote object index without transferring object bodies.  R2 object
 * names are immutable CAS identities, so the incremental verifier can use a
 * pre-upload inventory to distinguish objects that were already verified in a
 * previous publication from objects introduced by this release.
 */
export function inspectRcloneInventory(
  remote,
  { rcloneBinary = 'rclone', spawnImpl = spawn, timeoutMs = 15 * 60 * 1000 } = {},
) {
  if (typeof remote !== 'string' || !remote.includes(':')) throw new TypeError('Invalid rclone remote');
  const target = remote.replace(/\/+$/, '');
  return new Promise((resolve, reject) => {
    // `lsjson --fast-list` builds one large in-memory response and can stall
    // on a bucket containing years of immutable CAS objects. `lsf` with
    // ListR disabled streams one object per line, so the inventory remains
    // bounded and emits progress while R2 paginates.
    const child = spawnImpl(
      rcloneBinary,
      ['lsf', target, '--recursive', '--files-only', '--format', 'ps', '--separator', '\t', '--disable', 'ListR'],
      {
        stdio: ['ignore', 'pipe', 'pipe'],
      },
    );
    const inventory = new Map();
    let stderr = '';
    let pending = '';
    let listed = 0;
    let settled = false;
    const fail = (error) => {
      if (settled) return;
      settled = true;
      reject(error);
    };
    const timeout =
      Number.isFinite(timeoutMs) && timeoutMs > 0
        ? setTimeout(() => {
            child.kill?.('SIGTERM');
            fail(new Error(`rclone lsf inventory timed out after ${timeoutMs}ms`));
          }, timeoutMs)
        : null;
    timeout?.unref?.();
    const consumeLine = (line) => {
      if (!line) return;
      const separator = line.lastIndexOf('\t');
      const key = separator >= 0 ? line.slice(0, separator) : '';
      const rawSize = separator >= 0 ? line.slice(separator + 1) : '';
      if (!key || !rawSize || key.endsWith('/')) return;
      const size = Number(rawSize);
      if (!Number.isSafeInteger(size) || size < 0) {
        fail(new Error(`rclone lsf returned an invalid size for ${key}`));
        return;
      }
      inventory.set(key.replace(/^\/+/, ''), { bytes: size });
      listed += 1;
      if (listed % 1000 === 0) console.log(`R2 inventory: ${listed} objects`);
    };
    child.stdout.on('data', (chunk) => {
      if (settled) return;
      pending += chunk.toString('utf8');
      const lines = pending.split('\n');
      pending = lines.pop() || '';
      for (const line of lines) consumeLine(line.replace(/\r$/, ''));
    });
    child.stderr.on('data', (chunk) => {
      if (stderr.length < 16 * 1024) stderr += chunk.toString('utf8');
    });
    child.on('error', (error) => fail(error));
    child.on('close', (code, signal) => {
      if (timeout) clearTimeout(timeout);
      if (settled) return;
      if (code !== 0) {
        fail(new Error(`rclone lsf failed (${signal || code}): ${stderr.trim() || 'no diagnostic'}`));
        return;
      }
      consumeLine(pending.replace(/\r$/, ''));
      settled = true;
      resolve(inventory);
    });
  });
}

/**
 * Validate the post-upload inventory and return only objects whose bodies
 * need a SHA-256 transfer.  Existing CAS objects are still checked for
 * presence and exact size; their bodies are audited by the periodic full mode.
 */
export function planIncrementalRemoteVerification(objects, beforeInventory, afterInventory) {
  const toHash = [];
  let reusedObjects = 0;
  let reusedBytes = 0;
  for (const object of objects) {
    const after = afterInventory.get(object.key);
    if (!after) fail(`Remote immutable object is missing from the inventory: ${object.key}`);
    if (after.bytes !== object.bytes) {
      fail(
        `Remote immutable object has the wrong size for ${object.key}: expected ${object.bytes}, received ${after.bytes}`,
      );
    }
    const before = beforeInventory.get(object.key);
    if (before && before.bytes === object.bytes) {
      reusedObjects += 1;
      reusedBytes += object.bytes;
    } else {
      toHash.push(object);
    }
  }
  return { toHash, reusedObjects, reusedBytes };
}

export async function verifyRemoteRelease(publication, remote, options = {}) {
  if (typeof remote !== 'string' || !remote.includes(':')) throw new TypeError('Invalid rclone remote');
  const mode = options.remoteVerificationMode || 'full';
  if (mode !== 'full' && mode !== 'incremental') {
    throw new TypeError(`Unknown remote verification mode: ${mode}`);
  }
  if (mode === 'full') {
    const result = await verifyObjects(publication.objects, (object) => inspectRcloneObject(remote, object, options), {
      concurrency: options.concurrency || 4,
      label: 'Remote immutable object',
    });
    return {
      mode,
      objects: result.objects,
      bytes: result.bytes,
      verifiedObjects: result.objects,
      verifiedBytes: result.bytes,
      reusedObjects: 0,
      reusedBytes: 0,
    };
  }
  if (typeof options.remoteInventoryPath !== 'string' || !options.remoteInventoryPath) {
    throw new TypeError('Incremental remote verification requires --remote-inventory');
  }
  const beforeBytes = fs.readFileSync(options.remoteInventoryPath, 'utf8');
  let beforeEntries;
  try {
    beforeEntries = JSON.parse(beforeBytes || '[]');
  } catch (error) {
    throw new Error(
      `Incremental remote inventory is invalid JSON: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
  if (!Array.isArray(beforeEntries)) throw new TypeError('Incremental remote inventory must be an array');
  const beforeInventory = new Map(
    beforeEntries
      .filter((entry) => typeof entry?.Path === 'string' && Number.isSafeInteger(entry?.Size) && entry.Size >= 0)
      .map((entry) => [entry.Path.replace(/^\/+/, ''), { bytes: entry.Size }]),
  );
  const afterInventory = await inspectRcloneInventory(remote, options);
  const plan = planIncrementalRemoteVerification(publication.objects, beforeInventory, afterInventory);
  const result = await verifyObjects(plan.toHash, (object) => inspectRcloneObject(remote, object, options), {
    concurrency: options.concurrency || 4,
    label: 'Remote immutable object',
  });
  return {
    mode,
    objects: publication.objects.length,
    bytes: result.bytes + plan.reusedBytes,
    verifiedObjects: result.objects,
    verifiedBytes: result.bytes,
    reusedObjects: plan.reusedObjects,
    reusedBytes: plan.reusedBytes,
  };
}

function option(name, fallback) {
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : fallback;
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  const remote = option('--remote');
  const concurrency = Number(option('--concurrency', '4'));
  const v5Manifest = option('--v5-manifest');
  const commonOptions = {
    expectedPackageVersion: option('--expected-package-version'),
    expectedSourceCommit: option('--expected-source-commit'),
    fastCdcEvidenceMode: option('--fastcdc-evidence-mode', 'automatic'),
    fastCdcEvidencePath: option('--fastcdc-evidence'),
  };
  const publication = v5Manifest
    ? loadV5ManifestPublication(v5Manifest, {
        ...commonOptions,
        releaseId: option('--release-id'),
        expectedManifestSha256: option('--expected-manifest-sha256'),
      })
    : loadReleasePublication(option('--release-root', '.onlyoffice-release'), commonOptions);
  if (!v5Manifest) {
    const local = await verifyLocalRelease(publication, { concurrency });
    console.log(
      `Verified ${local.objects} local immutable objects (${local.bytes} bytes) for ${publication.releaseId}`,
    );
  } else if (!remote) {
    fail('--v5-manifest requires --remote because its object bodies are not local');
  }
  if (remote) {
    const remoteResult = await verifyRemoteRelease(publication, remote, {
      concurrency,
      rcloneBinary: option('--rclone-bin', 'rclone'),
      remoteVerificationMode: option('--remote-verification-mode', 'full'),
      remoteInventoryPath: option('--remote-inventory'),
    });
    const reportPath = option('--report-file');
    if (reportPath) {
      fs.mkdirSync(path.dirname(path.resolve(reportPath)), { recursive: true });
      fs.writeFileSync(
        reportPath,
        `${JSON.stringify(
          {
            version: 1,
            releaseId: publication.releaseId,
            manifestVersion: publication.manifest?.version,
            ...remoteResult,
          },
          null,
          2,
        )}\n`,
      );
    }
    console.log(
      `Verified ${remoteResult.objects} remote immutable objects (${remoteResult.bytes} bytes) for ${publication.releaseId}` +
        (remoteResult.reusedObjects === undefined
          ? ''
          : `; hashed ${remoteResult.verifiedObjects} new/replaced objects, reused ${remoteResult.reusedObjects}`),
    );
  }
}
