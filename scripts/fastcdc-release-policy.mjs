#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const SCRIPT_DIRECTORY = path.dirname(fileURLToPath(import.meta.url));
const FASTCDC_INDEXER_MANIFEST = path.resolve(SCRIPT_DIRECTORY, '../tools/fastcdc-index/Cargo.toml');
const DIGEST_PATTERN = /^[a-f0-9]{64}$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;

export const FASTCDC_RELEASE_POLICY = Object.freeze({
  id: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0',
  algorithm: 'fastcdc-v2020',
  minBytes: 64 * 1024,
  averageBytes: 256 * 1024,
  maxBytes: 1024 * 1024,
  normalization: 1,
  seed: 0,
  // Chrome can fail Cache.put() with NetworkError when a large network
  // ReadableStream is persisted as one response. Keep every canonical write
  // comfortably below that observed boundary by content-defining assets from
  // 8 MiB upward into <= 1 MiB immutable objects.
  minimumAssetBytes: 8 * 1024 * 1024,
  minimumHistoricalSamples: 2,
  minimumSavingsBytes: 1024 * 1024,
  minimumSavingsRatio: 0.25,
});

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function isNonNegativeSafeInteger(value) {
  return Number.isSafeInteger(value) && value >= 0;
}

function isSafeAssetPath(value) {
  return (
    typeof value === 'string' &&
    value.length > 0 &&
    !value.startsWith('/') &&
    !value.includes('\\') &&
    !value.includes('?') &&
    !value.includes('#') &&
    !value.split('/').some((segment) => !segment || segment === '.' || segment === '..')
  );
}

function parseEvidenceSample(value) {
  if (
    !value ||
    typeof value !== 'object' ||
    !RELEASE_ID_PATTERN.test(value.fromReleaseId || '') ||
    !RELEASE_ID_PATTERN.test(value.toReleaseId || '') ||
    value.fromReleaseId === value.toReleaseId ||
    !Number.isSafeInteger(value.wholeDownloadBytes) ||
    value.wholeDownloadBytes <= 0 ||
    !isNonNegativeSafeInteger(value.fastcdcDownloadBytes)
  ) {
    throw new Error('Invalid FastCDC evidence sample');
  }
  return {
    fromReleaseId: value.fromReleaseId,
    toReleaseId: value.toReleaseId,
    wholeDownloadBytes: value.wholeDownloadBytes,
    fastcdcDownloadBytes: value.fastcdcDownloadBytes,
  };
}

export function parseFastCdcEvidence(value) {
  if (!value || typeof value !== 'object' || value.version !== 1 || !Array.isArray(value.assets)) {
    throw new Error('Invalid FastCDC evidence document');
  }
  const paths = new Set();
  const assets = value.assets.map((asset) => {
    if (!asset || typeof asset !== 'object' || !isSafeAssetPath(asset.path) || !Array.isArray(asset.samples)) {
      throw new Error('Invalid FastCDC evidence asset');
    }
    if (paths.has(asset.path)) throw new Error(`Duplicate FastCDC evidence asset: ${asset.path}`);
    paths.add(asset.path);
    const transitions = new Set();
    const samples = asset.samples.map((sample) => {
      const parsed = parseEvidenceSample(sample);
      const transition = `${parsed.fromReleaseId}\0${parsed.toReleaseId}`;
      if (transitions.has(transition)) {
        throw new Error(`Duplicate FastCDC evidence transition for ${asset.path}`);
      }
      transitions.add(transition);
      return parsed;
    });
    return { path: asset.path, samples };
  });
  return { version: 1, assets };
}

export function readFastCdcEvidence(evidencePath) {
  let value;
  try {
    value = JSON.parse(fs.readFileSync(evidencePath, 'utf8'));
  } catch (error) {
    throw new Error(`Failed to read FastCDC evidence ${evidencePath}`, { cause: error });
  }
  return parseFastCdcEvidence(value);
}

export function evaluateFastCdcReleasePolicy({ assetPath, assetBytes, evidence }) {
  if (!isSafeAssetPath(assetPath) || !isNonNegativeSafeInteger(assetBytes)) {
    throw new TypeError('Invalid FastCDC policy input');
  }
  const parsedEvidence = parseFastCdcEvidence(evidence);
  const record = parsedEvidence.assets.find((asset) => asset.path === assetPath);
  const samples = record?.samples || [];
  const wholeDownloadBytes = samples.reduce((total, sample) => total + sample.wholeDownloadBytes, 0);
  const fastcdcDownloadBytes = samples.reduce((total, sample) => total + sample.fastcdcDownloadBytes, 0);
  const savingsBytes = wholeDownloadBytes - fastcdcDownloadBytes;
  const savingsRatio = wholeDownloadBytes > 0 ? savingsBytes / wholeDownloadBytes : 0;

  let reason = 'bounded-cache-write';
  if (assetBytes < FASTCDC_RELEASE_POLICY.minimumAssetBytes) reason = 'asset-too-small';

  return {
    selected: reason === 'bounded-cache-write',
    reason,
    samples: samples.length,
    wholeDownloadBytes,
    fastcdcDownloadBytes,
    savingsBytes,
    savingsRatio,
    policyId: FASTCDC_RELEASE_POLICY.id,
  };
}

export function runFastCdcIndexer(inputPath, { binaryPath, spawnSyncImpl = spawnSync, cargoTargetDirectory } = {}) {
  const command = binaryPath || 'cargo';
  const arguments_ = binaryPath
    ? [inputPath, String(FASTCDC_RELEASE_POLICY.averageBytes)]
    : [
        'run',
        '--quiet',
        '--release',
        '--manifest-path',
        FASTCDC_INDEXER_MANIFEST,
        '--',
        inputPath,
        String(FASTCDC_RELEASE_POLICY.averageBytes),
      ];
  const result = spawnSyncImpl(command, arguments_, {
    encoding: 'utf8',
    maxBuffer: 64 * 1024 * 1024,
    env: {
      ...process.env,
      ...(cargoTargetDirectory ? { CARGO_TARGET_DIR: cargoTargetDirectory } : {}),
    },
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    throw new Error(result.stderr || `FastCDC indexer exited with ${result.status}`);
  }
  try {
    return JSON.parse(result.stdout);
  } catch (error) {
    throw new Error('FastCDC indexer returned invalid JSON', { cause: error });
  }
}

function selectFixedConfiguration(index, expectedBytes) {
  if (!index || typeof index !== 'object' || index.bytes !== expectedBytes || !Array.isArray(index.configurations)) {
    throw new Error('FastCDC index does not match the input file');
  }
  const configuration = index.configurations.find(
    (candidate) =>
      candidate?.minimumBytes === FASTCDC_RELEASE_POLICY.minBytes &&
      candidate.averageBytes === FASTCDC_RELEASE_POLICY.averageBytes &&
      candidate.maximumBytes === FASTCDC_RELEASE_POLICY.maxBytes,
  );
  if (!configuration || !Array.isArray(configuration.chunks) || configuration.chunks.length === 0) {
    throw new Error('FastCDC index is missing the fixed release policy');
  }
  return configuration;
}

function validateChunkIndex(configuration, expectedBytes) {
  let expectedOffset = 0;
  return configuration.chunks.map((chunk, index) => {
    if (
      !chunk ||
      chunk.offset !== expectedOffset ||
      !Number.isSafeInteger(chunk.bytes) ||
      chunk.bytes <= 0 ||
      chunk.bytes > FASTCDC_RELEASE_POLICY.maxBytes ||
      (index < configuration.chunks.length - 1 && chunk.bytes < FASTCDC_RELEASE_POLICY.minBytes) ||
      !DIGEST_PATTERN.test(chunk.sha256 || '')
    ) {
      throw new Error(`Invalid FastCDC chunk at index ${index}`);
    }
    expectedOffset += chunk.bytes;
    if (expectedOffset > expectedBytes) throw new Error('FastCDC chunks exceed the input file');
    return {
      offset: chunk.offset,
      bytes: chunk.bytes,
      sha256: chunk.sha256,
    };
  });
}

function readExact(file, offset, bytes) {
  const output = Buffer.allocUnsafe(bytes);
  let read = 0;
  while (read < bytes) {
    const count = fs.readSync(file, output, read, bytes - read, offset + read);
    if (count === 0) throw new Error('FastCDC index extends beyond the input file');
    read += count;
  }
  return output;
}

function verifyExistingBlob(blobPath, expectedSha256, expectedBytes) {
  const bytes = fs.readFileSync(blobPath);
  if (bytes.byteLength !== expectedBytes || sha256(bytes) !== expectedSha256) {
    throw new Error(`Conflicting content-addressed blob: ${expectedSha256}`);
  }
}

function persistBlob(blobDirectory, digest, bytes, sequence) {
  const blobPath = path.join(blobDirectory, digest);
  if (fs.existsSync(blobPath)) {
    verifyExistingBlob(blobPath, digest, bytes.byteLength);
    return false;
  }
  const temporaryPath = path.join(blobDirectory, `.${digest}.${process.pid}.${sequence}.tmp`);
  fs.writeFileSync(temporaryPath, bytes, { flag: 'wx' });
  try {
    try {
      fs.linkSync(temporaryPath, blobPath);
      return true;
    } catch (error) {
      if (error?.code !== 'EEXIST') throw error;
      verifyExistingBlob(blobPath, digest, bytes.byteLength);
      return false;
    }
  } finally {
    fs.unlinkSync(temporaryPath);
  }
}

export function buildFastCdcRepresentation({
  inputPath,
  output,
  expectedBytes,
  expectedSha256,
  indexer = runFastCdcIndexer,
}) {
  if (!Number.isSafeInteger(expectedBytes) || expectedBytes <= 0 || !DIGEST_PATTERN.test(expectedSha256 || '')) {
    throw new TypeError('Invalid expected FastCDC asset identity');
  }
  const stat = fs.statSync(inputPath);
  if (!stat.isFile() || stat.size !== expectedBytes) {
    throw new Error('FastCDC input does not match the expected byte length');
  }

  const index = indexer(inputPath, FASTCDC_RELEASE_POLICY);
  const configuration = selectFixedConfiguration(index, expectedBytes);
  const chunks = validateChunkIndex(configuration, expectedBytes);
  const coveredBytes = chunks.reduce((total, chunk) => total + chunk.bytes, 0);
  if (coveredBytes !== expectedBytes) throw new Error('FastCDC chunks do not cover the complete input file');

  const file = fs.openSync(inputPath, 'r');
  try {
    const completeDigest = crypto.createHash('sha256');
    for (const [index_, chunk] of chunks.entries()) {
      const bytes = readExact(file, chunk.offset, chunk.bytes);
      if (sha256(bytes) !== chunk.sha256) {
        throw new Error(`FastCDC chunk digest mismatch at index ${index_}`);
      }
      completeDigest.update(bytes);
    }
    if (completeDigest.digest('hex') !== expectedSha256) {
      throw new Error('FastCDC reconstructed file digest mismatch');
    }

    const blobDirectory = path.join(output, 'blobs', 'sha256');
    fs.mkdirSync(blobDirectory, { recursive: true });
    const written = new Set();
    for (const [index_, chunk] of chunks.entries()) {
      if (written.has(chunk.sha256)) continue;
      const bytes = readExact(file, chunk.offset, chunk.bytes);
      persistBlob(blobDirectory, chunk.sha256, bytes, index_);
      written.add(chunk.sha256);
    }
  } finally {
    fs.closeSync(file);
  }

  return {
    algorithm: FASTCDC_RELEASE_POLICY.algorithm,
    minBytes: FASTCDC_RELEASE_POLICY.minBytes,
    averageBytes: FASTCDC_RELEASE_POLICY.averageBytes,
    maxBytes: FASTCDC_RELEASE_POLICY.maxBytes,
    normalization: FASTCDC_RELEASE_POLICY.normalization,
    seed: FASTCDC_RELEASE_POLICY.seed,
    chunks,
  };
}

export function buildFastCdcRepresentationFromEvidence({
  assetPath,
  inputPath,
  output,
  expectedBytes,
  expectedSha256,
  evidence,
  evidencePath,
  indexer,
}) {
  const parsedEvidence = evidence ? parseFastCdcEvidence(evidence) : readFastCdcEvidence(evidencePath);
  const decision = evaluateFastCdcReleasePolicy({
    assetPath,
    assetBytes: expectedBytes,
    evidence: parsedEvidence,
  });
  if (!decision.selected) return { selected: false, decision };
  return {
    selected: true,
    decision,
    representation: buildFastCdcRepresentation({
      inputPath,
      output,
      expectedBytes,
      expectedSha256,
      ...(indexer ? { indexer } : {}),
    }),
  };
}
