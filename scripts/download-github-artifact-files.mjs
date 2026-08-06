#!/usr/bin/env node

/**
 * Read selected entries from a GitHub Actions artifact without downloading the
 * artifact ZIP. GitHub's immutable artifact archive is a regular ZIP served
 * by an Azure blob endpoint, so the central directory and each selected entry
 * can be fetched with HTTP Range requests. This is the hand-off boundary for
 * staging: metadata and changed CAS objects are fetched, while unchanged
 * objects never cross the runner network boundary.
 */

import crypto from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { Readable, Transform } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import { createInflateRaw } from 'node:zlib';
import { fileURLToPath } from 'node:url';

const ZIP_LOCAL_HEADER = 0x04034b50;
const ZIP_CENTRAL_HEADER = 0x02014b50;
const ZIP_END_HEADER = 0x06054b50;
const ZIP64_END_HEADER = 0x06064b50;
const ZIP64_LOCATOR_HEADER = 0x07064b50;
const ZIP64_EXTRA_FIELD = 0x0001;
const RANGE_CHUNK_BYTES = 32 * 1024 * 1024;
const MAX_END_SCAN_BYTES = 128 * 1024;

function fail(message) {
  throw new Error(message);
}

function authHeaders(token) {
  return token ? { authorization: `Bearer ${token}`, accept: 'application/vnd.github+json' } : {};
}

function parseContentRange(value) {
  const match = /^bytes\s+(\d+)-(\d+)\/(\d+)$/.exec(value || '');
  if (!match) fail(`Range response is missing a valid Content-Range header: ${value || '<missing>'}`);
  return { start: Number(match[1]), end: Number(match[2]), total: Number(match[3]) };
}

async function resolveArchiveUrl({
  archiveUrl,
  repository,
  runId,
  artifactId,
  artifactName,
  token,
  fetchImpl = fetch,
}) {
  let source = archiveUrl;
  if (!source) {
    if (!repository || (!artifactId && !runId) || (!artifactId && !artifactName)) {
      fail('artifact URL or repository + artifact ID, or repository + run ID + artifact name, is required');
    }
    let artifact;
    if (artifactId) {
      const response = await fetchImpl(
        `https://api.github.com/repos/${repository}/actions/artifacts/${encodeURIComponent(artifactId)}`,
        { headers: authHeaders(token) },
      );
      if (!response.ok) fail(`GitHub artifact lookup failed (${response.status})`);
      artifact = await response.json();
    } else {
      const response = await fetchImpl(
        `https://api.github.com/repos/${repository}/actions/runs/${encodeURIComponent(runId)}/artifacts?per_page=100`,
        { headers: authHeaders(token) },
      );
      if (!response.ok) fail(`GitHub artifact listing failed (${response.status})`);
      const listing = await response.json();
      artifact = listing.artifacts?.find((candidate) => candidate.name === artifactName && !candidate.expired);
      if (!artifact) fail(`GitHub artifact ${artifactName} was not found in run ${runId}`);
    }
    if (!artifact?.archive_download_url || artifact.expired) fail('GitHub artifact is missing or expired');
    source = artifact.archive_download_url;
  }

  const response = await fetchImpl(source, { headers: authHeaders(token), redirect: 'manual' });
  if (response.status >= 300 && response.status < 400) {
    const location = response.headers.get('location');
    if (!location) fail('GitHub artifact redirect did not include a signed archive URL');
    return { url: location, artifactUrl: source };
  }
  if (response.status === 200 || response.status === 206) return { url: source, artifactUrl: source };
  fail(`GitHub artifact archive request failed (${response.status})`);
}

async function requestRange(url, start, end, { fetchImpl = fetch, token, stats } = {}) {
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start) {
    throw new TypeError('invalid archive range');
  }
  const response = await fetchImpl(url, {
    headers: { ...authHeaders(token), range: `bytes=${start}-${end}` },
  });
  if (response.status !== 206) {
    fail(`GitHub artifact server did not honor Range ${start}-${end} (HTTP ${response.status})`);
  }
  const range = parseContentRange(response.headers.get('content-range'));
  if (range.start !== start || range.end !== end || range.total < end + 1) {
    fail(`GitHub artifact returned an unexpected range ${JSON.stringify(range)} for ${start}-${end}`);
  }
  const bytes = Buffer.from(await response.arrayBuffer());
  if (bytes.byteLength !== end - start + 1) fail(`GitHub artifact range ${start}-${end} was truncated`);
  if (stats) stats.rangeBytes += bytes.byteLength;
  return { bytes, total: range.total };
}

function readZip64Extra(extra, values) {
  let offset = 0;
  while (offset + 4 <= extra.length) {
    const id = extra.readUInt16LE(offset);
    const size = extra.readUInt16LE(offset + 2);
    const end = offset + 4 + size;
    if (end > extra.length) break;
    if (id === ZIP64_EXTRA_FIELD) {
      let cursor = offset + 4;
      const read = () => {
        if (cursor + 8 > end) fail('ZIP64 extra field is truncated');
        const value = Number(extra.readBigUInt64LE(cursor));
        cursor += 8;
        return value;
      };
      if (values.uncompressedSize === 0xffffffff) values.uncompressedSize = read();
      if (values.compressedSize === 0xffffffff) values.compressedSize = read();
      if (values.localHeaderOffset === 0xffffffff) values.localHeaderOffset = read();
      break;
    }
    offset = end;
  }
  return values;
}

export function parseZipCentralDirectory(bytes) {
  const entries = [];
  let offset = 0;
  while (offset < bytes.length) {
    if (bytes.length - offset < 46) fail(`ZIP central directory ended at ${offset}`);
    if (bytes.readUInt32LE(offset) !== ZIP_CENTRAL_HEADER) fail(`invalid ZIP central entry at ${offset}`);
    const flags = bytes.readUInt16LE(offset + 8);
    const method = bytes.readUInt16LE(offset + 10);
    const crc32 = bytes.readUInt32LE(offset + 16);
    const compressedSize = bytes.readUInt32LE(offset + 20);
    const uncompressedSize = bytes.readUInt32LE(offset + 24);
    const nameBytes = bytes.readUInt16LE(offset + 28);
    const extraBytes = bytes.readUInt16LE(offset + 30);
    const commentBytes = bytes.readUInt16LE(offset + 32);
    const localHeaderOffset = bytes.readUInt32LE(offset + 42);
    const end = offset + 46 + nameBytes + extraBytes + commentBytes;
    if (end > bytes.length) fail(`ZIP central entry at ${offset} is truncated`);
    const name = bytes.subarray(offset + 46, offset + 46 + nameBytes).toString('utf8');
    const extra = bytes.subarray(offset + 46 + nameBytes, offset + 46 + nameBytes + extraBytes);
    const values = readZip64Extra(extra, { compressedSize, uncompressedSize, localHeaderOffset });
    if (!name || name.includes('\\') || name.startsWith('/') || name.split('/').includes('..')) {
      fail(`ZIP entry has an unsafe path: ${name || '<empty>'}`);
    }
    entries.push({
      name,
      flags,
      method,
      crc32,
      compressedSize: values.compressedSize,
      uncompressedSize: values.uncompressedSize,
      localHeaderOffset: values.localHeaderOffset,
      directory: name.endsWith('/'),
    });
    offset = end;
  }
  return entries;
}

function findLastSignature(bytes, signature) {
  for (let offset = bytes.length - 4; offset >= 0; offset -= 1) {
    if (bytes.readUInt32LE(offset) === signature) return offset;
  }
  return -1;
}

function parseZip64End(tail, eocdOffset, totalSize) {
  const locatorOffset = eocdOffset - 20;
  if (locatorOffset < 0 || tail.readUInt32LE(locatorOffset) !== ZIP64_LOCATOR_HEADER) {
    fail('ZIP64 archive is missing its locator');
  }
  const zip64Offset = Number(tail.readBigUInt64LE(locatorOffset + 8));
  const zip64End = zip64Offset >= totalSize ? null : zip64Offset;
  if (zip64End === null) fail('ZIP64 end record is outside the archive');
  return zip64End;
}

export async function readArtifactIndex({
  archiveUrl,
  repository,
  runId,
  artifactId,
  artifactName,
  token,
  fetchImpl = fetch,
  stats = { rangeBytes: 0 },
}) {
  const resolved = await resolveArchiveUrl({
    archiveUrl,
    repository,
    runId,
    artifactId,
    artifactName,
    token,
    fetchImpl,
  });
  const first = await requestRange(resolved.url, 0, 0, { fetchImpl, token: undefined, stats });
  const totalSize = first.total;
  const tailStart = Math.max(0, totalSize - MAX_END_SCAN_BYTES);
  const tail = (await requestRange(resolved.url, tailStart, totalSize - 1, { fetchImpl, stats })).bytes;
  const eocdOffset = findLastSignature(tail, ZIP_END_HEADER);
  if (eocdOffset < 0) fail('ZIP end-of-central-directory record was not found');
  let entries = tail.readUInt16LE(eocdOffset + 10);
  let centralSize = tail.readUInt32LE(eocdOffset + 12);
  let centralOffset = tail.readUInt32LE(eocdOffset + 16);
  if (entries === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) {
    const zip64Offset = parseZip64End(tail, eocdOffset, totalSize);
    const record = (await requestRange(resolved.url, zip64Offset, zip64Offset + 56 - 1, { fetchImpl, stats })).bytes;
    if (record.readUInt32LE(0) !== ZIP64_END_HEADER) fail('invalid ZIP64 end record');
    entries = Number(record.readBigUInt64LE(32));
    centralSize = Number(record.readBigUInt64LE(40));
    centralOffset = Number(record.readBigUInt64LE(48));
  }
  if (!Number.isSafeInteger(centralSize) || !Number.isSafeInteger(centralOffset) || centralSize <= 0) {
    fail('ZIP central directory has invalid bounds');
  }
  const central = (
    await requestRange(resolved.url, centralOffset, centralOffset + centralSize - 1, { fetchImpl, stats })
  ).bytes;
  const parsed = parseZipCentralDirectory(central);
  if (parsed.length !== entries) fail(`ZIP entry count mismatch: expected ${entries}, received ${parsed.length}`);
  return { ...resolved, totalSize, entries: parsed, stats };
}

function globRegex(pattern) {
  let source = '^';
  for (const character of pattern) {
    if (character === '*') source += '.*';
    else if (character === '?') source += '.';
    else source += /[\\^$+?.()|[\]{}]/.test(character) ? `\\${character}` : character;
  }
  return new RegExp(`${source}$`);
}

export function selectArtifactEntries(entries, { includes = [], files = [] } = {}) {
  const exact = new Set(files);
  const patterns = includes.map(globRegex);
  return entries.filter(
    (entry) => !entry.directory && (exact.has(entry.name) || patterns.some((pattern) => pattern.test(entry.name))),
  );
}

function crc32(bytes, previous = 0) {
  let crc = (previous ^ 0xffffffff) >>> 0;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

async function* entryChunks(url, entry, { fetchImpl, stats }) {
  for (let offset = 0; offset < entry.compressedSize; offset += RANGE_CHUNK_BYTES) {
    const end = Math.min(entry.compressedSize - 1, offset + RANGE_CHUNK_BYTES - 1);
    const range = await requestRange(url, entry.dataOffset + offset, entry.dataOffset + end, {
      fetchImpl,
      stats,
    });
    yield range.bytes;
  }
}

export async function extractArtifactEntry({ archive, entry, outputRoot, fetchImpl = fetch }) {
  const header = (
    await requestRange(archive.url, entry.localHeaderOffset, entry.localHeaderOffset + 30 - 1, {
      fetchImpl,
      stats: archive.stats,
    })
  ).bytes;
  if (header.readUInt32LE(0) !== ZIP_LOCAL_HEADER) fail(`invalid ZIP local header for ${entry.name}`);
  const nameBytes = header.readUInt16LE(26);
  const extraBytes = header.readUInt16LE(28);
  entry.dataOffset = entry.localHeaderOffset + 30 + nameBytes + extraBytes;
  const destination = path.join(outputRoot, entry.name);
  const relative = path.relative(path.resolve(outputRoot), path.resolve(destination));
  if (relative.startsWith('..') || path.isAbsolute(relative)) fail(`entry escapes output root: ${entry.name}`);
  fs.mkdirSync(path.dirname(destination), { recursive: true });
  const source = Readable.from(entryChunks(archive.url, entry, { fetchImpl, stats: archive.stats }));
  const decoded = entry.method === 0 ? source : entry.method === 8 ? source.pipe(createInflateRaw()) : null;
  if (!decoded) fail(`ZIP entry ${entry.name} uses unsupported compression method ${entry.method}`);
  const digest = crypto.createHash('sha256');
  let bytes = 0;
  let checksum = 0;
  const inspect = new Transform({
    transform(chunk, _encoding, callback) {
      bytes += chunk.byteLength;
      checksum = crc32(chunk, checksum);
      digest.update(chunk);
      callback(null, chunk);
    },
  });
  await pipeline(decoded, inspect, fs.createWriteStream(destination, { flags: 'w' }));
  if (bytes !== entry.uncompressedSize || checksum !== entry.crc32) {
    fs.rmSync(destination, { force: true });
    fail(`ZIP entry ${entry.name} failed size/CRC verification`);
  }
  return { path: entry.name, bytes, sha256: digest.digest('hex') };
}

function argumentValues(arguments_, name) {
  const values = [];
  for (let index = 0; index < arguments_.length; index += 1) {
    if (arguments_[index] === name && arguments_[index + 1]) values.push(arguments_[index + 1]);
  }
  return values;
}

function argumentValue(arguments_, name, fallback = '') {
  const index = arguments_.indexOf(name);
  return index >= 0 && arguments_[index + 1] ? arguments_[index + 1] : fallback;
}

async function main() {
  const arguments_ = process.argv.slice(2);
  const outputRoot = argumentValue(arguments_, '--output');
  const token = process.env.GH_TOKEN || process.env.GITHUB_TOKEN || '';
  const includes = argumentValues(arguments_, '--include');
  const fileList = argumentValue(arguments_, '--file-list');
  const files = fileList ? JSON.parse(fs.readFileSync(fileList, 'utf8')) : [];
  if (!outputRoot && !arguments_.includes('--list')) fail('--output is required unless --list is used');
  if (!includes.length && !files.length && !arguments_.includes('--list')) fail('provide --include or --file-list');
  const stats = { rangeBytes: 0 };
  const archive = await readArtifactIndex({
    archiveUrl: argumentValue(arguments_, '--archive-url'),
    repository: argumentValue(arguments_, '--repository', process.env.GITHUB_REPOSITORY),
    runId: argumentValue(arguments_, '--run-id'),
    artifactId: argumentValue(arguments_, '--artifact-id'),
    artifactName: argumentValue(arguments_, '--artifact-name'),
    token,
    stats,
  });
  if (arguments_.includes('--list')) {
    process.stdout.write(`${JSON.stringify({ totalBytes: archive.totalSize, entries: archive.entries }, null, 2)}\n`);
    return;
  }
  const selected = selectArtifactEntries(archive.entries, { includes, files });
  if (!selected.length) fail('selection did not match any artifact entries');
  const missing = [...new Set([...includes, ...files])].filter(
    (requested) => !selected.some((entry) => entry.name === requested || globRegex(requested).test(entry.name)),
  );
  if (missing.length) fail(`selected artifact entries were not found: ${missing.join(', ')}`);
  fs.mkdirSync(outputRoot, { recursive: true });
  const extracted = [];
  for (const entry of selected) extracted.push(await extractArtifactEntry({ archive, entry, outputRoot }));
  const reportPath = argumentValue(arguments_, '--report');
  const report = {
    version: 1,
    source: 'github-actions-artifact-range',
    totalArtifactBytes: archive.totalSize,
    selectedFiles: extracted.length,
    selectedUncompressedBytes: extracted.reduce((total, item) => total + item.bytes, 0),
    rangeBytes: stats.rangeBytes,
    entries: extracted,
  };
  if (reportPath) {
    fs.mkdirSync(path.dirname(path.resolve(reportPath)), { recursive: true });
    fs.writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`);
  }
  process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  });
}
