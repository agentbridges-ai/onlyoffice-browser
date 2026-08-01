#!/usr/bin/env node

import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { compareContentAddressedItems } from './benchmark-incremental-releases.mjs';

const SCRIPT_DIR = path.dirname(fileURLToPath(import.meta.url));
const FASTCDC_MANIFEST = path.resolve(SCRIPT_DIR, '../tools/fastcdc-index/Cargo.toml');
const FASTCDC_AVERAGES = [256 * 1024, 1024 * 1024];
const FIXED_SEGMENT_BYTES = 1024 * 1024;
const LOGICAL_TARGET_BYTES = 1024 * 1024;
const LOGICAL_MAX_BYTES = 1536 * 1024;

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function deterministicBytes(bytes, seed) {
  const output = Buffer.allocUnsafe(bytes);
  let state = seed >>> 0;
  for (let index = 0; index < output.length; index += 1) {
    state ^= state << 13;
    state ^= state >>> 17;
    state ^= state << 5;
    output[index] = state & 0xff;
  }
  return output;
}

function file(name, bytes, seed) {
  return { name, bytes: deterministicBytes(bytes, seed) };
}

function baseFiles() {
  return [
    file('010-shell.js', 320 * 1024, 10),
    file('020-word.js', 780 * 1024, 20),
    file('030-cell.js', 690 * 1024, 30),
    file('040-slide.js', 740 * 1024, 40),
    file('050-runtime.wasm', 4 * 1024 * 1024, 50),
    file('060-aptos.ttf', 920 * 1024, 60),
    file('070-dengxian.ttf', 1680 * 1024, 70),
    file('080-dictionary.bin', 540 * 1024, 80),
  ];
}

function assemble(files) {
  return Buffer.concat(files.map((entry) => entry.bytes));
}

function descriptors(chunks) {
  return chunks.map((bytes) => ({ sha256: sha256(bytes), bytes: bytes.byteLength }));
}

function fixedChunks(bytes, chunkBytes = FIXED_SEGMENT_BYTES) {
  const chunks = [];
  for (let offset = 0; offset < bytes.byteLength; offset += chunkBytes) {
    chunks.push(bytes.subarray(offset, Math.min(offset + chunkBytes, bytes.byteLength)));
  }
  return descriptors(chunks);
}

function logicalChunks(files) {
  const sorted = [...files].sort((left, right) => left.name.localeCompare(right.name));
  const chunks = [];
  let current = [];
  let currentBytes = 0;
  const flush = () => {
    if (!current.length) return;
    chunks.push(Buffer.concat(current.map((entry) => entry.bytes)));
    current = [];
    currentBytes = 0;
  };
  for (const entry of sorted) {
    if (
      current.length &&
      (currentBytes + entry.bytes.byteLength > LOGICAL_MAX_BYTES || currentBytes >= LOGICAL_TARGET_BYTES)
    ) {
      flush();
    }
    current.push(entry);
    currentBytes += entry.bytes.byteLength;
  }
  flush();
  return descriptors(chunks);
}

function fileCas(files) {
  return descriptors(files.map((entry) => entry.bytes));
}

function buildFastCdcBinary() {
  const targetDirectory =
    process.env.CARGO_TARGET_DIR || path.join(os.tmpdir(), 'onlyoffice-fastcdc-experiment-target');
  const result = spawnSync('cargo', ['build', '--quiet', '--release', '--manifest-path', FASTCDC_MANIFEST], {
    encoding: 'utf8',
    env: { ...process.env, CARGO_TARGET_DIR: targetDirectory },
  });
  if (result.status !== 0) throw new Error(result.stderr || 'FastCDC experiment build failed');
  return path.join(targetDirectory, 'release', 'onlyoffice-fastcdc-index');
}

function fastCdcIndexes(binary, filePath) {
  const result = spawnSync(binary, [filePath, ...FASTCDC_AVERAGES.map(String)], {
    encoding: 'utf8',
    maxBuffer: 16 * 1024 * 1024,
  });
  if (result.status !== 0) throw new Error(result.stderr || 'FastCDC experiment failed');
  return JSON.parse(result.stdout);
}

function mutateLargeFile(files) {
  return files.map((entry) => {
    if (entry.name !== '050-runtime.wasm') return entry;
    const changed = Buffer.from(entry.bytes);
    deterministicBytes(64 * 1024, 5050).copy(changed, 2 * 1024 * 1024);
    return { ...entry, bytes: changed };
  });
}

function scenarios() {
  const base = baseFiles();
  const inserted = file('000-new-font.ttf', 192 * 1024, 99);
  return [
    { name: 'prefix-insert', from: base, to: [inserted, ...base], actualChangedBytes: inserted.bytes.byteLength },
    {
      name: 'prefix-delete',
      from: [inserted, ...base],
      to: base,
      actualChangedBytes: 0,
    },
    {
      name: 'large-file-middle-replace',
      from: base,
      to: mutateLargeFile(base),
      actualChangedBytes: 64 * 1024,
    },
    {
      name: 'file-reorder',
      from: base,
      to: [base[1], base[0], ...base.slice(2)],
      actualChangedBytes: 0,
    },
  ];
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KiB`;
  return `${(bytes / 1024 / 1024).toFixed(2)} MiB`;
}

export function compareSyntheticScenario(scenario, fastCdc) {
  const fromBytes = assemble(scenario.from);
  const toBytes = assemble(scenario.to);
  const result = {
    name: scenario.name,
    actualChangedBytes: scenario.actualChangedBytes,
    fixed: compareContentAddressedItems(fixedChunks(fromBytes), fixedChunks(toBytes)),
    logical: compareContentAddressedItems(logicalChunks(scenario.from), logicalChunks(scenario.to)),
    fileCas: compareContentAddressedItems(fileCas(scenario.from), fileCas(scenario.to)),
    fastCdc: [],
  };
  if (fastCdc) {
    for (const fromConfiguration of fastCdc.from.configurations) {
      const toConfiguration = fastCdc.to.configurations.find(
        (candidate) => candidate.averageBytes === fromConfiguration.averageBytes,
      );
      result.fastCdc.push({
        averageBytes: fromConfiguration.averageBytes,
        ...compareContentAddressedItems(fromConfiguration.chunks, toConfiguration.chunks),
      });
    }
  }
  return result;
}

async function run() {
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-chunking-experiment-'));
  try {
    const binary = buildFastCdcBinary();
    const results = [];
    for (const scenario of scenarios()) {
      const fromPath = path.join(temporaryDirectory, `${scenario.name}-from.pack`);
      const toPath = path.join(temporaryDirectory, `${scenario.name}-to.pack`);
      fs.writeFileSync(fromPath, assemble(scenario.from));
      fs.writeFileSync(toPath, assemble(scenario.to));
      results.push(
        compareSyntheticScenario(scenario, {
          from: fastCdcIndexes(binary, fromPath),
          to: fastCdcIndexes(binary, toPath),
        }),
      );
    }
    const prefixInsert = results.find((result) => result.name === 'prefix-insert');
    const prefixDelete = results.find((result) => result.name === 'prefix-delete');
    const largeFileChange = results.find((result) => result.name === 'large-file-middle-replace');
    const fileReorder = results.find((result) => result.name === 'file-reorder');
    if (
      prefixInsert.fixed.downloadBytes !== prefixInsert.fixed.targetBytes ||
      prefixInsert.fileCas.downloadBytes !== prefixInsert.actualChangedBytes ||
      prefixDelete.fileCas.downloadBytes !== 0 ||
      largeFileChange.fastCdc[0].downloadBytes >= largeFileChange.fileCas.downloadBytes ||
      fileReorder.fileCas.downloadBytes !== 0
    ) {
      throw new Error(`chunking experiment invariants failed\n${JSON.stringify(results, null, 2)}`);
    }

    process.stdout.write(
      [
        '# Content-defined chunking mutation experiment',
        '',
        '| Mutation | Actual change | Fixed 1 MiB | Stable logical groups | File CAS | FastCDC 256 KiB | FastCDC 1 MiB |',
        '| --- | ---: | ---: | ---: | ---: | ---: | ---: |',
        ...results.map((result) => {
          const fast256 = result.fastCdc.find((entry) => entry.averageBytes === 256 * 1024);
          const fast1024 = result.fastCdc.find((entry) => entry.averageBytes === 1024 * 1024);
          return `| ${result.name} | ${formatBytes(result.actualChangedBytes)} | ${formatBytes(result.fixed.downloadBytes)} | ${formatBytes(result.logical.downloadBytes)} | ${formatBytes(result.fileCas.downloadBytes)} | ${formatBytes(fast256.downloadBytes)} | ${formatBytes(fast1024.downloadBytes)} |`;
        }),
        '',
        JSON.stringify({ version: 1, results }),
        '',
      ].join('\n'),
    );
  } finally {
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
  }
}

const isMain = process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1]);
if (isMain) {
  run().catch((error) => {
    console.error(error instanceof Error ? error.stack || error.message : String(error));
    process.exitCode = 1;
  });
}
