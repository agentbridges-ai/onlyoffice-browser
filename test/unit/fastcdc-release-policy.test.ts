import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
// The release policy is an executable Node ESM build helper and intentionally
// has no declaration output in the browser package.
// @ts-expect-error JavaScript build script has no declaration output.
import * as fastCdcReleasePolicy from '../../scripts/fastcdc-release-policy.mjs';

const {
  FASTCDC_RELEASE_POLICY,
  buildFastCdcRepresentation,
  buildFastCdcRepresentationFromEvidence,
  evaluateFastCdcReleasePolicy,
  parseFastCdcEvidence,
  readFastCdcEvidence,
  runFastCdcIndexer,
} = fastCdcReleasePolicy;

const temporaryDirectories: string[] = [];
const MiB = 1024 * 1024;

function temp(): string {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-fastcdc-policy-'));
  temporaryDirectories.push(directory);
  return directory;
}

function sha256(bytes: Uint8Array): string {
  return crypto.createHash('sha256').update(bytes).digest('hex');
}

function evidence(
  samples: Array<{
    fromReleaseId: string;
    toReleaseId: string;
    wholeDownloadBytes: number;
    fastcdcDownloadBytes: number;
  }>,
) {
  return {
    version: 1,
    assets: [{ path: 'wasm/runtime.wasm', samples }],
  };
}

function qualifyingSamples() {
  return [
    {
      fromReleaseId: 'release-a',
      toReleaseId: 'release-b',
      wholeDownloadBytes: 2 * MiB,
      fastcdcDownloadBytes: 1.5 * MiB,
    },
    {
      fromReleaseId: 'release-b',
      toReleaseId: 'release-c',
      wholeDownloadBytes: 2 * MiB,
      fastcdcDownloadBytes: 1.5 * MiB,
    },
  ];
}

function createIndexedFile(directory: string, chunks: Buffer[]) {
  const inputPath = path.join(directory, 'runtime.wasm');
  const source = Buffer.concat(chunks);
  fs.writeFileSync(inputPath, source);
  const records = [];
  let offset = 0;
  for (const chunk of chunks) {
    records.push({ offset, bytes: chunk.byteLength, sha256: sha256(chunk) });
    offset += chunk.byteLength;
  }
  return {
    inputPath,
    source,
    index: {
      path: inputPath,
      bytes: source.byteLength,
      configurations: [
        {
          minimumBytes: FASTCDC_RELEASE_POLICY.minBytes,
          averageBytes: FASTCDC_RELEASE_POLICY.averageBytes,
          maximumBytes: FASTCDC_RELEASE_POLICY.maxBytes,
          elapsedMs: 0,
          chunks: records,
        },
      ],
    },
  };
}

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    fs.rmSync(directory, { recursive: true, force: true });
  }
});

describe('FastCDC release eligibility evidence', () => {
  it('selects every >=8 MiB asset to keep canonical Cache Storage writes bounded', () => {
    const decision = evaluateFastCdcReleasePolicy({
      assetPath: 'wasm/runtime.wasm',
      assetBytes: 8 * MiB,
      evidence: evidence(qualifyingSamples()),
    });
    expect(decision).toMatchObject({
      selected: true,
      reason: 'bounded-cache-write',
      samples: 2,
      wholeDownloadBytes: 4 * MiB,
      fastcdcDownloadBytes: 3 * MiB,
      savingsBytes: MiB,
      savingsRatio: 0.25,
      policyId: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0',
    });
  });

  it('keeps file CAS below the bounded-write threshold', () => {
    expect(
      evaluateFastCdcReleasePolicy({
        assetPath: 'wasm/runtime.wasm',
        assetBytes: 8 * MiB - 1,
        evidence: evidence(qualifyingSamples()),
      }),
    ).toMatchObject({ selected: false, reason: 'asset-too-small' });
  });

  it('uses FastCDC even when historical samples show no incremental savings', () => {
    expect(
      evaluateFastCdcReleasePolicy({
        assetPath: 'wasm/runtime.wasm',
        assetBytes: 8 * MiB,
        evidence: evidence([]),
      }),
    ).toMatchObject({ selected: true, reason: 'bounded-cache-write', samples: 0 });
  });

  it('reads and validates a versioned evidence JSON file', () => {
    const directory = temp();
    const evidencePath = path.join(directory, 'fastcdc-evidence.json');
    fs.writeFileSync(evidencePath, JSON.stringify(evidence(qualifyingSamples())));
    expect(readFastCdcEvidence(evidencePath)).toEqual(evidence(qualifyingSamples()));

    expect(() =>
      parseFastCdcEvidence({
        version: 1,
        assets: [...evidence(qualifyingSamples()).assets, ...evidence(qualifyingSamples()).assets],
      }),
    ).toThrow(/Duplicate FastCDC evidence asset/);
    expect(() => parseFastCdcEvidence(evidence([qualifyingSamples()[0], { ...qualifyingSamples()[0] }]))).toThrow(
      /Duplicate FastCDC evidence transition/,
    );
  });
});

describe('FastCDC release representation builder', () => {
  it('invokes the maintained Rust v2020 indexer with the fixed 256 KiB policy', () => {
    const spawnSyncImpl = vi.fn((..._arguments: unknown[]) => ({
      status: 0,
      stdout: JSON.stringify({ bytes: 0, configurations: [] }),
      stderr: '',
    }));
    expect(
      runFastCdcIndexer('/tmp/runtime.wasm', {
        spawnSyncImpl,
        cargoTargetDirectory: '/tmp/onlyoffice-fastcdc-target',
      }),
    ).toEqual({ bytes: 0, configurations: [] });
    expect(spawnSyncImpl).toHaveBeenCalledOnce();
    const [command, arguments_, options] = spawnSyncImpl.mock.calls[0] as [
      string,
      string[],
      { env: Record<string, string> },
    ];
    expect(command).toBe('cargo');
    expect(arguments_).toEqual(
      expect.arrayContaining([
        'run',
        '--release',
        '--manifest-path',
        expect.stringMatching(/tools\/fastcdc-index\/Cargo\.toml$/),
        '/tmp/runtime.wasm',
        String(256 * 1024),
      ]),
    );
    expect(options.env.CARGO_TARGET_DIR).toBe('/tmp/onlyoffice-fastcdc-target');
  });

  it('verifies every chunk and the reconstructed asset before writing canonical blobs', () => {
    const directory = temp();
    const output = temp();
    const chunks = Array.from({ length: 16 }, (_, index) => Buffer.alloc(MiB, index));
    const indexed = createIndexedFile(directory, chunks);
    const evidencePath = path.join(directory, 'fastcdc-evidence.json');
    fs.writeFileSync(evidencePath, JSON.stringify(evidence(qualifyingSamples())));
    const indexer = vi.fn(() => indexed.index);

    const result = buildFastCdcRepresentationFromEvidence({
      assetPath: 'wasm/runtime.wasm',
      inputPath: indexed.inputPath,
      output,
      expectedBytes: indexed.source.byteLength,
      expectedSha256: sha256(indexed.source),
      evidencePath,
      indexer,
    });

    expect(result.selected).toBe(true);
    expect(result.representation).toMatchObject({
      algorithm: 'fastcdc-v2020',
      minBytes: 64 * 1024,
      averageBytes: 256 * 1024,
      maxBytes: MiB,
      normalization: 1,
      seed: 0,
      chunks: indexed.index.configurations[0].chunks,
    });
    expect(indexer).toHaveBeenCalledWith(indexed.inputPath, FASTCDC_RELEASE_POLICY);
    const blobDirectory = path.join(output, 'blobs', 'sha256');
    expect(fs.readdirSync(blobDirectory).sort()).toEqual(
      indexed.index.configurations[0].chunks.map((chunk) => chunk.sha256).sort(),
    );
    for (const [index, chunk] of indexed.index.configurations[0].chunks.entries()) {
      const blob = fs.readFileSync(path.join(blobDirectory, chunk.sha256));
      expect(blob.byteLength).toBe(chunks[index].byteLength);
      expect(sha256(blob)).toBe(sha256(chunks[index]));
    }

    expect(() =>
      buildFastCdcRepresentation({
        inputPath: indexed.inputPath,
        output,
        expectedBytes: indexed.source.byteLength,
        expectedSha256: sha256(indexed.source),
        indexer,
      }),
    ).not.toThrow();
  });

  it('does not invoke the indexer or create blobs below the bounded-write threshold', () => {
    const directory = temp();
    const output = temp();
    const inputPath = path.join(directory, 'small.wasm');
    fs.writeFileSync(inputPath, Buffer.alloc(1024));
    const indexer = vi.fn();
    const result = buildFastCdcRepresentationFromEvidence({
      assetPath: 'wasm/runtime.wasm',
      inputPath,
      output,
      expectedBytes: 1024,
      expectedSha256: sha256(Buffer.alloc(1024)),
      evidence: evidence(qualifyingSamples()),
      indexer,
    });
    expect(result).toMatchObject({
      selected: false,
      decision: { reason: 'asset-too-small' },
    });
    expect(indexer).not.toHaveBeenCalled();
    expect(fs.existsSync(path.join(output, 'blobs'))).toBe(false);
  });

  it.each([
    {
      name: 'an offset gap',
      mutate: (index: ReturnType<typeof createIndexedFile>['index']) => {
        index.configurations[0].chunks[1].offset += 1;
      },
      error: /Invalid FastCDC chunk/,
    },
    {
      name: 'an invalid chunk digest',
      mutate: (index: ReturnType<typeof createIndexedFile>['index']) => {
        index.configurations[0].chunks[0].sha256 = '0'.repeat(64);
      },
      error: /chunk digest mismatch/,
    },
    {
      name: 'incomplete coverage',
      mutate: (index: ReturnType<typeof createIndexedFile>['index']) => {
        index.configurations[0].chunks.pop();
      },
      error: /do not cover/,
    },
  ])('rejects $name before publishing any blob', ({ mutate, error }) => {
    const directory = temp();
    const output = temp();
    const indexed = createIndexedFile(directory, [Buffer.alloc(64 * 1024, 1), Buffer.alloc(64 * 1024, 2)]);
    mutate(indexed.index);
    expect(() =>
      buildFastCdcRepresentation({
        inputPath: indexed.inputPath,
        output,
        expectedBytes: indexed.source.byteLength,
        expectedSha256: sha256(indexed.source),
        indexer: () => indexed.index,
      }),
    ).toThrow(error);
    expect(fs.existsSync(path.join(output, 'blobs'))).toBe(false);
  });

  it('rejects a reconstructed whole-file digest mismatch', () => {
    const directory = temp();
    const output = temp();
    const indexed = createIndexedFile(directory, [Buffer.alloc(64 * 1024, 1), Buffer.alloc(64 * 1024, 2)]);
    expect(() =>
      buildFastCdcRepresentation({
        inputPath: indexed.inputPath,
        output,
        expectedBytes: indexed.source.byteLength,
        expectedSha256: 'f'.repeat(64),
        indexer: () => indexed.index,
      }),
    ).toThrow(/reconstructed file digest mismatch/);
    expect(fs.existsSync(path.join(output, 'blobs'))).toBe(false);
  });
});
