import { describe, expect, it } from 'vitest';
import {
  compareContentAddressedItems,
  compareReleasePair,
  formatBenchmarkMarkdown,
} from '../../scripts/benchmark-incremental-releases.mjs';

describe('incremental release benchmark', () => {
  it('counts each missing digest once and reuses target content by digest', () => {
    expect(
      compareContentAddressedItems(
        [
          { sha256: 'same', bytes: 10 },
          { sha256: 'old', bytes: 20 },
        ],
        [
          { sha256: 'same', bytes: 10 },
          { sha256: 'new', bytes: 30 },
          { sha256: 'new', bytes: 30 },
        ],
      ),
    ).toEqual({
      targetBytes: 40,
      targetObjects: 2,
      downloadBytes: 30,
      downloadObjects: 1,
      reusedBytes: 10,
      reusedObjects: 1,
    });
  });

  it('compares current, fixed digest, file CAS, and FastCDC strategies', () => {
    const from = {
      releaseId: 'v1',
      package: {
        bytes: 100,
        segments: [
          { sha256: 'segment-same', bytes: 50 },
          { sha256: 'segment-old', bytes: 50 },
        ],
      },
      assets: [
        { path: 'a', sha256: 'asset-same', bytes: 60 },
        { path: 'b', sha256: 'asset-old', bytes: 40 },
      ],
    };
    const to = {
      releaseId: 'v2',
      package: {
        bytes: 110,
        segments: [
          { sha256: 'segment-same', bytes: 50 },
          { sha256: 'segment-new', bytes: 60 },
        ],
      },
      assets: [
        { path: 'a', sha256: 'asset-same', bytes: 60 },
        { path: 'b', sha256: 'asset-new', bytes: 45 },
        { path: 'c', sha256: 'asset-added', bytes: 5 },
      ],
    };
    const pair = compareReleasePair(from, to, {
      from: {
        configurations: [
          {
            minimumBytes: 1,
            averageBytes: 4,
            maximumBytes: 16,
            elapsedMs: 2,
            chunks: [
              { sha256: 'chunk-same', bytes: 70 },
              { sha256: 'chunk-old', bytes: 30 },
            ],
          },
        ],
      },
      to: {
        configurations: [
          {
            minimumBytes: 1,
            averageBytes: 4,
            maximumBytes: 16,
            elapsedMs: 3,
            chunks: [
              { sha256: 'chunk-same', bytes: 70 },
              { sha256: 'chunk-new', bytes: 40 },
            ],
          },
        ],
      },
    });

    expect(pair.currentReleaseBoundSegments.downloadBytes).toBe(110);
    expect(pair.fixedDigestSegments.downloadBytes).toBe(60);
    expect(pair.fileCas.downloadBytes).toBe(50);
    expect(pair.fastCdc[0].downloadBytes).toBe(40);
    expect(pair.assetPaths).toEqual({ unchanged: 1, changed: 1, added: 1, removed: 0 });
    expect(
      formatBenchmarkMarkdown({
        generatedAt: '2026-07-30T00:00:00.000Z',
        pairs: [pair],
      }),
    ).toContain('| File-level CAS | 50 B | 2 | 3 | 60 B | 54.55% |');
  });
});
