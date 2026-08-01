import { describe, expect, it } from 'vitest';
import * as benchmarkModule from '../../scripts/benchmark-incremental-releases.mjs';

const { compareContentAddressedItems, compareReleasePair, formatBenchmarkMarkdown } = benchmarkModule;
const { compareHybridPlanner, releaseManifestPath } = benchmarkModule as typeof benchmarkModule & {
  compareHybridPlanner: (from: unknown, to: unknown) => any;
  releaseManifestPath: (releaseId: string, version?: 4 | 5) => string;
};

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

  it('plans unchanged paths, ordinary whole blobs, and per-file FastCDC chunks independently', () => {
    const from = {
      version: 5,
      releaseId: 'v1',
      package: {
        bytes: 220,
        segments: [
          { sha256: 'segment-same', bytes: 100 },
          { sha256: 'segment-old', bytes: 120 },
        ],
      },
      assets: [
        { path: 'unchanged.js', sha256: 'asset-same', bytes: 60 },
        { path: 'ordinary.js', sha256: 'asset-old', bytes: 40 },
        {
          path: 'runtime.wasm',
          sha256: 'runtime-old',
          bytes: 110,
          representations: {
            whole: { sha256: 'runtime-old', bytes: 110 },
            fastcdc: {
              chunks: [
                { offset: 0, sha256: 'chunk-same', bytes: 70 },
                { offset: 70, sha256: 'chunk-old', bytes: 40 },
              ],
            },
          },
        },
        { path: 'removed.js', sha256: 'removed', bytes: 10 },
      ],
    };
    const to = {
      version: 5,
      releaseId: 'v2',
      package: {
        bytes: 230,
        segments: [
          { sha256: 'segment-same', bytes: 100 },
          { sha256: 'segment-new', bytes: 130 },
        ],
      },
      assets: [
        { path: 'unchanged.js', sha256: 'asset-same', bytes: 60 },
        { path: 'ordinary.js', sha256: 'asset-new', bytes: 45 },
        {
          path: 'runtime.wasm',
          sha256: 'runtime-new',
          bytes: 110,
          representations: {
            whole: { sha256: 'runtime-new', bytes: 110 },
            fastcdc: {
              chunks: [
                { offset: 0, sha256: 'chunk-same', bytes: 70 },
                { offset: 70, sha256: 'chunk-new', bytes: 40 },
              ],
            },
          },
        },
        { path: 'added.js', sha256: 'asset-added', bytes: 5 },
      ],
    };

    const pair = compareReleasePair(from as any, to as any) as any;

    expect(pair.currentReleaseBoundSegments.downloadBytes).toBe(230);
    expect(pair.fixedDigestSegments.downloadBytes).toBe(130);
    expect(pair.fileCas.downloadBytes).toBe(160);
    expect(pair.hybridPlanner).toMatchObject({
      targetBytes: 220,
      targetObjects: 5,
      downloadBytes: 90,
      downloadObjects: 3,
      reusedBytes: 130,
      reusedObjects: 2,
      paths: { unchanged: 1, whole: 2, fastCdc: 1, removed: 1 },
    });
    expect(pair.hybridPlanner.decisions).toEqual([
      {
        path: 'added.js',
        change: 'added',
        strategy: 'whole',
        targetBytes: 5,
        downloadBytes: 5,
        downloadObjects: 1,
      },
      {
        path: 'ordinary.js',
        change: 'changed',
        strategy: 'whole',
        targetBytes: 45,
        downloadBytes: 45,
        downloadObjects: 1,
      },
      {
        path: 'runtime.wasm',
        change: 'changed',
        strategy: 'fastcdc',
        targetBytes: 110,
        downloadBytes: 40,
        downloadObjects: 1,
      },
      {
        path: 'unchanged.js',
        change: 'unchanged',
        strategy: 'unchanged',
        targetBytes: 60,
        downloadBytes: 0,
        downloadObjects: 0,
      },
      {
        path: 'removed.js',
        change: 'removed',
        strategy: 'removed',
        targetBytes: 0,
        downloadBytes: 0,
        downloadObjects: 0,
      },
    ]);
    expect(pair.assetPaths).toEqual({ unchanged: 1, changed: 2, added: 1, removed: 1 });

    const markdown = formatBenchmarkMarkdown({
      generatedAt: '2026-07-30T00:00:00.000Z',
      pairs: [pair],
    });
    expect(markdown).toContain('| Hybrid path planner | 90 B | 3 | 5 | 130 B | 59.09% |');
    expect(markdown).toContain(
      'Hybrid paths: 1 unchanged (zero download), 2 whole blobs, 1 per-file FastCDC, 1 removed.',
    );
    expect(markdown).not.toContain('FastCDC avg');
  });

  it('uses whole blobs for a v4 companion and cannot infer chunk reuse from absent v5 metadata', () => {
    const fromV4 = {
      version: 4,
      releaseId: 'v1',
      assets: [
        { path: 'unchanged.js', sha256: 'same', bytes: 10 },
        { path: 'runtime.wasm', sha256: 'old-runtime', bytes: 100 },
      ],
    };
    const toV4 = {
      version: 4,
      releaseId: 'v2',
      assets: [
        { path: 'unchanged.js', sha256: 'same', bytes: 10 },
        { path: 'runtime.wasm', sha256: 'new-runtime', bytes: 100 },
      ],
    };
    expect(compareHybridPlanner(fromV4, toV4)).toMatchObject({
      downloadBytes: 100,
      downloadObjects: 1,
      paths: { unchanged: 1, whole: 1, fastCdc: 0, removed: 0 },
    });

    const toV5 = {
      ...toV4,
      version: 5,
      assets: [
        toV4.assets[0],
        {
          ...toV4.assets[1],
          representations: {
            whole: { sha256: 'new-runtime', bytes: 100 },
            fastcdc: {
              chunks: [
                { offset: 0, sha256: 'new-chunk-a', bytes: 40 },
                { offset: 40, sha256: 'new-chunk-b', bytes: 60 },
              ],
            },
          },
        },
      ],
    };
    expect(compareHybridPlanner(fromV4, toV5)).toMatchObject({
      downloadBytes: 100,
      downloadObjects: 2,
      paths: { unchanged: 1, whole: 0, fastCdc: 1, removed: 0 },
    });
  });

  it('deduplicates one missing object across paths and validates FastCDC coverage', () => {
    const result = compareHybridPlanner(
      { version: 5, releaseId: 'v1', assets: [] },
      {
        version: 5,
        releaseId: 'v2',
        assets: [
          { path: 'a.bin', sha256: 'shared', bytes: 8 },
          { path: 'b.bin', sha256: 'shared', bytes: 8 },
        ],
      },
    );
    expect(result).toMatchObject({
      targetBytes: 16,
      targetObjects: 1,
      downloadBytes: 8,
      downloadObjects: 1,
    });
    expect(result.decisions.map((decision: any) => decision.downloadBytes)).toEqual([8, 0]);

    expect(() =>
      compareHybridPlanner(
        { version: 5, releaseId: 'v1', assets: [] },
        {
          version: 5,
          releaseId: 'v2',
          assets: [
            {
              path: 'broken.wasm',
              sha256: 'broken',
              bytes: 8,
              representations: {
                whole: { sha256: 'broken', bytes: 8 },
                fastcdc: { chunks: [{ offset: 1, sha256: 'chunk', bytes: 8 }] },
              },
            },
          ],
        },
      ),
    ).toThrow(/invalid FastCDC chunk/);
  });

  it('defaults to v5 manifest paths and supports the v4 companion explicitly', () => {
    expect(releaseManifestPath('v0.6.0+canary')).toBe('/releases/v0.6.0%2Bcanary/manifest.json');
    expect(releaseManifestPath('v0.6.0+canary', 4)).toBe('/releases/v0.6.0%2Bcanary/manifest-v4.json');
    expect(() => releaseManifestPath('release', 3 as 4)).toThrow(/must be 4 or 5/);
  });
});
