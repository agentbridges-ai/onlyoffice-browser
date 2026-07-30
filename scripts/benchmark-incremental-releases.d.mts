export interface ContentAddressedItem {
  sha256: string;
  bytes: number;
}

export interface AssetItem extends ContentAddressedItem {
  path: string;
}

export interface ContentComparison {
  targetBytes: number;
  targetObjects: number;
  downloadBytes: number;
  downloadObjects: number;
  reusedBytes: number;
  reusedObjects: number;
}

export interface FastCdcConfiguration {
  minimumBytes: number;
  averageBytes: number;
  maximumBytes: number;
  elapsedMs: number;
  chunks: ContentAddressedItem[];
}

export interface FastCdcIndex {
  configurations: FastCdcConfiguration[];
}

export interface BenchmarkManifest {
  releaseId: string;
  package: {
    bytes: number;
    segments: ContentAddressedItem[];
  };
  assets: AssetItem[];
}

export interface ReleaseComparison {
  fromReleaseId: string;
  toReleaseId: string;
  logicalTargetBytes: number;
  assetPaths: {
    unchanged: number;
    changed: number;
    added: number;
    removed: number;
  };
  currentReleaseBoundSegments: ContentComparison;
  fixedDigestSegments: ContentComparison;
  fileCas: ContentComparison;
  fastCdc: Array<
    ContentComparison & {
      minimumBytes: number;
      averageBytes: number;
      maximumBytes: number;
      fromElapsedMs: number;
      toElapsedMs: number;
    }
  >;
}

export function compareContentAddressedItems(
  fromItems: ContentAddressedItem[],
  toItems: ContentAddressedItem[],
): ContentComparison;

export function compareAssetPaths(fromAssets: AssetItem[], toAssets: AssetItem[]): ReleaseComparison['assetPaths'];

export function compareReleasePair(
  fromManifest: BenchmarkManifest,
  toManifest: BenchmarkManifest,
  fastCdcIndexes?: { from: FastCdcIndex; to: FastCdcIndex },
): ReleaseComparison;

export function formatBenchmarkMarkdown(report: { generatedAt: string; pairs: ReleaseComparison[] }): string;
