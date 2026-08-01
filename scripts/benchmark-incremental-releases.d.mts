export interface ContentAddressedItem {
  sha256: string;
  bytes: number;
}

export interface BenchmarkFastCdcChunk extends ContentAddressedItem {
  offset: number;
}

export interface BenchmarkAsset extends ContentAddressedItem {
  path: string;
  representations?: {
    whole: ContentAddressedItem;
    fastcdc?: {
      chunks: BenchmarkFastCdcChunk[];
    };
  };
}

export interface BenchmarkManifest {
  version: 4 | 5;
  releaseId: string;
  package?: {
    bytes: number;
    segments: ContentAddressedItem[];
  };
  assets: BenchmarkAsset[];
}

export interface ContentComparison {
  targetBytes: number;
  targetObjects: number;
  downloadBytes: number;
  downloadObjects: number;
  reusedBytes: number;
  reusedObjects: number;
}

export interface AssetPathComparison {
  unchanged: number;
  changed: number;
  added: number;
  removed: number;
}

export type HybridPlannerChange = 'unchanged' | 'changed' | 'added' | 'removed';
export type HybridPlannerStrategy = 'unchanged' | 'whole' | 'fastcdc' | 'removed';

export interface HybridPlannerDecision {
  path: string;
  change: HybridPlannerChange;
  strategy: HybridPlannerStrategy;
  targetBytes: number;
  downloadBytes: number;
  downloadObjects: number;
}

export interface HybridPlannerComparison extends ContentComparison {
  paths: {
    unchanged: number;
    whole: number;
    fastCdc: number;
    removed: number;
  };
  decisions: HybridPlannerDecision[];
}

export interface ReleaseComparison {
  fromReleaseId: string;
  toReleaseId: string;
  fromManifestVersion: 4 | 5;
  toManifestVersion: 4 | 5;
  logicalTargetBytes: number;
  assetPaths: AssetPathComparison;
  currentReleaseBoundSegments: ContentComparison;
  fixedDigestSegments: ContentComparison;
  fileCas: ContentComparison;
  hybridPlanner: HybridPlannerComparison;
  /** Compatibility-only report-v1 field. Whole-package FastCDC is no longer planned. */
  fastCdc: [];
}

export interface IncrementalBenchmarkReportV2 {
  version: 2;
  generatedAt: string;
  origin: string;
  releases: string[];
  manifestVersion: 4 | 5;
  pairs: ReleaseComparison[];
}

export function compareContentAddressedItems(
  fromItems: ContentAddressedItem[],
  toItems: ContentAddressedItem[],
): ContentComparison;

export function compareAssetPaths(fromAssets: BenchmarkAsset[], toAssets: BenchmarkAsset[]): AssetPathComparison;

export function compareHybridPlanner(
  fromManifest: BenchmarkManifest,
  toManifest: BenchmarkManifest,
): HybridPlannerComparison;

export function compareReleasePair(fromManifest: BenchmarkManifest, toManifest: BenchmarkManifest): ReleaseComparison;

export function formatBenchmarkMarkdown(report: Pick<IncrementalBenchmarkReportV2, 'generatedAt' | 'pairs'>): string;

export function releaseManifestPath(releaseId: string, version?: 4 | 5): string;
