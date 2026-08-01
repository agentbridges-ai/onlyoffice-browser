import {
  FASTCDC_V1_POLICY_ID,
  validateReleaseContentModel,
  type AssetContentRepresentation,
  type AssetContentSpan,
  type ContentObjectDescriptor,
  type ReleaseContentAsset,
  type ReleaseContentModel,
} from './release-content-model';
import type { ReleaseAssetV5, ReleaseManifestV5, ReleasePackageSegment } from './release-resources';

const SHA256_PATTERN = /^[a-f0-9]{64}$/;

function addObject(objects: Map<string, ContentObjectDescriptor>, sha256: string, bytes: number): void {
  const existing = objects.get(sha256);
  if (existing && existing.bytes !== bytes) {
    throw new TypeError(`Content digest ${sha256} has conflicting byte lengths`);
  }
  objects.set(sha256, { sha256, bytes });
}

function packageSpans(asset: ReleaseAssetV5, segments: ReleasePackageSegment[]): AssetContentSpan[] {
  const assetStart = asset.packageOffset;
  const assetEnd = assetStart + asset.bytes;
  if (asset.bytes === 0) {
    const container =
      segments.find((segment) => assetStart >= segment.offset && assetStart <= segment.offset + segment.bytes) ||
      segments.at(-1);
    if (!container) throw new TypeError(`Package does not cover empty asset ${asset.path}`);
    return [
      {
        objectSha256: container.sha256,
        objectOffset: assetStart - container.offset,
        assetOffset: 0,
        bytes: 0,
      },
    ];
  }

  const spans: AssetContentSpan[] = [];
  for (const segment of segments) {
    const segmentEnd = segment.offset + segment.bytes;
    const overlapStart = Math.max(assetStart, segment.offset);
    const overlapEnd = Math.min(assetEnd, segmentEnd);
    if (overlapEnd <= overlapStart) continue;
    spans.push({
      objectSha256: segment.sha256,
      objectOffset: overlapStart - segment.offset,
      assetOffset: overlapStart - assetStart,
      bytes: overlapEnd - overlapStart,
    });
  }
  return spans;
}

function assetRepresentations(asset: ReleaseAssetV5, segments: ReleasePackageSegment[]): AssetContentRepresentation[] {
  const representations: AssetContentRepresentation[] = [
    {
      id: 'whole',
      kind: 'whole',
      spans: [
        {
          objectSha256: asset.representations.whole.sha256,
          objectOffset: 0,
          assetOffset: 0,
          bytes: asset.representations.whole.bytes,
        },
      ],
    },
    {
      id: 'package',
      kind: 'package',
      spans: packageSpans(asset, segments),
    },
  ];
  const fastcdc = asset.representations.fastcdc;
  if (fastcdc) {
    representations.push({
      id: FASTCDC_V1_POLICY_ID,
      kind: 'fastcdc',
      fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
      spans: fastcdc.chunks.map((chunk) => ({
        objectSha256: chunk.sha256,
        objectOffset: 0,
        assetOffset: chunk.offset,
        bytes: chunk.bytes,
      })),
    });
  }
  return representations;
}

/**
 * Converts an already parsed v5 release manifest into the representation graph
 * used by the transactional planner. No bytes are copied: package spans point
 * directly into the immutable package segment objects.
 */
export function releaseManifestV5ToContentModel(
  manifest: ReleaseManifestV5,
  manifestSha256: string,
): ReleaseContentModel {
  if (!SHA256_PATTERN.test(manifestSha256)) throw new TypeError('Invalid manifest SHA-256');
  const objects = new Map<string, ContentObjectDescriptor>();
  for (const segment of manifest.package.segments) addObject(objects, segment.sha256, segment.bytes);

  const assets: ReleaseContentAsset[] = manifest.assets.map((asset) => {
    addObject(objects, asset.representations.whole.sha256, asset.representations.whole.bytes);
    for (const chunk of asset.representations.fastcdc?.chunks || []) {
      addObject(objects, chunk.sha256, chunk.bytes);
    }
    return {
      path: asset.path,
      bytes: asset.bytes,
      sha256: asset.sha256,
      mime: asset.mime,
      representations: assetRepresentations(asset, manifest.package.segments),
    };
  });
  const model: ReleaseContentModel = {
    releaseId: manifest.releaseId,
    manifestSha256,
    storageSetSha256: manifest.contentProtocol.storageSetSha256,
    objects: [...objects.values()].sort((left, right) => left.sha256.localeCompare(right.sha256)),
    assets,
  };
  validateReleaseContentModel(model);
  return model;
}
