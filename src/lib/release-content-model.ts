import { sha256 } from '@noble/hashes/sha2.js';

export const FASTCDC_V1_POLICY_ID = 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0' as const;

export type AssetRepresentationKind = 'whole' | 'fastcdc' | 'package';
export type ReleaseTransactionState = 'installing' | 'prepared' | 'active' | 'retained' | 'failed';

export type ReleaseContentModelErrorCode = 'invalid-model' | 'integrity-conflict' | 'invalid-state' | 'incomplete';

export class ReleaseContentModelError extends Error {
  constructor(
    readonly code: ReleaseContentModelErrorCode,
    message: string,
  ) {
    super(message);
    this.name = 'ReleaseContentModelError';
  }
}

export interface ContentObjectDescriptor {
  sha256: string;
  bytes: number;
}

export interface ContentObjectRecord extends ContentObjectDescriptor {
  verifiedAt: number;
}

export interface AssetContentSpan {
  objectSha256: string;
  objectOffset: number;
  assetOffset: number;
  bytes: number;
}

export interface AssetContentRepresentation {
  id: string;
  kind: AssetRepresentationKind;
  spans: AssetContentSpan[];
  fastCdcPolicyId?: typeof FASTCDC_V1_POLICY_ID;
}

export interface ReleaseContentAsset {
  path: string;
  bytes: number;
  sha256: string;
  mime?: string;
  representations: AssetContentRepresentation[];
}

export interface ReleaseContentModel {
  releaseId: string;
  manifestSha256: string;
  storageSetSha256: string;
  objects: ContentObjectDescriptor[];
  assets: ReleaseContentAsset[];
}

export interface AssetContentMapping {
  releaseId: string;
  path: string;
  assetSha256: string;
  assetBytes: number;
  mime?: string;
  representationId: string;
  representationKind: AssetRepresentationKind;
  spans: AssetContentSpan[];
  sourceReleaseId?: string;
}

export interface RepresentationCost {
  representationId: string;
  kind: AssetRepresentationKind;
  referencedBytes: number;
  incrementalRequiredBytes: number;
  downloadBytes: number;
  availableBytes: number;
  alreadyPlannedBytes: number;
  requiredObjects: ContentObjectDescriptor[];
  missingObjects: ContentObjectDescriptor[];
}

export interface AssetRepresentationCosts {
  path: string;
  representations: RepresentationCost[];
}

export interface ReleaseContentPlan {
  releaseId: string;
  totalLogicalBytes: number;
  requiredObjectBytes: number;
  downloadBytes: number;
  reusedObjectBytes: number;
  requiredObjects: ContentObjectDescriptor[];
  missingObjects: ContentObjectDescriptor[];
  mappings: AssetContentMapping[];
  representationCosts: AssetRepresentationCosts[];
}

export interface ReleaseTransactionRecord {
  releaseId: string;
  manifestSha256: string;
  storageSetSha256: string;
  planFingerprint: string;
  state: ReleaseTransactionState;
  previousActiveReleaseId: string | null;
  requiredObjects: ContentObjectDescriptor[];
  plannedMappings: AssetContentMapping[];
  committedMappings: AssetContentMapping[];
  activationRollbackState?: 'failed' | 'retained';
  failureCode?: string;
  updatedAt: number;
}

export interface ReleaseContentJournalSnapshot {
  version: 1;
  activeReleaseId: string | null;
  objects: ContentObjectRecord[];
  transactions: ReleaseTransactionRecord[];
}

export interface ReleaseResumeState {
  releaseId: string;
  state: ReleaseTransactionState;
  missingObjects: ContentObjectDescriptor[];
  pendingAssetPaths: string[];
}

/**
 * A transactionally consistent, release-scoped read view for runtime health
 * checks and Broker reads. `objects` contains only records referenced by the
 * selected release, so callers never need to scan the global object ledger.
 */
export interface ReleaseContentReadView {
  transaction: ReleaseTransactionRecord | null;
  objects: ContentObjectRecord[];
}

type ObjectMap<T extends ContentObjectDescriptor = ContentObjectDescriptor> = ReadonlyMap<string, T>;

const DIGEST_PATTERN = /^[a-f0-9]{64}$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const MIME_PATTERN =
  /^[a-zA-Z0-9][a-zA-Z0-9!#$&^_.+-]*\/[a-zA-Z0-9][a-zA-Z0-9!#$&^_.+-]*(?:\s*;\s*[a-zA-Z0-9!#$&^_.+-]+\s*=\s*[a-zA-Z0-9!#$&^_.+-]+)*$/;

function fail(code: ReleaseContentModelErrorCode, message: string): never {
  throw new ReleaseContentModelError(code, message);
}

function isSafeInteger(value: number, minimum = 0): boolean {
  return Number.isSafeInteger(value) && value >= minimum;
}

function assertDigest(value: string, label: string): void {
  if (!DIGEST_PATTERN.test(value)) fail('invalid-model', `${label} must be a lowercase SHA-256 digest`);
}

function assertSafePath(path: string): void {
  if (
    path.length === 0 ||
    path.length > 2_048 ||
    path.startsWith('/') ||
    path.includes('\\') ||
    path.includes('\0') ||
    path.includes('?') ||
    path.includes('#') ||
    /%(?:2e|2f|5c)/i.test(path) ||
    path.split('/').some((part) => part.length === 0 || part === '.' || part === '..')
  ) {
    fail('invalid-model', `unsafe asset path: ${path}`);
  }
}

function assertOptionalMime(mime: string | undefined, label: string): void {
  if (mime !== undefined && (typeof mime !== 'string' || mime.length > 256 || !MIME_PATTERN.test(mime))) {
    fail('invalid-model', `${label} must be a safe non-empty MIME type`);
  }
}

function descriptorMap(
  descriptors: Iterable<ContentObjectDescriptor>,
  errorCode: ReleaseContentModelErrorCode = 'invalid-model',
): Map<string, ContentObjectDescriptor> {
  const result = new Map<string, ContentObjectDescriptor>();
  for (const descriptor of descriptors) {
    assertDigest(descriptor.sha256, 'object sha256');
    if (!isSafeInteger(descriptor.bytes)) fail(errorCode, `invalid object size for ${descriptor.sha256}`);
    const existing = result.get(descriptor.sha256);
    if (existing && existing.bytes !== descriptor.bytes) {
      fail('integrity-conflict', `digest ${descriptor.sha256} is associated with conflicting sizes`);
    }
    result.set(descriptor.sha256, { sha256: descriptor.sha256, bytes: descriptor.bytes });
  }
  return result;
}

function recordMap(records: Iterable<ContentObjectRecord>): Map<string, ContentObjectRecord> {
  const materialized = [...records];
  const descriptors = descriptorMap(materialized, 'integrity-conflict');
  const result = new Map<string, ContentObjectRecord>();
  for (const record of materialized) {
    if (!Number.isFinite(record.verifiedAt) || record.verifiedAt < 0) {
      fail('invalid-model', `invalid verification timestamp for ${record.sha256}`);
    }
    const descriptor = descriptors.get(record.sha256)!;
    result.set(record.sha256, { ...descriptor, verifiedAt: record.verifiedAt });
  }
  return result;
}

function cloneSpan(span: AssetContentSpan): AssetContentSpan {
  return { ...span };
}

function cloneDescriptor(descriptor: ContentObjectDescriptor): ContentObjectDescriptor {
  return { ...descriptor };
}

function cloneRecord(record: ContentObjectRecord): ContentObjectRecord {
  return { ...record };
}

function cloneMapping(mapping: AssetContentMapping): AssetContentMapping {
  return {
    ...mapping,
    spans: mapping.spans.map(cloneSpan),
  };
}

function cloneTransaction(transaction: ReleaseTransactionRecord): ReleaseTransactionRecord {
  return {
    ...transaction,
    requiredObjects: transaction.requiredObjects.map(cloneDescriptor),
    plannedMappings: transaction.plannedMappings.map(cloneMapping),
    committedMappings: transaction.committedMappings.map(cloneMapping),
  };
}

function compareDescriptors(left: ContentObjectDescriptor, right: ContentObjectDescriptor): number {
  return left.sha256.localeCompare(right.sha256);
}

function compareMappings(left: AssetContentMapping, right: AssetContentMapping): number {
  return (
    left.path.localeCompare(right.path) ||
    left.representationKind.localeCompare(right.representationKind) ||
    left.representationId.localeCompare(right.representationId)
  );
}

function mappingIdentity(mapping: AssetContentMapping): string {
  return JSON.stringify({
    path: mapping.path,
    assetSha256: mapping.assetSha256,
    assetBytes: mapping.assetBytes,
    mime: mapping.mime,
    representationId: mapping.representationId,
    representationKind: mapping.representationKind,
    spans: mapping.spans,
  });
}

function digestText(value: string): string {
  return Array.from(sha256(new TextEncoder().encode(value)), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function fingerprintPlanParts(
  releaseId: string,
  requiredObjects: ContentObjectDescriptor[],
  mappings: AssetContentMapping[],
): string {
  return digestText(
    JSON.stringify({
      releaseId,
      objects: [...requiredObjects].sort(compareDescriptors),
      mappings: [...mappings].sort(compareMappings).map((mapping) => ({
        path: mapping.path,
        assetSha256: mapping.assetSha256,
        assetBytes: mapping.assetBytes,
        mime: mapping.mime,
        representationId: mapping.representationId,
        representationKind: mapping.representationKind,
        spans: mapping.spans,
      })),
    }),
  );
}

function fingerprintPlan(plan: ReleaseContentPlan): string {
  return fingerprintPlanParts(plan.releaseId, plan.requiredObjects, plan.mappings);
}

export function validateAssetSpanCoverage(
  asset: Pick<ReleaseContentAsset, 'path' | 'bytes' | 'sha256'>,
  representation: AssetContentRepresentation,
  objects: ObjectMap,
): void {
  assertSafePath(asset.path);
  assertDigest(asset.sha256, `asset ${asset.path} sha256`);
  if (!isSafeInteger(asset.bytes)) fail('invalid-model', `invalid asset size for ${asset.path}`);
  if (!representation.id || representation.id.length > 128) {
    fail('invalid-model', `invalid representation id for ${asset.path}`);
  }
  if (!['whole', 'fastcdc', 'package'].includes(representation.kind)) {
    fail('invalid-model', `invalid representation kind for ${asset.path}`);
  }
  if (representation.kind === 'fastcdc' && representation.fastCdcPolicyId !== FASTCDC_V1_POLICY_ID) {
    fail('invalid-model', `unsupported FastCDC policy for ${asset.path}`);
  }
  if (representation.kind !== 'fastcdc' && representation.fastCdcPolicyId !== undefined) {
    fail('invalid-model', `unexpected FastCDC policy for ${asset.path}`);
  }
  if (asset.bytes === 0 && representation.spans.length !== 1) {
    fail('invalid-model', `empty asset ${asset.path} must retain one content-addressed object`);
  }
  if (asset.bytes > 0 && representation.spans.length === 0) {
    fail('invalid-model', `representation ${representation.id} does not cover ${asset.path}`);
  }

  let expectedAssetOffset = 0;
  for (const span of representation.spans) {
    assertDigest(span.objectSha256, `span object for ${asset.path}`);
    const object = objects.get(span.objectSha256);
    if (!object) fail('invalid-model', `span for ${asset.path} references an undeclared object`);
    if (
      !isSafeInteger(span.objectOffset) ||
      !isSafeInteger(span.assetOffset) ||
      !isSafeInteger(span.bytes) ||
      (asset.bytes > 0 && span.bytes === 0)
    ) {
      fail('invalid-model', `invalid span bounds for ${asset.path}`);
    }
    if (span.assetOffset !== expectedAssetOffset) {
      fail('invalid-model', `span coverage for ${asset.path} has a gap or overlap`);
    }
    if (span.objectOffset + span.bytes > object.bytes || span.assetOffset + span.bytes > asset.bytes) {
      fail('invalid-model', `span for ${asset.path} exceeds its object or asset bounds`);
    }
    expectedAssetOffset += span.bytes;
  }
  if (expectedAssetOffset !== asset.bytes) {
    fail('invalid-model', `span coverage for ${asset.path} is incomplete`);
  }

  if (representation.kind === 'whole') {
    const span = representation.spans[0];
    const object = objects.get(span.objectSha256)!;
    if (
      representation.spans.length !== 1 ||
      span.objectSha256 !== asset.sha256 ||
      span.objectOffset !== 0 ||
      span.assetOffset !== 0 ||
      span.bytes !== asset.bytes ||
      object.bytes !== asset.bytes
    ) {
      fail('invalid-model', `whole representation for ${asset.path} is not the complete asset object`);
    }
  }

  if (
    representation.kind === 'fastcdc' &&
    representation.spans.some((span) => span.objectOffset !== 0 || objects.get(span.objectSha256)!.bytes !== span.bytes)
  ) {
    fail('invalid-model', `FastCDC representation for ${asset.path} must reference complete chunk objects`);
  }
}

export function validateReleaseContentModel(model: ReleaseContentModel): void {
  if (!RELEASE_ID_PATTERN.test(model.releaseId)) fail('invalid-model', 'invalid release id');
  assertDigest(model.manifestSha256, 'manifest sha256');
  assertDigest(model.storageSetSha256, 'storage set sha256');
  const objects = descriptorMap(model.objects);
  const paths = new Set<string>();
  for (const asset of model.assets) {
    if (paths.has(asset.path)) fail('invalid-model', `duplicate asset path: ${asset.path}`);
    paths.add(asset.path);
    assertOptionalMime(asset.mime, `asset ${asset.path} MIME`);
    if (asset.representations.length === 0) {
      fail('invalid-model', `asset ${asset.path} has no content representation`);
    }
    const representationIds = new Set<string>();
    for (const representation of asset.representations) {
      if (representationIds.has(representation.id)) {
        fail('invalid-model', `duplicate representation id ${representation.id} for ${asset.path}`);
      }
      representationIds.add(representation.id);
      validateAssetSpanCoverage(asset, representation, objects);
    }
  }
}

export function estimateRepresentationCost(
  representation: AssetContentRepresentation,
  objects: ObjectMap,
  availableObjects: ObjectMap<ContentObjectRecord>,
  alreadyRequired: ReadonlySet<string> = new Set(),
): RepresentationCost {
  const referenced = new Map<string, ContentObjectDescriptor>();
  for (const span of representation.spans) {
    const object = objects.get(span.objectSha256);
    if (!object) fail('invalid-model', `representation ${representation.id} references an undeclared object`);
    const available = availableObjects.get(object.sha256);
    if (available && available.bytes !== object.bytes) {
      fail('integrity-conflict', `stored object ${object.sha256} has an unexpected size`);
    }
    referenced.set(object.sha256, object);
  }

  const requiredObjects = [...referenced.values()].sort(compareDescriptors);
  const missingObjects = requiredObjects.filter(
    (object) => !availableObjects.has(object.sha256) && !alreadyRequired.has(object.sha256),
  );
  return {
    representationId: representation.id,
    kind: representation.kind,
    referencedBytes: requiredObjects.reduce((sum, object) => sum + object.bytes, 0),
    incrementalRequiredBytes: requiredObjects
      .filter((object) => !alreadyRequired.has(object.sha256))
      .reduce((sum, object) => sum + object.bytes, 0),
    downloadBytes: missingObjects.reduce((sum, object) => sum + object.bytes, 0),
    availableBytes: requiredObjects
      .filter((object) => availableObjects.has(object.sha256))
      .reduce((sum, object) => sum + object.bytes, 0),
    alreadyPlannedBytes: requiredObjects
      .filter((object) => alreadyRequired.has(object.sha256))
      .reduce((sum, object) => sum + object.bytes, 0),
    requiredObjects,
    missingObjects,
  };
}

function representationRank(kind: AssetRepresentationKind): number {
  if (kind === 'whole') return 0;
  if (kind === 'fastcdc') return 1;
  return 2;
}

function compareCosts(left: RepresentationCost, right: RepresentationCost): number {
  return (
    left.downloadBytes - right.downloadBytes ||
    left.missingObjects.length - right.missingObjects.length ||
    left.incrementalRequiredBytes - right.incrementalRequiredBytes ||
    left.requiredObjects.length - right.requiredObjects.length ||
    representationRank(left.kind) - representationRank(right.kind) ||
    left.representationId.localeCompare(right.representationId)
  );
}

function createMapping(
  releaseId: string,
  asset: ReleaseContentAsset,
  representation: AssetContentRepresentation,
  sourceReleaseId?: string,
): AssetContentMapping {
  return {
    releaseId,
    path: asset.path,
    assetSha256: asset.sha256,
    assetBytes: asset.bytes,
    mime: asset.mime,
    representationId: representation.id,
    representationKind: representation.kind,
    spans: representation.spans.map(cloneSpan),
    ...(sourceReleaseId && sourceReleaseId !== releaseId ? { sourceReleaseId } : {}),
  };
}

function selectPreviousMapping(
  releaseId: string,
  asset: ReleaseContentAsset,
  previousMappings: AssetContentMapping[],
  availableObjects: ObjectMap<ContentObjectRecord>,
): AssetContentMapping | null {
  const candidates = previousMappings
    .filter(
      (mapping) =>
        mapping.assetSha256 === asset.sha256 &&
        mapping.assetBytes === asset.bytes &&
        mapping.spans.every((span) => availableObjects.has(span.objectSha256)),
    )
    .sort(
      (left, right) =>
        left.releaseId.localeCompare(right.releaseId) ||
        left.path.localeCompare(right.path) ||
        left.representationId.localeCompare(right.representationId),
    );
  for (const previous of candidates) {
    validateMapping(previous, availableObjects);
    const previousRepresentation: AssetContentRepresentation = {
      id: previous.representationId,
      kind: previous.representationKind,
      spans: previous.spans.map(cloneSpan),
      ...(previous.representationKind === 'fastcdc' ? { fastCdcPolicyId: FASTCDC_V1_POLICY_ID } : {}),
    };
    validateAssetSpanCoverage(asset, previousRepresentation, availableObjects);
    return {
      ...cloneMapping(previous),
      releaseId,
      path: asset.path,
      assetSha256: asset.sha256,
      assetBytes: asset.bytes,
      mime: asset.mime,
      sourceReleaseId: previous.releaseId,
    };
  }
  return null;
}

type PlanVariant = {
  mappings: AssetContentMapping[];
  requiredObjects: Map<string, ContentObjectDescriptor>;
};

function buildPlanVariant(
  model: ReleaseContentModel,
  objectMap: Map<string, ContentObjectDescriptor>,
  availableObjects: Map<string, ContentObjectRecord>,
  aliases: Map<string, AssetContentMapping>,
  preferredKind?: AssetRepresentationKind,
): PlanVariant {
  const mappings: AssetContentMapping[] = [];
  const requiredObjects = new Map<string, ContentObjectDescriptor>();
  const sortedAssets = [...model.assets].sort((left, right) => left.path.localeCompare(right.path));

  for (const asset of sortedAssets) {
    const alias = aliases.get(asset.path);
    if (alias) {
      mappings.push(cloneMapping(alias));
      for (const span of alias.spans) requiredObjects.set(span.objectSha256, objectMap.get(span.objectSha256)!);
      continue;
    }

    const preferred = preferredKind
      ? asset.representations.filter((representation) => representation.kind === preferredKind)
      : [];
    const candidates = preferred.length > 0 ? preferred : asset.representations;
    const costs = candidates
      .map((representation) => ({
        representation,
        cost: estimateRepresentationCost(representation, objectMap, availableObjects, new Set(requiredObjects.keys())),
      }))
      .sort((left, right) => compareCosts(left.cost, right.cost));
    const selected = costs[0].representation;
    mappings.push(createMapping(model.releaseId, asset, selected));
    for (const descriptor of costs[0].cost.requiredObjects) requiredObjects.set(descriptor.sha256, descriptor);
  }

  return { mappings, requiredObjects };
}

function buildColdCanonicalPlanVariant(
  model: ReleaseContentModel,
  objectMap: Map<string, ContentObjectDescriptor>,
  availableObjects: Map<string, ContentObjectRecord>,
): PlanVariant {
  const mappings: AssetContentMapping[] = [];
  const requiredObjects = new Map<string, ContentObjectDescriptor>();
  const sortedAssets = [...model.assets].sort((left, right) => left.path.localeCompare(right.path));
  for (const asset of sortedAssets) {
    const fastCdc = asset.representations.filter((representation) => representation.kind === 'fastcdc');
    const canonical = fastCdc.length
      ? fastCdc
      : asset.representations.filter((representation) => representation.kind === 'whole');
    if (canonical.length === 0) fail('invalid-model', `asset ${asset.path} has no canonical storage representation`);
    const selected = canonical
      .map((representation) => ({
        representation,
        cost: estimateRepresentationCost(representation, objectMap, availableObjects, new Set(requiredObjects.keys())),
      }))
      .sort((left, right) => compareCosts(left.cost, right.cost))[0];
    mappings.push(createMapping(model.releaseId, asset, selected.representation));
    for (const descriptor of selected.cost.requiredObjects) requiredObjects.set(descriptor.sha256, descriptor);
  }
  return { mappings, requiredObjects };
}

function variantSortKey(variant: PlanVariant, availableObjects: ObjectMap<ContentObjectRecord>): readonly unknown[] {
  const required = [...variant.requiredObjects.values()];
  const missing = required.filter((object) => !availableObjects.has(object.sha256));
  return [
    missing.reduce((sum, object) => sum + object.bytes, 0),
    missing.length,
    required.reduce((sum, object) => sum + object.bytes, 0),
    required.length,
    variant.mappings.reduce((sum, mapping) => sum + representationRank(mapping.representationKind), 0),
    variant.mappings
      .map((mapping) => `${mapping.path}:${mapping.representationKind}:${mapping.representationId}`)
      .join('|'),
  ];
}

function compareVariantKeys(left: readonly unknown[], right: readonly unknown[]): number {
  for (let index = 0; index < left.length; index += 1) {
    const leftValue = left[index];
    const rightValue = right[index];
    const comparison =
      typeof leftValue === 'number' && typeof rightValue === 'number'
        ? leftValue - rightValue
        : String(leftValue).localeCompare(String(rightValue));
    if (comparison !== 0) return comparison;
  }
  return 0;
}

export function planReleaseContent(
  model: ReleaseContentModel,
  inventory: Iterable<ContentObjectRecord>,
  previousMappings: Iterable<AssetContentMapping> = [],
  options: {
    preferPackageForColdInstall?: boolean;
    preferWholeForColdInstall?: boolean;
    preferCanonicalForColdInstall?: boolean;
  } = {},
): ReleaseContentPlan {
  validateReleaseContentModel(model);
  const targetObjects = descriptorMap(model.objects);
  const availableObjects = recordMap(inventory);
  for (const [digest, available] of availableObjects) {
    const expected = targetObjects.get(digest);
    if (expected && expected.bytes !== available.bytes) {
      fail('integrity-conflict', `stored object ${digest} conflicts with the target release`);
    }
  }

  const previous = [...previousMappings].map(cloneMapping);
  const aliases = new Map<string, AssetContentMapping>();
  for (const asset of model.assets) {
    const alias = selectPreviousMapping(model.releaseId, asset, previous, availableObjects);
    if (alias) aliases.set(asset.path, alias);
  }
  const planningObjects = new Map(targetObjects);
  for (const alias of aliases.values()) {
    for (const span of alias.spans) {
      const object = availableObjects.get(span.objectSha256)!;
      planningObjects.set(object.sha256, { sha256: object.sha256, bytes: object.bytes });
    }
  }

  const coldPreferences = [
    options.preferPackageForColdInstall,
    options.preferWholeForColdInstall,
    options.preferCanonicalForColdInstall,
  ].filter(Boolean).length;
  if (coldPreferences > 1) {
    fail('invalid-state', 'cold install cannot prefer multiple representation policies');
  }
  const coldPreferredKind = options.preferPackageForColdInstall
    ? 'package'
    : options.preferWholeForColdInstall
      ? 'whole'
      : undefined;
  const variants = options.preferCanonicalForColdInstall
    ? [buildColdCanonicalPlanVariant(model, planningObjects, availableObjects)]
    : coldPreferredKind
      ? [buildPlanVariant(model, planningObjects, availableObjects, new Map(), coldPreferredKind)]
      : [aliases, new Map<string, AssetContentMapping>()]
          .flatMap((candidateAliases) =>
            ([undefined, 'whole', 'fastcdc', 'package'] as const).map((preferredKind) =>
              buildPlanVariant(model, planningObjects, availableObjects, candidateAliases, preferredKind),
            ),
          )
          .sort((left, right) =>
            compareVariantKeys(variantSortKey(left, availableObjects), variantSortKey(right, availableObjects)),
          );
  const selected = variants[0];
  const requiredObjects = [...selected.requiredObjects.values()].sort(compareDescriptors);
  const missingObjects = requiredObjects.filter((object) => !availableObjects.has(object.sha256));
  const representationCosts = [...model.assets]
    .sort((left, right) => left.path.localeCompare(right.path))
    .map((asset) => ({
      path: asset.path,
      representations: asset.representations
        .map((representation) => estimateRepresentationCost(representation, targetObjects, availableObjects))
        .sort(compareCosts),
    }));

  return {
    releaseId: model.releaseId,
    totalLogicalBytes: model.assets.reduce((sum, asset) => sum + asset.bytes, 0),
    requiredObjectBytes: requiredObjects.reduce((sum, object) => sum + object.bytes, 0),
    downloadBytes: missingObjects.reduce((sum, object) => sum + object.bytes, 0),
    reusedObjectBytes: requiredObjects
      .filter((object) => availableObjects.has(object.sha256))
      .reduce((sum, object) => sum + object.bytes, 0),
    requiredObjects,
    missingObjects,
    mappings: selected.mappings.sort(compareMappings),
    representationCosts,
  };
}

function validateMapping(mapping: AssetContentMapping, objects: ObjectMap): void {
  if (!RELEASE_ID_PATTERN.test(mapping.releaseId)) fail('invalid-model', 'invalid mapping release id');
  if (mapping.sourceReleaseId && !RELEASE_ID_PATTERN.test(mapping.sourceReleaseId)) {
    fail('invalid-model', 'invalid mapping source release id');
  }
  assertOptionalMime(mapping.mime, `mapping ${mapping.path} MIME`);
  validateAssetSpanCoverage(
    {
      path: mapping.path,
      bytes: mapping.assetBytes,
      sha256: mapping.assetSha256,
    },
    {
      id: mapping.representationId,
      kind: mapping.representationKind,
      spans: mapping.spans,
      ...(mapping.representationKind === 'fastcdc' ? { fastCdcPolicyId: FASTCDC_V1_POLICY_ID } : {}),
    },
    objects,
  );
}

function validateTransaction(transaction: ReleaseTransactionRecord, objects: ObjectMap<ContentObjectRecord>): void {
  if (!RELEASE_ID_PATTERN.test(transaction.releaseId)) fail('invalid-model', 'invalid transaction release id');
  assertDigest(transaction.manifestSha256, 'transaction manifest sha256');
  assertDigest(transaction.storageSetSha256, 'transaction storage set sha256');
  assertDigest(transaction.planFingerprint, 'transaction plan fingerprint');
  if (!['installing', 'prepared', 'active', 'retained', 'failed'].includes(transaction.state)) {
    fail('invalid-model', `invalid transaction state for ${transaction.releaseId}`);
  }
  if (transaction.activationRollbackState !== undefined && transaction.state !== 'active') {
    fail('invalid-model', `non-active transaction ${transaction.releaseId} retains activation rollback state`);
  }
  if (transaction.previousActiveReleaseId !== null && !RELEASE_ID_PATTERN.test(transaction.previousActiveReleaseId)) {
    fail('invalid-model', 'invalid previous active release id');
  }
  if (!Number.isFinite(transaction.updatedAt) || transaction.updatedAt < 0) {
    fail('invalid-model', `invalid transaction timestamp for ${transaction.releaseId}`);
  }

  const requiredObjects = descriptorMap(transaction.requiredObjects);
  for (const object of requiredObjects.values()) {
    const stored = objects.get(object.sha256);
    if (stored && stored.bytes !== object.bytes) {
      fail('integrity-conflict', `stored object ${object.sha256} conflicts with transaction ${transaction.releaseId}`);
    }
  }
  const plannedPaths = new Set<string>();
  for (const mapping of transaction.plannedMappings) {
    if (mapping.releaseId !== transaction.releaseId || plannedPaths.has(mapping.path)) {
      fail('invalid-model', `invalid planned mapping for ${transaction.releaseId}`);
    }
    plannedPaths.add(mapping.path);
    validateMapping(mapping, requiredObjects);
  }
  const plannedIdentities = new Map(
    transaction.plannedMappings.map((mapping) => [mapping.path, mappingIdentity(mapping)]),
  );
  const committedPaths = new Set<string>();
  for (const mapping of transaction.committedMappings) {
    if (committedPaths.has(mapping.path) || plannedIdentities.get(mapping.path) !== mappingIdentity(mapping)) {
      fail('invalid-model', `committed mapping does not match the release plan for ${mapping.path}`);
    }
    committedPaths.add(mapping.path);
  }

  const complete =
    [...requiredObjects.values()].every((object) => objects.get(object.sha256)?.bytes === object.bytes) &&
    committedPaths.size === plannedPaths.size;
  if (['prepared', 'active', 'retained'].includes(transaction.state) && !complete) {
    fail('incomplete', `transaction ${transaction.releaseId} is marked complete without all content`);
  }
  if (transaction.state === 'failed' && !transaction.failureCode) {
    fail('invalid-model', `failed transaction ${transaction.releaseId} has no failure code`);
  }
  if (transaction.state !== 'failed' && transaction.failureCode !== undefined) {
    fail('invalid-model', `non-failed transaction ${transaction.releaseId} retains a failure code`);
  }
  if (
    transaction.planFingerprint !==
    fingerprintPlanParts(transaction.releaseId, transaction.requiredObjects, transaction.plannedMappings)
  ) {
    fail('integrity-conflict', `transaction ${transaction.releaseId} has a mismatched plan fingerprint`);
  }
}

export class MemoryReleaseContentJournal {
  private activeReleaseId: string | null;
  private readonly objects: Map<string, ContentObjectRecord>;
  private readonly transactions: Map<string, ReleaseTransactionRecord>;
  private readonly now: () => number;

  constructor(options: { snapshot?: ReleaseContentJournalSnapshot; now?: () => number } = {}) {
    this.now = options.now || Date.now;
    const snapshot = options.snapshot || {
      version: 1 as const,
      activeReleaseId: null,
      objects: [],
      transactions: [],
    };
    if (snapshot.version !== 1) fail('invalid-model', 'unsupported release content journal version');
    this.objects = recordMap(snapshot.objects);
    this.transactions = new Map();
    for (const transaction of snapshot.transactions) {
      if (this.transactions.has(transaction.releaseId)) {
        fail('invalid-model', `duplicate transaction for ${transaction.releaseId}`);
      }
      validateTransaction(transaction, this.objects);
      this.transactions.set(transaction.releaseId, cloneTransaction(transaction));
    }
    this.activeReleaseId = snapshot.activeReleaseId;
    const activeTransactions = [...this.transactions.values()].filter((transaction) => transaction.state === 'active');
    if (
      (this.activeReleaseId === null && activeTransactions.length !== 0) ||
      (this.activeReleaseId !== null &&
        (activeTransactions.length !== 1 || activeTransactions[0].releaseId !== this.activeReleaseId))
    ) {
      fail('invalid-state', 'active release pointer and transaction state disagree');
    }
  }

  static fromSnapshot(
    snapshot: ReleaseContentJournalSnapshot,
    now: () => number = Date.now,
  ): MemoryReleaseContentJournal {
    return new MemoryReleaseContentJournal({ snapshot, now });
  }

  beginRelease(
    plan: ReleaseContentPlan,
    identity: { manifestSha256: string; storageSetSha256: string },
  ): ReleaseTransactionRecord {
    if (!RELEASE_ID_PATTERN.test(plan.releaseId)) fail('invalid-model', 'invalid plan release id');
    assertDigest(identity.manifestSha256, 'manifest sha256');
    assertDigest(identity.storageSetSha256, 'storage set sha256');
    const requiredObjects = descriptorMap(plan.requiredObjects);
    const plannedMappings = plan.mappings.map(cloneMapping).sort(compareMappings);
    const paths = new Set<string>();
    for (const mapping of plannedMappings) {
      if (mapping.releaseId !== plan.releaseId || paths.has(mapping.path)) {
        fail('invalid-model', `invalid or duplicate mapping for ${mapping.path}`);
      }
      paths.add(mapping.path);
      validateMapping(mapping, requiredObjects);
    }
    const planFingerprint = fingerprintPlan(plan);
    const existing = this.transactions.get(plan.releaseId);
    if (existing) {
      if (
        existing.manifestSha256 !== identity.manifestSha256 ||
        existing.storageSetSha256 !== identity.storageSetSha256 ||
        existing.planFingerprint !== planFingerprint
      ) {
        fail('integrity-conflict', `release ${plan.releaseId} was already journaled with different content`);
      }
      return cloneTransaction(existing);
    }

    const transaction: ReleaseTransactionRecord = {
      releaseId: plan.releaseId,
      manifestSha256: identity.manifestSha256,
      storageSetSha256: identity.storageSetSha256,
      planFingerprint,
      state: 'installing',
      previousActiveReleaseId: this.activeReleaseId,
      requiredObjects: [...requiredObjects.values()].sort(compareDescriptors),
      plannedMappings,
      committedMappings: [],
      updatedAt: this.now(),
    };
    this.transactions.set(transaction.releaseId, transaction);
    return cloneTransaction(transaction);
  }

  recordVerifiedObject(releaseId: string, object: ContentObjectDescriptor, verifiedAt: number = this.now()): void {
    const transaction = this.requireMutableTransaction(releaseId);
    const expected = transaction.requiredObjects.find((candidate) => candidate.sha256 === object.sha256);
    if (!expected || expected.bytes !== object.bytes) {
      fail('integrity-conflict', `object ${object.sha256} does not match the plan for ${releaseId}`);
    }
    if (!Number.isFinite(verifiedAt) || verifiedAt < 0) fail('invalid-model', 'invalid verification timestamp');
    const existing = this.objects.get(object.sha256);
    if (existing && existing.bytes !== object.bytes) {
      fail('integrity-conflict', `object ${object.sha256} conflicts with the verified object inventory`);
    }
    this.objects.set(object.sha256, { ...object, verifiedAt: existing?.verifiedAt || verifiedAt });
    transaction.updatedAt = this.now();
  }

  commitAssetMapping(releaseId: string, path: string): AssetContentMapping {
    const transaction = this.requireMutableTransaction(releaseId);
    const mapping = transaction.plannedMappings.find((candidate) => candidate.path === path);
    if (!mapping) fail('invalid-model', `asset ${path} is not part of release ${releaseId}`);
    for (const span of mapping.spans) {
      const expected = transaction.requiredObjects.find((object) => object.sha256 === span.objectSha256)!;
      if (this.objects.get(span.objectSha256)?.bytes !== expected.bytes) {
        fail('incomplete', `asset ${path} references an object that is not verified`);
      }
    }
    if (!transaction.committedMappings.some((candidate) => candidate.path === path)) {
      transaction.committedMappings.push(cloneMapping(mapping));
      transaction.committedMappings.sort(compareMappings);
      transaction.updatedAt = this.now();
    }
    return cloneMapping(mapping);
  }

  prepareRelease(releaseId: string): ReleaseTransactionRecord {
    const transaction = this.requireMutableTransaction(releaseId);
    const resume = this.getResumeState(releaseId);
    if (resume.missingObjects.length > 0 || resume.pendingAssetPaths.length > 0) {
      fail('incomplete', `release ${releaseId} is not completely verified and mapped`);
    }
    transaction.state = 'prepared';
    transaction.updatedAt = this.now();
    return cloneTransaction(transaction);
  }

  activateRelease(releaseId: string): ReleaseTransactionRecord {
    const transaction = this.requireTransaction(releaseId);
    if (transaction.state !== 'prepared') {
      fail('invalid-state', `only a prepared release can be activated`);
    }
    if (this.activeReleaseId !== transaction.previousActiveReleaseId) {
      fail('invalid-state', `the active release changed while ${releaseId} was being installed`);
    }
    if (this.activeReleaseId) {
      const previous = this.requireTransaction(this.activeReleaseId);
      if (previous.state !== 'active') fail('invalid-state', 'active release transaction is inconsistent');
      previous.state = 'retained';
      delete previous.activationRollbackState;
      previous.updatedAt = this.now();
    }
    transaction.state = 'active';
    transaction.activationRollbackState = 'failed';
    transaction.updatedAt = this.now();
    this.activeReleaseId = releaseId;
    return cloneTransaction(transaction);
  }

  reactivateRetainedRelease(releaseId: string): ReleaseTransactionRecord {
    const transaction = this.requireTransaction(releaseId);
    if (transaction.state !== 'retained') {
      fail('invalid-state', `only a retained release can be reactivated`);
    }
    if (!this.activeReleaseId || this.activeReleaseId === releaseId) {
      fail('invalid-state', `reactivation requires a different active release`);
    }
    const previous = this.requireTransaction(this.activeReleaseId);
    if (previous.state !== 'active') fail('invalid-state', 'active release transaction is inconsistent');
    previous.state = 'retained';
    delete previous.activationRollbackState;
    previous.updatedAt = this.now();
    transaction.previousActiveReleaseId = previous.releaseId;
    transaction.state = 'active';
    transaction.activationRollbackState = 'retained';
    transaction.updatedAt = this.now();
    this.activeReleaseId = releaseId;
    return cloneTransaction(transaction);
  }

  rollbackActivation(releaseId: string, failureCode: string): ReleaseTransactionRecord {
    const transaction = this.requireTransaction(releaseId);
    if (transaction.state !== 'active' || this.activeReleaseId !== releaseId || !failureCode) {
      fail('invalid-state', `release ${releaseId} is not the active rollback target`);
    }
    if (transaction.previousActiveReleaseId) {
      const previous = this.requireTransaction(transaction.previousActiveReleaseId);
      if (previous.state !== 'retained') {
        fail('invalid-state', `previous release for ${releaseId} is not retained`);
      }
      previous.state = 'active';
      previous.updatedAt = this.now();
    }
    const rollbackState = transaction.activationRollbackState || 'failed';
    transaction.state = rollbackState;
    delete transaction.activationRollbackState;
    if (rollbackState === 'failed') transaction.failureCode = failureCode;
    else delete transaction.failureCode;
    transaction.updatedAt = this.now();
    this.activeReleaseId = transaction.previousActiveReleaseId;
    return cloneTransaction(transaction);
  }

  failRelease(releaseId: string, failureCode: string): ReleaseTransactionRecord {
    const transaction = this.requireTransaction(releaseId);
    if (transaction.state !== 'installing' && transaction.state !== 'prepared') {
      fail('invalid-state', `release ${releaseId} cannot fail from state ${transaction.state}`);
    }
    if (!failureCode) fail('invalid-model', 'failure code is required');
    transaction.state = 'failed';
    transaction.failureCode = failureCode;
    transaction.updatedAt = this.now();
    return cloneTransaction(transaction);
  }

  resumeRelease(releaseId: string): ReleaseTransactionRecord {
    const transaction = this.requireTransaction(releaseId);
    if (transaction.state !== 'failed') fail('invalid-state', `release ${releaseId} is not failed`);
    if (this.activeReleaseId !== transaction.previousActiveReleaseId) {
      fail('invalid-state', `the active release changed before ${releaseId} could resume`);
    }
    transaction.state = 'installing';
    delete transaction.failureCode;
    transaction.updatedAt = this.now();
    return cloneTransaction(transaction);
  }

  getResumeState(releaseId: string): ReleaseResumeState {
    const transaction = this.requireTransaction(releaseId);
    const committed = new Set(transaction.committedMappings.map((mapping) => mapping.path));
    return {
      releaseId,
      state: transaction.state,
      missingObjects: transaction.requiredObjects
        .filter((object) => this.objects.get(object.sha256)?.bytes !== object.bytes)
        .map(cloneDescriptor)
        .sort(compareDescriptors),
      pendingAssetPaths: transaction.plannedMappings
        .filter((mapping) => !committed.has(mapping.path))
        .map((mapping) => mapping.path)
        .sort(),
    };
  }

  getActiveRelease(): ReleaseTransactionRecord | null {
    return this.activeReleaseId ? cloneTransaction(this.requireTransaction(this.activeReleaseId)) : null;
  }

  getTransaction(releaseId: string): ReleaseTransactionRecord | null {
    const transaction = this.transactions.get(releaseId);
    return transaction ? cloneTransaction(transaction) : null;
  }

  listObjects(): ContentObjectRecord[] {
    return [...this.objects.values()].map(cloneRecord).sort(compareDescriptors);
  }

  listAssetMappings(releaseId: string): AssetContentMapping[] {
    return this.requireTransaction(releaseId).committedMappings.map(cloneMapping).sort(compareMappings);
  }

  getReleaseReadView(releaseId?: string): ReleaseContentReadView {
    const transaction = releaseId ? this.transactions.get(releaseId) : this.getActiveRelease();
    if (!transaction) return { transaction: null, objects: [] };
    const objects: ContentObjectRecord[] = [];
    for (const descriptor of transaction.requiredObjects) {
      const record = this.objects.get(descriptor.sha256);
      if (record) objects.push(cloneRecord(record));
    }
    objects.sort(compareDescriptors);
    return {
      transaction: cloneTransaction(transaction),
      objects,
    };
  }

  snapshot(): ReleaseContentJournalSnapshot {
    return {
      version: 1,
      activeReleaseId: this.activeReleaseId,
      objects: this.listObjects(),
      transactions: [...this.transactions.values()]
        .map(cloneTransaction)
        .sort((left, right) => left.releaseId.localeCompare(right.releaseId)),
    };
  }

  private requireTransaction(releaseId: string): ReleaseTransactionRecord {
    const transaction = this.transactions.get(releaseId);
    if (!transaction) fail('invalid-state', `release ${releaseId} has no transaction`);
    return transaction;
  }

  private requireMutableTransaction(releaseId: string): ReleaseTransactionRecord {
    const transaction = this.requireTransaction(releaseId);
    if (transaction.state !== 'installing') {
      fail('invalid-state', `release ${releaseId} is not accepting installation writes`);
    }
    return transaction;
  }
}
