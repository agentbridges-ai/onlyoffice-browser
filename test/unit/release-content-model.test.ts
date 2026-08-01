import { describe, expect, it } from 'vitest';
import {
  FASTCDC_V1_POLICY_ID,
  MemoryReleaseContentJournal,
  ReleaseContentModelError,
  estimateRepresentationCost,
  planReleaseContent,
  validateAssetSpanCoverage,
  validateReleaseContentModel,
  type AssetContentMapping,
  type AssetContentRepresentation,
  type ContentObjectDescriptor,
  type ContentObjectRecord,
  type ReleaseContentAsset,
  type ReleaseContentModel,
} from '../../src/lib/release-content-model';

const digest = (character: string): string => character.repeat(64);
const manifestA = digest('a');
const manifestB = digest('b');
const storageA = digest('c');
const storageB = digest('d');

function wholeRepresentation(assetSha256: string, bytes: number): AssetContentRepresentation {
  return {
    id: 'whole',
    kind: 'whole',
    spans: [{ objectSha256: assetSha256, objectOffset: 0, assetOffset: 0, bytes }],
  };
}

function model(
  releaseId: string,
  assets: ReleaseContentAsset[],
  objects: ContentObjectDescriptor[],
  manifestSha256 = manifestA,
  storageSetSha256 = storageA,
): ReleaseContentModel {
  return {
    releaseId,
    manifestSha256,
    storageSetSha256,
    assets,
    objects,
  };
}

function installAndActivate(
  journal: MemoryReleaseContentJournal,
  release: ReleaseContentModel,
  inventory: ContentObjectRecord[] = journal.listObjects(),
  previousMappings: AssetContentMapping[] = [],
): void {
  const plan = planReleaseContent(release, inventory, previousMappings);
  journal.beginRelease(plan, release);
  for (const object of plan.missingObjects) journal.recordVerifiedObject(release.releaseId, object);
  for (const mapping of plan.mappings) journal.commitAssetMapping(release.releaseId, mapping.path);
  journal.prepareRelease(release.releaseId);
  journal.activateRelease(release.releaseId);
}

describe('release content span model', () => {
  it('requires exact contiguous coverage and physical object bounds', () => {
    const objects = new Map<string, ContentObjectDescriptor>([
      [digest('1'), { sha256: digest('1'), bytes: 4 }],
      [digest('2'), { sha256: digest('2'), bytes: 4 }],
    ]);
    const asset = { path: 'sdkjs/word/word.js', bytes: 8, sha256: digest('3') };
    const valid: AssetContentRepresentation = {
      id: 'package',
      kind: 'package',
      spans: [
        { objectSha256: digest('1'), objectOffset: 0, assetOffset: 0, bytes: 4 },
        { objectSha256: digest('2'), objectOffset: 0, assetOffset: 4, bytes: 4 },
      ],
    };
    expect(() => validateAssetSpanCoverage(asset, valid, objects)).not.toThrow();
    expect(() =>
      validateAssetSpanCoverage(
        asset,
        {
          ...valid,
          spans: [valid.spans[0], { ...valid.spans[1], assetOffset: 5 }],
        },
        objects,
      ),
    ).toThrowError(/gap or overlap/);
    expect(() =>
      validateAssetSpanCoverage(
        asset,
        {
          ...valid,
          spans: [valid.spans[0], { ...valid.spans[1], objectOffset: 1 }],
        },
        objects,
      ),
    ).toThrowError(/exceeds/);
  });

  it('enforces whole-file identity and complete FastCDC chunk objects', () => {
    const wholeSha = digest('4');
    const chunkOne = digest('5');
    const chunkTwo = digest('6');
    const objects = new Map<string, ContentObjectDescriptor>([
      [wholeSha, { sha256: wholeSha, bytes: 8 }],
      [chunkOne, { sha256: chunkOne, bytes: 4 }],
      [chunkTwo, { sha256: chunkTwo, bytes: 4 }],
    ]);
    const asset = { path: 'sdkjs/common/data.bin', bytes: 8, sha256: wholeSha };
    expect(() => validateAssetSpanCoverage(asset, wholeRepresentation(wholeSha, 8), objects)).not.toThrow();
    expect(() =>
      validateAssetSpanCoverage(
        asset,
        {
          id: 'fastcdc',
          kind: 'fastcdc',
          fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
          spans: [
            { objectSha256: chunkOne, objectOffset: 0, assetOffset: 0, bytes: 4 },
            { objectSha256: chunkTwo, objectOffset: 0, assetOffset: 4, bytes: 4 },
          ],
        },
        objects,
      ),
    ).not.toThrow();
    expect(() =>
      validateAssetSpanCoverage(
        asset,
        {
          ...wholeRepresentation(wholeSha, 8),
          spans: [{ objectSha256: chunkOne, objectOffset: 0, assetOffset: 0, bytes: 8 }],
        },
        objects,
      ),
    ).toThrowError();
    expect(() =>
      validateAssetSpanCoverage(
        asset,
        {
          id: 'fastcdc',
          kind: 'fastcdc',
          fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
          spans: [
            { objectSha256: chunkOne, objectOffset: 1, assetOffset: 0, bytes: 4 },
            { objectSha256: chunkTwo, objectOffset: 0, assetOffset: 4, bytes: 4 },
          ],
        },
        objects,
      ),
    ).toThrowError();
  });

  it('rejects digest-size conflicts before planning', () => {
    const sha = digest('7');
    const release = model(
      'v1',
      [{ path: 'a.bin', bytes: 4, sha256: sha, representations: [wholeRepresentation(sha, 4)] }],
      [
        { sha256: sha, bytes: 4 },
        { sha256: sha, bytes: 5 },
      ],
    );
    expect(() => validateReleaseContentModel(release)).toThrowError(
      expect.objectContaining<Partial<ReleaseContentModelError>>({ code: 'integrity-conflict' }),
    );
  });

  it('uses stable package segments for a cold install even when whole blobs are marginally smaller', () => {
    const wholeOne = digest('4');
    const wholeTwo = digest('5');
    const packageSha = digest('6');
    const objects = [
      { sha256: wholeOne, bytes: 7 },
      { sha256: wholeTwo, bytes: 7 },
      { sha256: packageSha, bytes: 16 },
    ];
    const assets: ReleaseContentAsset[] = [
      {
        path: 'a.bin',
        bytes: 7,
        sha256: wholeOne,
        representations: [
          wholeRepresentation(wholeOne, 7),
          {
            id: 'package',
            kind: 'package',
            spans: [{ objectSha256: packageSha, objectOffset: 1, assetOffset: 0, bytes: 7 }],
          },
        ],
      },
      {
        path: 'b.bin',
        bytes: 7,
        sha256: wholeTwo,
        representations: [
          wholeRepresentation(wholeTwo, 7),
          {
            id: 'package',
            kind: 'package',
            spans: [{ objectSha256: packageSha, objectOffset: 8, assetOffset: 0, bytes: 7 }],
          },
        ],
      },
    ];

    const normal = planReleaseContent(model('v-cold-normal', assets, objects), []);
    const cold = planReleaseContent(model('v-cold-package', assets, objects), [], [], {
      preferPackageForColdInstall: true,
    });
    const coldWhole = planReleaseContent(model('v-cold-whole', assets, objects), [], [], {
      preferWholeForColdInstall: true,
    });

    expect(normal.requiredObjects).toEqual([
      { sha256: wholeOne, bytes: 7 },
      { sha256: wholeTwo, bytes: 7 },
    ]);
    expect(cold.requiredObjects).toEqual([{ sha256: packageSha, bytes: 16 }]);
    expect(cold.mappings.every((mapping) => mapping.representationKind === 'package')).toBe(true);
    expect(coldWhole.requiredObjects).toEqual([
      { sha256: wholeOne, bytes: 7 },
      { sha256: wholeTwo, bytes: 7 },
    ]);
    expect(coldWhole.mappings.every((mapping) => mapping.representationKind === 'whole')).toBe(true);
    expect(() =>
      planReleaseContent(model('v-cold-conflict', assets, objects), [], [], {
        preferPackageForColdInstall: true,
        preferWholeForColdInstall: true,
      }),
    ).toThrow('cannot prefer multiple');
  });

  it('keeps package transport out of the cold active mapping and seeds FastCDC inventory when available', () => {
    const whole = digest('1');
    const chunkOne = digest('2');
    const chunkTwo = digest('3');
    const packaged = digest('4');
    const release = model(
      'v-cold-canonical-fastcdc',
      [
        {
          path: 'runtime.bin',
          bytes: 8,
          sha256: whole,
          representations: [
            wholeRepresentation(whole, 8),
            {
              id: FASTCDC_V1_POLICY_ID,
              kind: 'fastcdc',
              fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
              spans: [
                { objectSha256: chunkOne, objectOffset: 0, assetOffset: 0, bytes: 4 },
                { objectSha256: chunkTwo, objectOffset: 0, assetOffset: 4, bytes: 4 },
              ],
            },
            {
              id: 'package',
              kind: 'package',
              spans: [{ objectSha256: packaged, objectOffset: 1, assetOffset: 0, bytes: 8 }],
            },
          ],
        },
      ],
      [
        { sha256: whole, bytes: 8 },
        { sha256: chunkOne, bytes: 4 },
        { sha256: chunkTwo, bytes: 4 },
        { sha256: packaged, bytes: 10 },
      ],
    );

    const plan = planReleaseContent(release, [], [], { preferCanonicalForColdInstall: true });

    expect(plan.requiredObjects).toEqual([
      { sha256: chunkOne, bytes: 4 },
      { sha256: chunkTwo, bytes: 4 },
    ]);
    expect(plan.mappings).toMatchObject([
      {
        representationKind: 'fastcdc',
        representationId: FASTCDC_V1_POLICY_ID,
      },
    ]);
  });
});

describe('release representation planner', () => {
  it('reports whole, FastCDC, and package missing-byte cost and selects the local change', () => {
    const wholeSha = digest('8');
    const oldOne = digest('9');
    const changed = digest('a');
    const oldThree = digest('b');
    const packageSha = digest('c');
    const objects: ContentObjectDescriptor[] = [
      { sha256: wholeSha, bytes: 9 },
      { sha256: oldOne, bytes: 3 },
      { sha256: changed, bytes: 3 },
      { sha256: oldThree, bytes: 3 },
      { sha256: packageSha, bytes: 20 },
    ];
    const asset: ReleaseContentAsset = {
      path: 'sdkjs/word/large.bin',
      bytes: 9,
      sha256: wholeSha,
      representations: [
        wholeRepresentation(wholeSha, 9),
        {
          id: 'fastcdc',
          kind: 'fastcdc',
          fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
          spans: [
            { objectSha256: oldOne, objectOffset: 0, assetOffset: 0, bytes: 3 },
            { objectSha256: changed, objectOffset: 0, assetOffset: 3, bytes: 3 },
            { objectSha256: oldThree, objectOffset: 0, assetOffset: 6, bytes: 3 },
          ],
        },
        {
          id: 'package',
          kind: 'package',
          spans: [{ objectSha256: packageSha, objectOffset: 5, assetOffset: 0, bytes: 9 }],
        },
      ],
    };
    const inventory: ContentObjectRecord[] = [
      { sha256: oldOne, bytes: 3, verifiedAt: 1 },
      { sha256: oldThree, bytes: 3, verifiedAt: 1 },
    ];
    const objectMap = new Map(objects.map((object) => [object.sha256, object]));
    const inventoryMap = new Map(inventory.map((object) => [object.sha256, object]));

    expect(
      asset.representations.map((representation) => ({
        kind: representation.kind,
        bytes: estimateRepresentationCost(representation, objectMap, inventoryMap).downloadBytes,
      })),
    ).toEqual([
      { kind: 'whole', bytes: 9 },
      { kind: 'fastcdc', bytes: 3 },
      { kind: 'package', bytes: 20 },
    ]);

    const plan = planReleaseContent(model('v2', [asset], objects, manifestB, storageB), inventory);
    expect(plan).toMatchObject({
      downloadBytes: 3,
      reusedObjectBytes: 6,
      requiredObjectBytes: 9,
      missingObjects: [{ sha256: changed, bytes: 3 }],
    });
    expect(plan.mappings[0]).toMatchObject({ representationKind: 'fastcdc', representationId: 'fastcdc' });
  });

  it('aliases an unchanged A asset into B with zero network and zero byte copy', () => {
    const sha = digest('d');
    const assetA: ReleaseContentAsset = {
      path: 'sdkjs/word/word.js',
      bytes: 12,
      sha256: sha,
      mime: 'text/javascript',
      representations: [wholeRepresentation(sha, 12)],
    };
    const releaseA = model('release-a', [assetA], [{ sha256: sha, bytes: 12 }]);
    const journal = new MemoryReleaseContentJournal({ now: () => 1 });
    installAndActivate(journal, releaseA);

    const assetB = {
      ...assetA,
      path: 'sdkjs/word/renamed-word.js',
      mime: 'application/javascript',
    };
    const releaseB = model('release-b', [assetB], [{ sha256: sha, bytes: 12 }], manifestB, storageB);
    const planB = planReleaseContent(releaseB, journal.listObjects(), journal.listAssetMappings('release-a'));

    expect(planB.downloadBytes).toBe(0);
    expect(planB.requiredObjectBytes).toBe(12);
    expect(planB.reusedObjectBytes).toBe(12);
    expect(planB.mappings).toEqual([
      expect.objectContaining({
        releaseId: 'release-b',
        path: assetB.path,
        mime: 'application/javascript',
        sourceReleaseId: 'release-a',
        representationKind: 'whole',
      }),
    ]);
  });

  it('aliases A package spans after a B prefix insertion changes every package offset', () => {
    const assetSha = digest('e');
    const oldPackageSha = digest('f');
    const newPackageSha = digest('0');
    const oldAsset: ReleaseContentAsset = {
      path: 'sdkjs/word/word.js',
      bytes: 9,
      sha256: assetSha,
      representations: [
        {
          id: 'package-a',
          kind: 'package',
          spans: [{ objectSha256: oldPackageSha, objectOffset: 5, assetOffset: 0, bytes: 9 }],
        },
      ],
    };
    const releaseA = model('release-a', [oldAsset], [{ sha256: oldPackageSha, bytes: 20 }]);
    const journal = new MemoryReleaseContentJournal({ now: () => 1 });
    installAndActivate(journal, releaseA);

    const releaseB = model(
      'release-b',
      [
        {
          ...oldAsset,
          representations: [
            wholeRepresentation(assetSha, 9),
            {
              id: 'package-b-after-prefix-insert',
              kind: 'package',
              spans: [{ objectSha256: newPackageSha, objectOffset: 6, assetOffset: 0, bytes: 9 }],
            },
          ],
        },
      ],
      [
        { sha256: assetSha, bytes: 9 },
        { sha256: newPackageSha, bytes: 21 },
      ],
      manifestB,
      storageB,
    );
    const planB = planReleaseContent(releaseB, journal.listObjects(), journal.listAssetMappings('release-a'));

    expect(planB.downloadBytes).toBe(0);
    expect(planB.requiredObjects).toEqual([{ sha256: oldPackageSha, bytes: 20 }]);
    expect(planB.mappings).toEqual([
      expect.objectContaining({
        releaseId: 'release-b',
        representationId: 'package-a',
        representationKind: 'package',
        sourceReleaseId: 'release-a',
        spans: [{ objectSha256: oldPackageSha, objectOffset: 5, assetOffset: 0, bytes: 9 }],
      }),
    ]);
  });

  it('deduplicates a shared package object across assets before comparing plans', () => {
    const wholeOne = digest('1');
    const wholeTwo = digest('2');
    const packageSha = digest('3');
    const objects = [
      { sha256: wholeOne, bytes: 7 },
      { sha256: wholeTwo, bytes: 7 },
      { sha256: packageSha, bytes: 10 },
    ];
    const assets: ReleaseContentAsset[] = [
      {
        path: 'a.bin',
        bytes: 7,
        sha256: wholeOne,
        representations: [
          wholeRepresentation(wholeOne, 7),
          {
            id: 'package',
            kind: 'package',
            spans: [{ objectSha256: packageSha, objectOffset: 0, assetOffset: 0, bytes: 7 }],
          },
        ],
      },
      {
        path: 'b.bin',
        bytes: 7,
        sha256: wholeTwo,
        representations: [
          wholeRepresentation(wholeTwo, 7),
          {
            id: 'package',
            kind: 'package',
            spans: [{ objectSha256: packageSha, objectOffset: 3, assetOffset: 0, bytes: 7 }],
          },
        ],
      },
    ];

    const plan = planReleaseContent(model('v-package', assets, objects), []);
    expect(plan.downloadBytes).toBe(10);
    expect(plan.requiredObjects).toEqual([{ sha256: packageSha, bytes: 10 }]);
    expect(plan.mappings.every((mapping) => mapping.representationKind === 'package')).toBe(true);
  });
});

describe('memory release content journal', () => {
  it('includes MIME in the immutable release plan fingerprint', () => {
    const sha = digest('0');
    const release = model(
      'release-mime-fingerprint',
      [
        {
          path: 'sdkjs/word/word.js',
          bytes: 4,
          sha256: sha,
          mime: 'text/javascript',
          representations: [wholeRepresentation(sha, 4)],
        },
      ],
      [{ sha256: sha, bytes: 4 }],
    );
    const journal = new MemoryReleaseContentJournal();
    journal.beginRelease(planReleaseContent(release, []), release);

    const changedMimeRelease: ReleaseContentModel = {
      ...release,
      assets: release.assets.map((asset) => ({ ...asset, mime: 'application/javascript' })),
    };
    expect(() => journal.beginRelease(planReleaseContent(changedMimeRelease, []), changedMimeRelease)).toThrowError(
      expect.objectContaining<Partial<ReleaseContentModelError>>({ code: 'integrity-conflict' }),
    );
  });

  it('persists an interrupted transaction and resumes only its missing objects and mappings', () => {
    const first = digest('1');
    const second = digest('2');
    const asset: ReleaseContentAsset = {
      path: 'large.bin',
      bytes: 8,
      sha256: digest('3'),
      representations: [
        {
          id: 'fastcdc',
          kind: 'fastcdc',
          fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
          spans: [
            { objectSha256: first, objectOffset: 0, assetOffset: 0, bytes: 4 },
            { objectSha256: second, objectOffset: 0, assetOffset: 4, bytes: 4 },
          ],
        },
      ],
    };
    const release = model(
      'release-interrupted',
      [asset],
      [
        { sha256: first, bytes: 4 },
        { sha256: second, bytes: 4 },
      ],
    );
    const plan = planReleaseContent(release, []);
    const journal = new MemoryReleaseContentJournal({ now: () => 10 });
    journal.beginRelease(plan, release);
    journal.recordVerifiedObject(release.releaseId, { sha256: first, bytes: 4 });
    expect(() => journal.commitAssetMapping(release.releaseId, asset.path)).toThrowError(/not verified/);
    expect(() => journal.prepareRelease(release.releaseId)).toThrowError(
      expect.objectContaining<Partial<ReleaseContentModelError>>({ code: 'incomplete' }),
    );

    const restarted = MemoryReleaseContentJournal.fromSnapshot(journal.snapshot(), () => 20);
    expect(restarted.getResumeState(release.releaseId)).toEqual({
      releaseId: release.releaseId,
      state: 'installing',
      missingObjects: [{ sha256: second, bytes: 4 }],
      pendingAssetPaths: [asset.path],
    });
    restarted.recordVerifiedObject(release.releaseId, { sha256: second, bytes: 4 });
    restarted.commitAssetMapping(release.releaseId, asset.path);
    restarted.prepareRelease(release.releaseId);
    restarted.activateRelease(release.releaseId);
    expect(restarted.getActiveRelease()).toMatchObject({ releaseId: release.releaseId, state: 'active' });
  });

  it('keeps A active when B fails and never activates the failed transaction', () => {
    const aSha = digest('4');
    const bSha = digest('5');
    const releaseA = model(
      'release-a',
      [
        {
          path: 'a.bin',
          bytes: 4,
          sha256: aSha,
          representations: [wholeRepresentation(aSha, 4)],
        },
      ],
      [{ sha256: aSha, bytes: 4 }],
    );
    const releaseB = model(
      'release-b',
      [
        {
          path: 'a.bin',
          bytes: 4,
          sha256: bSha,
          representations: [wholeRepresentation(bSha, 4)],
        },
      ],
      [{ sha256: bSha, bytes: 4 }],
      manifestB,
      storageB,
    );
    const journal = new MemoryReleaseContentJournal({ now: () => 1 });
    installAndActivate(journal, releaseA);
    const planB = planReleaseContent(releaseB, journal.listObjects(), journal.listAssetMappings('release-a'));
    journal.beginRelease(planB, releaseB);
    journal.recordVerifiedObject('release-b', { sha256: bSha, bytes: 4 });
    journal.commitAssetMapping('release-b', 'a.bin');
    journal.prepareRelease('release-b');
    journal.failRelease('release-b', 'integrity');

    expect(journal.getActiveRelease()).toMatchObject({ releaseId: 'release-a', state: 'active' });
    expect(journal.getTransaction('release-b')).toMatchObject({ state: 'failed', failureCode: 'integrity' });
    expect(() => journal.activateRelease('release-b')).toThrowError(
      expect.objectContaining<Partial<ReleaseContentModelError>>({ code: 'invalid-state' }),
    );

    const restarted = MemoryReleaseContentJournal.fromSnapshot(journal.snapshot());
    expect(restarted.getActiveRelease()).toMatchObject({ releaseId: 'release-a', state: 'active' });
    restarted.resumeRelease('release-b');
    expect(restarted.getResumeState('release-b')).toMatchObject({
      missingObjects: [],
      pendingAssetPaths: [],
    });
    restarted.prepareRelease('release-b');
    restarted.activateRelease('release-b');
    expect(restarted.getActiveRelease()).toMatchObject({ releaseId: 'release-b', state: 'active' });
    expect(restarted.getTransaction('release-a')).toMatchObject({ state: 'retained' });
  });

  it('rolls an activated B readiness failure back to A and resumes B without re-downloading objects', () => {
    const aSha = digest('9');
    const bSha = digest('a');
    const createRelease = (releaseId: string, sha256: string, manifestSha256: string, storageSetSha256: string) =>
      model(
        releaseId,
        [
          {
            path: 'asset.bin',
            bytes: 4,
            sha256,
            representations: [wholeRepresentation(sha256, 4)],
          },
        ],
        [{ sha256, bytes: 4 }],
        manifestSha256,
        storageSetSha256,
      );
    const releaseA = createRelease('release-a', aSha, manifestA, storageA);
    const releaseB = createRelease('release-b', bSha, manifestB, storageB);
    const journal = new MemoryReleaseContentJournal();
    installAndActivate(journal, releaseA);
    installAndActivate(journal, releaseB, journal.listObjects(), journal.listAssetMappings(releaseA.releaseId));

    journal.rollbackActivation(releaseB.releaseId, 'storage');
    expect(journal.getActiveRelease()).toMatchObject({
      releaseId: releaseA.releaseId,
      state: 'active',
    });
    expect(journal.getTransaction(releaseB.releaseId)).toMatchObject({
      state: 'failed',
      failureCode: 'storage',
    });
    expect(journal.getResumeState(releaseB.releaseId)).toMatchObject({
      missingObjects: [],
      pendingAssetPaths: [],
    });

    journal.resumeRelease(releaseB.releaseId);
    journal.prepareRelease(releaseB.releaseId);
    journal.activateRelease(releaseB.releaseId);
    expect(journal.getActiveRelease()).toMatchObject({
      releaseId: releaseB.releaseId,
      state: 'active',
    });
    expect(journal.getTransaction(releaseA.releaseId)).toMatchObject({ state: 'retained' });
  });

  it('reactivates retained A, rolls it back to retained, and can retry the reactivation', () => {
    const createRelease = (releaseId: string, sha256: string, manifestSha256: string, storageSetSha256: string) =>
      model(
        releaseId,
        [
          {
            path: 'asset.bin',
            bytes: 4,
            sha256,
            representations: [wholeRepresentation(sha256, 4)],
          },
        ],
        [{ sha256, bytes: 4 }],
        manifestSha256,
        storageSetSha256,
      );
    const releaseA = createRelease('release-a', digest('c'), manifestA, storageA);
    const releaseB = createRelease('release-b', digest('d'), manifestB, storageB);
    const journal = new MemoryReleaseContentJournal();
    installAndActivate(journal, releaseA);
    installAndActivate(journal, releaseB, journal.listObjects(), journal.listAssetMappings(releaseA.releaseId));

    journal.reactivateRetainedRelease(releaseA.releaseId);
    expect(journal.getActiveRelease()).toMatchObject({ releaseId: releaseA.releaseId, state: 'active' });
    expect(journal.getTransaction(releaseB.releaseId)).toMatchObject({ state: 'retained' });

    journal.rollbackActivation(releaseA.releaseId, 'probe-failed');
    expect(journal.getActiveRelease()).toMatchObject({ releaseId: releaseB.releaseId, state: 'active' });
    expect(journal.getTransaction(releaseA.releaseId)).toMatchObject({ state: 'retained' });
    expect(journal.getTransaction(releaseA.releaseId)).not.toHaveProperty('failureCode');

    journal.reactivateRetainedRelease(releaseA.releaseId);
    expect(journal.getActiveRelease()).toMatchObject({ releaseId: releaseA.releaseId, state: 'active' });
  });

  it('rolls a failed first activation back to no active release', () => {
    const sha = digest('b');
    const release = model(
      'release-first',
      [
        {
          path: 'asset.bin',
          bytes: 4,
          sha256: sha,
          representations: [wholeRepresentation(sha, 4)],
        },
      ],
      [{ sha256: sha, bytes: 4 }],
    );
    const journal = new MemoryReleaseContentJournal();
    installAndActivate(journal, release);

    journal.rollbackActivation(release.releaseId, 'storage');
    expect(journal.getActiveRelease()).toBeNull();
    expect(journal.getTransaction(release.releaseId)).toMatchObject({
      state: 'failed',
      failureCode: 'storage',
    });
    expect(() => journal.rollbackActivation(release.releaseId, 'storage')).toThrowError(
      expect.objectContaining<Partial<ReleaseContentModelError>>({ code: 'invalid-state' }),
    );
  });

  it('rejects a resumed release if another transaction changed the active pointer', () => {
    const objectA = digest('6');
    const objectB = digest('7');
    const objectC = digest('8');
    const createRelease = (releaseId: string, sha256: string, manifestSha256: string, storageSetSha256: string) =>
      model(
        releaseId,
        [
          {
            path: 'asset.bin',
            bytes: 1,
            sha256,
            representations: [wholeRepresentation(sha256, 1)],
          },
        ],
        [{ sha256, bytes: 1 }],
        manifestSha256,
        storageSetSha256,
      );
    const releaseA = createRelease('release-a', objectA, manifestA, storageA);
    const releaseB = createRelease('release-b', objectB, manifestB, storageB);
    const releaseC = createRelease('release-c', objectC, digest('e'), digest('f'));
    const journal = new MemoryReleaseContentJournal();
    installAndActivate(journal, releaseA);
    const planB = planReleaseContent(releaseB, journal.listObjects());
    journal.beginRelease(planB, releaseB);
    journal.failRelease('release-b', 'network');
    installAndActivate(journal, releaseC);

    expect(() => journal.resumeRelease('release-b')).toThrowError(/active release changed/);
    expect(journal.getActiveRelease()).toMatchObject({ releaseId: 'release-c' });
  });
});
