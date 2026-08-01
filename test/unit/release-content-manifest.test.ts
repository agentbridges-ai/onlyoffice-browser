import { describe, expect, it } from 'vitest';
import { releaseManifestV5ToContentModel } from '../../src/lib/release-content-manifest';
import {
  FASTCDC_V1_POLICY_ID,
  planReleaseContent,
  type ContentObjectRecord,
} from '../../src/lib/release-content-model';
import {
  computeStorageSetSha256,
  parseReleaseManifest,
  type ReleaseAssetV5,
  type ReleaseManifestV5,
} from '../../src/lib/release-resources';

const digest = (character: string) => character.repeat(64);

function manifest(): ReleaseManifestV5 {
  const assets: ReleaseAssetV5[] = [
    {
      path: 'sdkjs/word/word.js',
      bytes: 6,
      mime: 'text/javascript',
      sha256: digest('a'),
      profile: 'word',
      chunk: 'word-001',
      packageOffset: 16,
      representations: {
        whole: { sha256: digest('a'), bytes: 6 },
        fastcdc: {
          algorithm: 'fastcdc-v2020',
          minBytes: 65_536,
          averageBytes: 262_144,
          maxBytes: 1_048_576,
          normalization: 1,
          seed: 0,
          chunks: [
            { offset: 0, bytes: 2, sha256: digest('b') },
            { offset: 2, bytes: 4, sha256: digest('c') },
          ],
        },
      },
    },
  ];
  const packageDescriptor = {
    format: 'onlyoffice-pack-v1' as const,
    path: 'office-resources.oobpack' as const,
    bytes: 22,
    sha256: digest('d'),
    headerBytes: 16,
    segmentBytes: 18,
    segments: [
      { id: digest('e'), offset: 0, bytes: 18, sha256: digest('e') },
      { id: digest('f'), offset: 18, bytes: 4, sha256: digest('f') },
    ],
  };
  return parseReleaseManifest({
    version: 5,
    releaseId: 'v5-content-model',
    packageVersion: '0.6.0',
    hostBuildId: digest('1'),
    shellRevision: digest('2'),
    runtimeManifestSha256: digest('3'),
    fontManifestSha256: digest('4'),
    x2t: { version: '9.3.0+2', commit: 'abc', sha256: digest('5') },
    profiles: {
      base: [],
      word: ['sdkjs/word/word.js'],
      cell: [],
      slide: [],
      'fonts-basic': [],
      'fonts-office-compat': [],
    },
    chunks: [{ id: 'word-001', profile: 'word', bytes: 6, paths: ['sdkjs/word/word.js'] }],
    package: packageDescriptor,
    assets,
    contentProtocol: {
      version: 1,
      digest: 'sha256',
      cacheKeyFormat: 'canonical-sha256-v1',
      storageSetSha256: computeStorageSetSha256(packageDescriptor, assets),
      fastcdcPolicyId: FASTCDC_V1_POLICY_ID,
    },
  }) as ReleaseManifestV5;
}

describe('v5 manifest content model adapter', () => {
  it('maps a package-crossing asset without copying its bytes', () => {
    const model = releaseManifestV5ToContentModel(manifest(), digest('9'));
    expect(model.assets[0].mime).toBe('text/javascript');
    expect(model.objects).toEqual(
      expect.arrayContaining([
        { sha256: digest('a'), bytes: 6 },
        { sha256: digest('b'), bytes: 2 },
        { sha256: digest('c'), bytes: 4 },
        { sha256: digest('e'), bytes: 18 },
        { sha256: digest('f'), bytes: 4 },
      ]),
    );
    expect(model.assets[0].representations).toEqual(
      expect.arrayContaining([
        {
          id: 'package',
          kind: 'package',
          spans: [
            { objectSha256: digest('e'), objectOffset: 16, assetOffset: 0, bytes: 2 },
            { objectSha256: digest('f'), objectOffset: 0, assetOffset: 2, bytes: 4 },
          ],
        },
        {
          id: FASTCDC_V1_POLICY_ID,
          kind: 'fastcdc',
          fastCdcPolicyId: FASTCDC_V1_POLICY_ID,
          spans: [
            { objectSha256: digest('b'), objectOffset: 0, assetOffset: 0, bytes: 2 },
            { objectSha256: digest('c'), objectOffset: 0, assetOffset: 2, bytes: 4 },
          ],
        },
      ]),
    );
  });

  it('chooses the least missing representation deterministically', () => {
    const model = releaseManifestV5ToContentModel(manifest(), digest('9'));
    const available: ContentObjectRecord[] = [
      { sha256: digest('b'), bytes: 2, verifiedAt: 1 },
      { sha256: digest('c'), bytes: 4, verifiedAt: 1 },
    ];
    const plan = planReleaseContent(model, available);
    expect(plan.downloadBytes).toBe(0);
    expect(plan.mappings[0]).toMatchObject({
      mime: 'text/javascript',
      representationKind: 'fastcdc',
      representationId: FASTCDC_V1_POLICY_ID,
    });
  });

  it('rejects a conflicting object size before planning', () => {
    const release = manifest();
    release.assets[0].representations.fastcdc!.chunks[0].sha256 = release.package.segments[0].sha256;
    expect(() => releaseManifestV5ToContentModel(release, digest('9'))).toThrow(/conflicting byte lengths/);
  });
});
