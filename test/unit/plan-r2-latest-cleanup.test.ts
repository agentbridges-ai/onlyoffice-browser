import crypto from 'node:crypto';
import { describe, expect, it } from 'vitest';

// @ts-expect-error Executable Node release scripts are intentionally shipped as plain Node ESM.
import { buildLatestOnlyCleanupPlan } from '../../scripts/plan-r2-latest-cleanup.mjs';

const digest = (value: string) => crypto.createHash('sha256').update(value).digest('hex');

function fixture() {
  const releaseId = 'v0.5.12-test';
  const manifest = {
    version: 5,
    releaseId,
    package: {
      sha256: digest('office-pack'),
      segments: [{ sha256: digest('office-segment') }],
    },
    assets: [
      { path: 'sdk.js', sha256: digest('sdk'), representations: {} },
      {
        path: 'x2t.wasm',
        sha256: digest('wasm'),
        representations: { fastcdc: { chunks: [{ sha256: digest('wasm-chunk') }] } },
      },
    ],
  };
  const manifestBytes = Buffer.from(`${JSON.stringify(manifest)}\n`);
  const pointer = {
    version: 1,
    releaseId,
    manifestUrl: `/releases/${releaseId}/manifest.json`,
    manifestSha256: digest(manifestBytes.toString()),
  };
  const legacyPointer = {
    version: 1,
    releaseId,
    manifestUrl: `/releases/${releaseId}/manifest-v4.json`,
    manifestSha256: digest('legacy-manifest'),
  };
  const inventory = [
    { Path: 'channels/stable.json', Size: 100 },
    { Path: 'channels/stable-v5.json', Size: 100 },
    { Path: `releases/${releaseId}/manifest.json`, Size: manifestBytes.length },
    { Path: `releases/${releaseId}/manifest-v4.json`, Size: 100 },
    { Path: `releases/${releaseId}/source-identity.json`, Size: 200 },
    { Path: `promotions/${releaseId}/receipt.json`, Size: 300 },
    { Path: `blobs/sha256/${digest('sdk')}`, Size: 3 },
    { Path: `blobs/sha256/${digest('wasm')}`, Size: 4 },
    { Path: `blobs/sha256/${digest('wasm-chunk')}`, Size: 5 },
    { Path: `packages/sha256/${digest('office-pack')}.oobpack`, Size: 11 },
    { Path: `segments/sha256/${digest('office-segment')}`, Size: 12 },
    { Path: 'releases/v0.5.11/manifest.json', Size: 99 },
    { Path: `blobs/sha256/${digest('old')}`, Size: 77 },
  ];
  return { releaseId, manifest, manifestBytes, pointer, legacyPointer, inventory };
}

describe('latest-only R2 cleanup planning', () => {
  it('retains the active release and every manifest-referenced object', () => {
    const input = fixture();
    const plan = buildLatestOnlyCleanupPlan(input);

    expect(plan.mode).toBe('latest-only');
    expect(plan.releaseId).toBe(input.releaseId);
    expect(plan.retainedObjects).toBe(11);
    expect(plan.deleteObjects).toBe(2);
    expect(plan.deleteBytes).toBe(176);
  });

  it('rejects a pointer that is not bound to the supplied manifest', () => {
    const input = fixture();
    expect(() =>
      buildLatestOnlyCleanupPlan({
        ...input,
        pointer: { ...input.pointer, releaseId: 'v0.5.11', manifestUrl: '/releases/v0.5.11/manifest.json' },
      }),
    ).toThrow('legacy stable pointer is not bound to the latest v5 release');
  });

  it('rejects missing referenced objects instead of generating a destructive list', () => {
    const input = fixture();
    expect(() =>
      buildLatestOnlyCleanupPlan({
        ...input,
        inventory: input.inventory.filter((entry) => entry.Path !== `blobs/sha256/${digest('sdk')}`),
      }),
    ).toThrow('latest release references missing R2 objects');
  });

  it('rejects unsafe inventory paths', () => {
    const input = fixture();
    expect(() =>
      buildLatestOnlyCleanupPlan({
        ...input,
        inventory: [...input.inventory, { Path: '../outside', Size: 1 }],
      }),
    ).toThrow('not a safe R2 key');
  });

  it('ignores an S3 root marker and normalizes a harmless leading ./ prefix', () => {
    const input = fixture();
    const plan = buildLatestOnlyCleanupPlan({
      ...input,
      inventory: [{ Path: '', Size: 0 }, { Path: './channels/stable.json', Size: 100 }, ...input.inventory.slice(1)],
    });

    expect(plan.inventoryObjects).toBe(input.inventory.length);
    expect(plan.deleteObjects).toBe(2);
  });
});
