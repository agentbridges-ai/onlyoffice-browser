import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';

// @ts-expect-error Executable Node release scripts are intentionally shipped as plain Node ESM.
import { buildLatestOnlyCleanupPlan } from '../../scripts/plan-r2-latest-cleanup.mjs';

const digest = (value: string) => crypto.createHash('sha256').update(value).digest('hex');
const repositoryRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const cleanupPlanner = path.join(repositoryRoot, 'scripts/plan-r2-latest-cleanup.mjs');

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

  it('normalizes the raw inventory once when invoked through the CLI', () => {
    const input = fixture();
    const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'onlyoffice-r2-cleanup-'));
    try {
      const pointerPath = path.join(temporaryDirectory, 'pointer.json');
      const legacyPointerPath = path.join(temporaryDirectory, 'legacy-pointer.json');
      const manifestPath = path.join(temporaryDirectory, 'manifest.json');
      const inventoryPath = path.join(temporaryDirectory, 'inventory.json');
      const outputPath = path.join(temporaryDirectory, 'plan.json');
      const deleteListPath = path.join(temporaryDirectory, 'delete-list.txt');
      fs.writeFileSync(pointerPath, `${JSON.stringify(input.pointer)}\n`);
      fs.writeFileSync(legacyPointerPath, `${JSON.stringify(input.legacyPointer)}\n`);
      fs.writeFileSync(manifestPath, input.manifestBytes);
      fs.writeFileSync(inventoryPath, `${JSON.stringify(input.inventory)}\n`);

      const result = spawnSync(
        process.execPath,
        [
          cleanupPlanner,
          '--pointer',
          pointerPath,
          '--legacy-pointer',
          legacyPointerPath,
          '--manifest',
          manifestPath,
          '--inventory',
          inventoryPath,
          '--output',
          outputPath,
          '--delete-list',
          deleteListPath,
        ],
        { cwd: repositoryRoot, encoding: 'utf8' },
      );

      expect(result.status, result.stderr || result.stdout).toBe(0);
      expect(result.stderr).toBe('');
      expect(JSON.parse(fs.readFileSync(outputPath, 'utf8'))).toMatchObject({
        mode: 'latest-only',
        releaseId: input.releaseId,
        deleteObjects: 2,
        deleteBytes: 176,
      });
    } finally {
      fs.rmSync(temporaryDirectory, { recursive: true, force: true });
    }
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

  it('rejects dot paths instead of treating them as root markers', () => {
    const input = fixture();
    expect(() =>
      buildLatestOnlyCleanupPlan({
        ...input,
        inventory: [...input.inventory, { Path: '.', Size: 0 }],
      }),
    ).toThrow('not a safe R2 key');
    expect(() =>
      buildLatestOnlyCleanupPlan({
        ...input,
        inventory: [...input.inventory, { Path: './', Size: 0 }],
      }),
    ).toThrow('not a safe R2 key');
  });

  it('accepts an S3 slash root marker only without object metadata', () => {
    const input = fixture();
    const plan = buildLatestOnlyCleanupPlan({
      ...input,
      inventory: [{ Path: '/', Size: undefined }, ...input.inventory],
    });
    expect(plan.inventoryObjects).toBe(input.inventory.length);
    expect(() =>
      buildLatestOnlyCleanupPlan({
        ...input,
        inventory: [{ Path: '/', Size: 1 }, ...input.inventory],
      }),
    ).toThrow('not a safe R2 key');
  });

  it('ignores an S3 root marker and normalizes a harmless leading ./ prefix', () => {
    const input = fixture();
    const plan = buildLatestOnlyCleanupPlan({
      ...input,
      inventory: [
        { Path: '', Size: undefined },
        { Path: './channels/stable.json', Size: 100 },
        ...input.inventory.slice(1),
      ],
    });

    expect(plan.inventoryObjects).toBe(input.inventory.length);
    expect(plan.deleteObjects).toBe(2);
  });
});
