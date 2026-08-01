import { IDBDatabase, IDBFactory, IDBIndex, IDBKeyRange as FakeIDBKeyRange, IDBObjectStore } from 'fake-indexeddb';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { IndexedDbCanonicalResourceJournal } from '../../src/lib/canonical-resource-store';
import { planReleaseContent, type ReleaseContentModel } from '../../src/lib/release-content-model';

function largeRelease(objectCount: number): ReleaseContentModel {
  const objects = Array.from({ length: objectCount }, (_, index) => ({
    sha256: index.toString(16).padStart(64, '0'),
    bytes: 1,
  }));
  return {
    releaseId: 'release-ten-thousand',
    manifestSha256: 'a'.repeat(64),
    storageSetSha256: 'b'.repeat(64),
    objects,
    assets: objects.map((object, index) => ({
      path: index === 0 ? 'assets/base-first.bin' : `sdkjs/assets/${index.toString().padStart(5, '0')}.bin`,
      bytes: object.bytes,
      sha256: object.sha256,
      representations: [
        {
          id: 'whole',
          kind: 'whole',
          spans: [
            {
              objectSha256: object.sha256,
              objectOffset: 0,
              assetOffset: 0,
              bytes: object.bytes,
            },
          ],
        },
      ],
    })),
  };
}

function callsForStore(spy: { mock: { instances: unknown[] } }, storeName: string): number {
  return spy.mock.instances.filter((instance) => (instance as { name?: string } | undefined)?.name === storeName)
    .length;
}

afterEach(() => {
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe('IndexedDbCanonicalResourceJournal scale', () => {
  it('checkpoints and commits 10,000 objects without per-object full scans or mapping rewrites', async () => {
    vi.stubGlobal('IDBKeyRange', FakeIDBKeyRange);
    const transactionSpy = vi.spyOn(IDBDatabase.prototype, 'transaction');
    const getSpy = vi.spyOn(IDBObjectStore.prototype, 'get');
    const getAllSpy = vi.spyOn(IDBObjectStore.prototype, 'getAll');
    const putSpy = vi.spyOn(IDBObjectStore.prototype, 'put');
    const deleteSpy = vi.spyOn(IDBObjectStore.prototype, 'delete');
    const openCursorSpy = vi.spyOn(IDBObjectStore.prototype, 'openCursor');
    const openKeyCursorSpy = vi.spyOn(IDBIndex.prototype, 'openKeyCursor');
    const objectCount = 10_000;
    const model = largeRelease(objectCount);
    const plan = planReleaseContent(model, [], []);
    const journal = new IndexedDbCanonicalResourceJournal(new IDBFactory(), {
      databaseName: `journal-scale-${crypto.randomUUID()}`,
      now: () => 1,
    });

    await journal.beginRelease(plan, model);
    for (const spy of [transactionSpy, getSpy, getAllSpy, putSpy, deleteSpy, openCursorSpy, openKeyCursorSpy]) {
      spy.mockClear();
    }

    await journal.recordVerifiedObjects(model.releaseId, plan.requiredObjects, 1);

    const batchCount = Math.ceil(objectCount / 256);
    expect(transactionSpy).toHaveBeenCalledTimes(batchCount);
    expect(getSpy).toHaveBeenCalledTimes(objectCount * 2 + batchCount);
    expect(getAllSpy).not.toHaveBeenCalled();
    expect(openCursorSpy).not.toHaveBeenCalled();
    expect(openKeyCursorSpy).not.toHaveBeenCalled();
    expect(deleteSpy).not.toHaveBeenCalled();
    expect(callsForStore(putSpy, 'objects')).toBe(objectCount);
    expect(callsForStore(putSpy, 'releaseTransactions')).toBe(0);
    for (const spy of [transactionSpy, getSpy, getAllSpy, putSpy, deleteSpy, openCursorSpy, openKeyCursorSpy]) {
      spy.mockClear();
    }

    await journal.commitAssetMappings(
      model.releaseId,
      plan.mappings.map((mapping) => mapping.path),
    );

    expect(transactionSpy).toHaveBeenCalledTimes(batchCount);
    expect(getAllSpy).not.toHaveBeenCalled();
    expect(openCursorSpy).not.toHaveBeenCalled();
    expect(openKeyCursorSpy).not.toHaveBeenCalled();
    expect(deleteSpy).not.toHaveBeenCalled();
    expect(callsForStore(putSpy, 'assetMappings')).toBe(objectCount);
    expect(callsForStore(putSpy, 'releaseTransactions')).toBe(0);

    await journal.prepareRelease(model.releaseId);
    await journal.activateRelease(model.releaseId);
    for (const spy of [transactionSpy, getSpy, getAllSpy, openCursorSpy, openKeyCursorSpy]) spy.mockClear();

    const lastPath = plan.mappings.at(-1)!.path;
    const assetView = await journal.getAssetReadView(model.releaseId, lastPath);
    expect(assetView.mapping?.path).toBe(lastPath);
    expect(assetView.objects).toHaveLength(1);
    expect(transactionSpy).toHaveBeenCalledTimes(1);
    expect(getSpy.mock.calls.length).toBeLessThanOrEqual(4);
    expect(getAllSpy).not.toHaveBeenCalled();
    expect(openCursorSpy).not.toHaveBeenCalled();
    expect(openKeyCursorSpy).not.toHaveBeenCalled();
    for (const spy of [transactionSpy, getSpy, getAllSpy, openCursorSpy, openKeyCursorSpy]) spy.mockClear();

    const probeView = await journal.getReleaseProbeView(model.releaseId);
    expect(probeView.mapping?.path).toBe('sdkjs/assets/00001.bin');
    expect(probeView.object?.sha256).toBe(model.objects[1].sha256);
    expect(transactionSpy).toHaveBeenCalledTimes(1);
    expect(getSpy.mock.calls.length).toBeLessThanOrEqual(4);
    expect(getAllSpy).not.toHaveBeenCalled();
    expect(openCursorSpy).not.toHaveBeenCalled();
    expect(openKeyCursorSpy).not.toHaveBeenCalled();
  }, 120_000);
});
