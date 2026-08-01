import { sha256 } from '@noble/hashes/sha2.js';
import { describe, expect, it, vi } from 'vitest';
import {
  CANONICAL_CONTENT_BYTES_HEADER,
  CANONICAL_CONTENT_CACHE_NAME,
  CANONICAL_CONTENT_DIGEST_HEADER,
  CanonicalResourceStore,
  MemoryCanonicalResourceJournal,
  canonicalContentCacheKey,
  type CanonicalResourceJournal,
} from '../../src/lib/canonical-resource-store';
import type {
  AssetContentMapping,
  ContentObjectDescriptor,
  ContentObjectRecord,
  ReleaseContentAsset,
  ReleaseContentModel,
  ReleaseTransactionRecord,
} from '../../src/lib/release-content-model';

function digestBytes(bytes: Uint8Array): string {
  return Array.from(sha256(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function fixedDigest(character: string): string {
  return character.repeat(64);
}

function body(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

function wholeAsset(
  path: string,
  bytes: Uint8Array,
): {
  asset: ReleaseContentAsset;
  object: ContentObjectDescriptor;
} {
  const object = { sha256: digestBytes(bytes), bytes: bytes.byteLength };
  return {
    object,
    asset: {
      path,
      bytes: bytes.byteLength,
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
    },
  };
}

function release(
  releaseId: string,
  content: Array<{ path: string; bytes: Uint8Array }>,
  manifestSha256 = fixedDigest('a'),
  storageSetSha256 = fixedDigest('b'),
): ReleaseContentModel {
  const entries = content.map(({ path, bytes }) => wholeAsset(path, bytes));
  return {
    releaseId,
    manifestSha256,
    storageSetSha256,
    objects: entries.map(({ object }) => object),
    assets: entries.map(({ asset }) => asset),
  };
}

type MemoryCacheHarness = {
  cacheStorage: CacheStorage;
  entries: Map<string, { bytes: Uint8Array; headers: Headers }>;
  putChunkSizes: number[][];
  open: ReturnType<typeof vi.fn>;
  match: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
  globalMatch: ReturnType<typeof vi.fn>;
};

function cacheKey(input: RequestInfo | URL): string {
  return input instanceof Request ? input.url : String(input);
}

function memoryCacheStorage(): MemoryCacheHarness {
  const entries = new Map<string, { bytes: Uint8Array; headers: Headers }>();
  const putChunkSizes: number[][] = [];
  const match = vi.fn(async (input: RequestInfo | URL) => {
    const entry = entries.get(cacheKey(input));
    if (!entry) return undefined;
    return new Response(entry.bytes.slice().buffer, {
      status: 200,
      headers: new Headers(entry.headers),
    });
  });
  const deleteCached = vi.fn(async (input: RequestInfo | URL) => entries.delete(cacheKey(input)));
  const cache = {
    put: vi.fn(async (input: RequestInfo | URL, response: Response) => {
      const reader = response.body?.getReader();
      const chunks: Uint8Array[] = [];
      const chunkSizes: number[] = [];
      let total = 0;
      if (reader) {
        while (true) {
          const next = await reader.read();
          if (next.done) break;
          chunks.push(next.value.slice());
          chunkSizes.push(next.value.byteLength);
          total += next.value.byteLength;
        }
      }
      const bytes = new Uint8Array(total);
      let offset = 0;
      for (const chunk of chunks) {
        bytes.set(chunk, offset);
        offset += chunk.byteLength;
      }
      putChunkSizes.push(chunkSizes);
      entries.set(cacheKey(input), { bytes, headers: new Headers(response.headers) });
    }),
    match,
    delete: deleteCached,
  };
  const open = vi.fn(async (name: string) => {
    expect(name).toBe(CANONICAL_CONTENT_CACHE_NAME);
    return cache as unknown as Cache;
  });
  const globalMatch = vi.fn(async () => {
    throw new Error('the global CacheStorage matcher must not be used');
  });
  return {
    cacheStorage: {
      open,
      match: globalMatch,
    } as unknown as CacheStorage,
    entries,
    putChunkSizes,
    open,
    match,
    delete: deleteCached,
    globalMatch,
  };
}

function strictStreamingResponse(bytes: Uint8Array, chunkSizes: number[]): Response {
  let offset = 0;
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) {
      if (offset >= bytes.byteLength) {
        controller.close();
        return;
      }
      const requested = chunkSizes.shift() ?? bytes.byteLength - offset;
      const end = Math.min(bytes.byteLength, offset + requested);
      controller.enqueue(bytes.slice(offset, end));
      offset = end;
    },
  });
  const response = new Response(stream, {
    status: 200,
    headers: {
      'Content-Length': String(bytes.byteLength),
      'Content-Type': 'application/octet-stream',
    },
  });
  Object.defineProperties(response, {
    arrayBuffer: {
      value: vi.fn(async () => {
        throw new Error('arrayBuffer must not be called while installing');
      }),
    },
    clone: {
      value: vi.fn(() => {
        throw new Error('clone must not be called while installing');
      }),
    },
  });
  return response;
}

function objectUrl(releaseId: string, object: ContentObjectDescriptor): URL {
  return new URL(`https://assets.example.test/objects/${releaseId}/sha256/${object.sha256}`);
}

function createStore(options: {
  cache: MemoryCacheHarness;
  journal: MemoryCanonicalResourceJournal;
  fetch: typeof fetch;
  progress?: (chunkBytes: number) => void;
  maxConcurrentDownloads?: number;
}): CanonicalResourceStore {
  return new CanonicalResourceStore({
    cacheStorage: options.cache.cacheStorage,
    journal: options.journal,
    fetch: options.fetch,
    objectUrl,
    cacheKeyOrigin: 'https://onlyoffice.example.test',
    maxConcurrentDownloads: options.maxConcurrentDownloads,
    now: () => 100,
    onObjectProgress: ({ chunkBytes }) => options.progress?.(chunkBytes),
  });
}

describe('canonical resource store', () => {
  it('downloads once into one dedicated cache while hashing the original stream', async () => {
    const bytes = body('streamed-only-once');
    const target = release('release-a', [{ path: 'sdkjs/word/word.js', bytes }]);
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal({ now: () => 100 });
    const progress: number[] = [];
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      expect(new URL(String(input)).pathname).toBe(`/objects/release-a/sha256/${target.objects[0].sha256}`);
      expect(init).toMatchObject({ cache: 'no-store', credentials: 'omit' });
      return strictStreamingResponse(bytes, [2, 3, 4]);
    }) as unknown as typeof fetch;
    const store = createStore({
      cache,
      journal,
      fetch: fetchImplementation,
      progress: (chunkBytes) => progress.push(chunkBytes),
    });

    const [firstInstall, concurrentInstall] = await Promise.all([
      store.installAndActivate(target),
      store.installAndActivate(target),
    ]);

    expect(fetchImplementation).toHaveBeenCalledTimes(1);
    expect(firstInstall).toMatchObject({ releaseId: 'release-a', state: 'active' });
    expect(concurrentInstall).toMatchObject({ releaseId: 'release-a', state: 'active' });
    expect(cache.entries.size).toBe(1);
    expect(
      cache.entries.has(canonicalContentCacheKey(target.objects[0].sha256, 'https://onlyoffice.example.test')),
    ).toBe(true);
    expect(progress).toEqual([2, 3, 4, bytes.byteLength - 9]);
    expect(cache.putChunkSizes).toEqual([progress]);
    expect(cache.globalMatch).not.toHaveBeenCalled();
    expect(await store.checkHealth({ probe: true })).toMatchObject({
      releaseId: 'release-a',
      state: 'active',
      ready: true,
      missingObjects: [],
      probeSucceeded: true,
    });
  });

  it('completes a real readiness read without waiting for large Cache Storage cancellation', async () => {
    const bytes = body('persisted-large-object-probe');
    const target = release('release-probe', [{ path: 'sdkjs/word/word.js', bytes }]);
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const store = createStore({
      cache,
      journal,
      fetch: vi.fn(async () => strictStreamingResponse(bytes, [4])) as unknown as typeof fetch,
    });
    await store.installAndActivate(target);

    const cancel = vi.fn(() => new Promise<void>(() => undefined));
    cache.match.mockImplementation(async () => {
      const stream = new ReadableStream<Uint8Array>({
        start(controller) {
          controller.enqueue(bytes.slice(0, 4));
        },
        cancel,
      });
      return new Response(stream, {
        status: 200,
        headers: {
          'content-length': String(bytes.byteLength),
          [CANONICAL_CONTENT_DIGEST_HEADER]: target.objects[0].sha256,
          [CANONICAL_CONTENT_BYTES_HEADER]: String(bytes.byteLength),
        },
      });
    });

    await expect(store.probeRelease(target.releaseId)).resolves.toMatchObject({
      ready: true,
      probeSucceeded: true,
    });
    expect(cancel).toHaveBeenCalledTimes(1);
  });

  it('removes an integrity failure and never writes it to the verified ledger', async () => {
    const expected = body('expected-content');
    const corrupted = body('corrupt-content!');
    expect(corrupted.byteLength).toBe(expected.byteLength);
    const target = release('release-bad', [{ path: 'wasm/module.wasm', bytes: expected }]);
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const store = createStore({
      cache,
      journal,
      fetch: vi.fn(async () => strictStreamingResponse(corrupted, [3, 3])) as unknown as typeof fetch,
    });

    await expect(store.installAndActivate(target)).rejects.toMatchObject({
      code: 'integrity',
    });

    expect(cache.entries).toHaveLength(0);
    expect(await journal.listObjects()).toEqual([]);
    expect(await journal.getTransaction(target.releaseId)).toMatchObject({
      state: 'failed',
      failureCode: 'integrity',
    });
    expect(await journal.getActiveRelease()).toBeNull();
  });

  it('aliases unchanged A content into B without network or a second cache entry', async () => {
    const shared = body('same-content-across-releases');
    const releaseA = release('release-a', [{ path: 'sdkjs/word/word.js', bytes: shared }]);
    const releaseB = release(
      'release-b',
      [{ path: 'sdkjs/word/renamed-word.js', bytes: shared }],
      fixedDigest('c'),
      fixedDigest('d'),
    );
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const fetchImplementation = vi.fn(async () => strictStreamingResponse(shared, [5, 5])) as unknown as typeof fetch;
    const store = createStore({ cache, journal, fetch: fetchImplementation });

    await store.installAndActivate(releaseA);
    await store.installAndActivate(releaseB);

    expect(fetchImplementation).toHaveBeenCalledTimes(1);
    expect(cache.entries.size).toBe(1);
    expect(await journal.listAssetMappings('release-b')).toEqual([
      expect.objectContaining({
        releaseId: 'release-b',
        path: 'sdkjs/word/renamed-word.js',
        sourceReleaseId: 'release-a',
      }),
    ]);
    expect(await journal.getActiveRelease()).toMatchObject({ releaseId: 'release-b', state: 'active' });
    expect(await journal.getTransaction('release-a')).toMatchObject({ state: 'retained' });
  });

  it('checkpoints each verified object and resumes by fetching only the missing object', async () => {
    const first = body('first-object');
    const second = body('second-object');
    const target = release('release-resume', [
      { path: 'a.bin', bytes: first },
      { path: 'b.bin', bytes: second },
    ]);
    const bodies = new Map(target.objects.map((object, index) => [object.sha256, index === 0 ? first : second]));
    const calls = new Map<string, number>();
    let injectedFailure = false;
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
      const objectDigest = new URL(String(input)).pathname.split('/').at(-1)!;
      calls.set(objectDigest, (calls.get(objectDigest) || 0) + 1);
      if (!injectedFailure && calls.size === 2) {
        injectedFailure = true;
        throw new TypeError('simulated disconnect');
      }
      const bytes = bodies.get(objectDigest)!;
      return strictStreamingResponse(bytes, [4]);
    }) as unknown as typeof fetch;
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const store = createStore({ cache, journal, fetch: fetchImplementation });

    await expect(store.installAndActivate(target)).rejects.toMatchObject({
      code: 'network',
    });
    expect(cache.entries.size).toBe(1);
    expect(await journal.listObjects()).toHaveLength(1);
    expect(await journal.getTransaction(target.releaseId)).toMatchObject({
      state: 'failed',
      failureCode: 'network',
    });

    await store.installAndActivate(target);

    expect(fetchImplementation).toHaveBeenCalledTimes(3);
    expect([...calls.values()].sort()).toEqual([1, 2]);
    expect(cache.entries.size).toBe(2);
    expect(await journal.listObjects()).toHaveLength(2);
    expect(await journal.getActiveRelease()).toMatchObject({
      releaseId: target.releaseId,
      state: 'active',
    });
  });

  it('downloads independent immutable objects with a bounded four-way worker pool', async () => {
    const content = Array.from({ length: 6 }, (_, index) => ({
      path: `asset-${index}.bin`,
      bytes: body(`independent-object-${index}`),
    }));
    const target = release('release-concurrent', content);
    const bodies = new Map(target.objects.map((object, index) => [object.sha256, content[index].bytes]));
    let active = 0;
    let maximumActive = 0;
    let releaseFirstBatch!: () => void;
    const firstBatch = new Promise<void>((resolve) => {
      releaseFirstBatch = resolve;
    });
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
      active += 1;
      maximumActive = Math.max(maximumActive, active);
      if (active === 4) releaseFirstBatch();
      await firstBatch;
      active -= 1;
      const objectDigest = new URL(String(input)).pathname.split('/').at(-1)!;
      return strictStreamingResponse(bodies.get(objectDigest)!, [4]);
    }) as unknown as typeof fetch;
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const store = createStore({ cache, journal, fetch: fetchImplementation });

    await store.installAndActivate(target);

    expect(maximumActive).toBe(4);
    expect(fetchImplementation).toHaveBeenCalledTimes(content.length);
    expect(cache.entries.size).toBe(content.length);
  });

  it('keeps A active when preparing B fails and activates B only after a complete retry', async () => {
    const bytesA = body('release-a-object');
    const bytesB = body('release-b-object');
    const releaseA = release('release-a', [{ path: 'asset.bin', bytes: bytesA }]);
    const releaseB = release('release-b', [{ path: 'asset.bin', bytes: bytesB }], fixedDigest('e'), fixedDigest('f'));
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    let failB = true;
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
      const objectDigest = new URL(String(input)).pathname.split('/').at(-1)!;
      if (objectDigest === releaseB.objects[0].sha256 && failB) {
        return strictStreamingResponse(body('broken-release-b'), [5]);
      }
      const bytes = objectDigest === releaseA.objects[0].sha256 ? bytesA : bytesB;
      return strictStreamingResponse(bytes, [5]);
    }) as unknown as typeof fetch;
    const store = createStore({ cache, journal, fetch: fetchImplementation });

    await store.installAndActivate(releaseA);
    await expect(store.installAndActivate(releaseB)).rejects.toMatchObject({
      code: 'integrity',
    });

    expect(await journal.getActiveRelease()).toMatchObject({ releaseId: 'release-a', state: 'active' });
    expect(await journal.getTransaction('release-b')).toMatchObject({ state: 'failed' });
    expect(cache.entries.size).toBe(1);

    failB = false;
    await store.installAndActivate(releaseB);

    expect(await journal.getActiveRelease()).toMatchObject({ releaseId: 'release-b', state: 'active' });
    expect(await journal.getTransaction('release-a')).toMatchObject({ state: 'retained' });
    expect(cache.entries.size).toBe(2);
  });

  it('atomically reactivates retained A, restores B on failure, and keeps A readable throughout', async () => {
    const bytesA = body('retained-release-a');
    const bytesB = body('active-release-b');
    const releaseA = release('release-a', [{ path: 'sdkjs/a.bin', bytes: bytesA }]);
    const releaseB = release('release-b', [{ path: 'sdkjs/b.bin', bytes: bytesB }], fixedDigest('e'), fixedDigest('f'));
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
      const objectDigest = new URL(String(input)).pathname.split('/').at(-1)!;
      return strictStreamingResponse(objectDigest === releaseA.objects[0].sha256 ? bytesA : bytesB, [4]);
    }) as unknown as typeof fetch;
    const store = createStore({ cache, journal, fetch: fetchImplementation });

    await store.installAndActivate(releaseA);
    await store.installAndActivate(releaseB);
    expect(await journal.getTransaction(releaseA.releaseId)).toMatchObject({ state: 'retained' });
    expect(await store.matchObject(releaseA.objects[0])).toBeDefined();

    await store.installAndActivate(releaseA);
    expect(fetchImplementation).toHaveBeenCalledTimes(2);
    expect(await journal.getActiveRelease()).toMatchObject({ releaseId: releaseA.releaseId, state: 'active' });
    expect(await journal.getTransaction(releaseB.releaseId)).toMatchObject({ state: 'retained' });
    expect(await store.matchObject(releaseA.objects[0])).toBeDefined();

    await store.rollbackActivation(releaseA.releaseId, 'storage');
    expect(await journal.getActiveRelease()).toMatchObject({ releaseId: releaseB.releaseId, state: 'active' });
    const rolledBackReleaseA = await journal.getTransaction(releaseA.releaseId);
    expect(rolledBackReleaseA).toMatchObject({ state: 'retained' });
    expect(rolledBackReleaseA).not.toHaveProperty('failureCode');
    expect(await store.matchObject(releaseA.objects[0])).toBeDefined();

    await store.installAndActivate(releaseA);
    expect(fetchImplementation).toHaveBeenCalledTimes(2);
    expect(await journal.getActiveRelease()).toMatchObject({ releaseId: releaseA.releaseId, state: 'active' });
  });

  it('deeply verifies cached bytes, ignores trusted-looking headers, and removes only the corrupted object', async () => {
    const healthyBytes = body('healthy-object');
    const corruptedBytes = body('corrupted-object');
    const target = release('release-deep-verify', [
      { path: 'sdkjs/healthy.bin', bytes: healthyBytes },
      { path: 'sdkjs/corrupted.bin', bytes: corruptedBytes },
    ]);
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const objectBytes = new Map([
      [target.objects[0].sha256, healthyBytes],
      [target.objects[1].sha256, corruptedBytes],
    ]);
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
      const digest = new URL(String(input)).pathname.split('/').at(-1)!;
      return strictStreamingResponse(objectBytes.get(digest)!, [3, 4]);
    }) as unknown as typeof fetch;
    const store = createStore({ cache, journal, fetch: fetchImplementation });
    await store.installAndActivate(target);

    const corruptedObject = target.objects[1];
    const corruptedKey = canonicalContentCacheKey(corruptedObject.sha256, 'https://onlyoffice.example.test');
    const cached = cache.entries.get(corruptedKey)!;
    const tampered = cached.bytes.slice();
    tampered[0] ^= 0xff;
    cached.bytes = tampered;
    expect(cached.headers.get(CANONICAL_CONTENT_DIGEST_HEADER)).toBe(corruptedObject.sha256);
    expect(cached.headers.get(CANONICAL_CONTENT_BYTES_HEADER)).toBe(String(corruptedObject.bytes));
    await expect(store.checkHealth()).resolves.toMatchObject({ ready: true, missingObjects: [] });
    cache.delete.mockClear();

    await expect(store.verifyReleaseIntegrity()).resolves.toMatchObject({
      releaseId: target.releaseId,
      state: 'active',
      status: 'complete',
      ready: false,
      checkedObjects: 2,
      verifiedObjects: 1,
      verifiedBytes: healthyBytes.byteLength,
      failures: [
        {
          object: corruptedObject,
          code: 'integrity',
          removed: true,
          actualBytes: corruptedObject.bytes,
          actualSha256: digestBytes(tampered),
        },
      ],
    });
    expect(cache.delete).toHaveBeenCalledTimes(1);
    expect(cache.delete).toHaveBeenCalledWith(corruptedKey);
    expect(cache.entries.has(corruptedKey)).toBe(false);
    expect(
      cache.entries.has(canonicalContentCacheKey(target.objects[0].sha256, 'https://onlyoffice.example.test')),
    ).toBe(true);
    await expect(store.checkHealth()).resolves.toMatchObject({
      ready: false,
      missingObjects: [corruptedObject],
    });
  });

  it('returns a structured read failure and clears only the unreadable canonical object', async () => {
    const bytes = body('unreadable-object');
    const target = release('release-read-failure', [{ path: 'sdkjs/unreadable.bin', bytes }]);
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const store = createStore({
      cache,
      journal,
      fetch: vi.fn(async () => strictStreamingResponse(bytes, [4])) as unknown as typeof fetch,
    });
    await store.installAndActivate(target);
    const object = target.objects[0];
    const key = canonicalContentCacheKey(object.sha256, 'https://onlyoffice.example.test');
    cache.match.mockImplementationOnce(
      async () =>
        new Response(
          new ReadableStream<Uint8Array>({
            pull() {
              throw new Error('cache body read failed');
            },
          }),
          {
            status: 200,
            headers: new Headers(cache.entries.get(key)!.headers),
          },
        ),
    );
    cache.delete.mockClear();

    await expect(
      store.verifyReleaseIntegrity({
        releaseId: target.releaseId,
        objectSha256s: [object.sha256],
      }),
    ).resolves.toMatchObject({
      status: 'complete',
      ready: false,
      checkedObjects: 1,
      verifiedObjects: 0,
      failures: [
        {
          object,
          code: 'read',
          removed: true,
        },
      ],
    });
    expect(cache.delete).toHaveBeenCalledWith(key);
    expect(cache.entries.has(key)).toBe(false);
  });

  it('returns an aborted deep-verification result without deleting an object whose read was cancelled', async () => {
    const bytes = body('cancelled-object');
    const target = release('release-cancel-verify', [{ path: 'sdkjs/cancelled.bin', bytes }]);
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const store = createStore({
      cache,
      journal,
      fetch: vi.fn(async () => strictStreamingResponse(bytes, [4])) as unknown as typeof fetch,
    });
    await store.installAndActivate(target);
    const object = target.objects[0];
    const key = canonicalContentCacheKey(object.sha256, 'https://onlyoffice.example.test');
    let startedRead!: () => void;
    const readStarted = new Promise<void>((resolve) => {
      startedRead = resolve;
    });
    let releasePull: (() => void) | undefined;
    cache.match.mockImplementationOnce(
      async () =>
        new Response(
          new ReadableStream<Uint8Array>({
            pull() {
              startedRead();
              return new Promise<void>((resolve) => {
                releasePull = resolve;
              });
            },
            cancel() {
              releasePull?.();
            },
          }),
          {
            status: 200,
            headers: new Headers(cache.entries.get(key)!.headers),
          },
        ),
    );
    cache.delete.mockClear();
    const controller = new AbortController();
    const verification = store.verifyReleaseIntegrity({ signal: controller.signal });
    await readStarted;
    controller.abort();

    await expect(verification).resolves.toMatchObject({
      releaseId: target.releaseId,
      state: 'active',
      status: 'aborted',
      ready: false,
      failures: [],
    });
    expect(cache.delete).not.toHaveBeenCalled();
    expect(cache.entries.has(key)).toBe(true);
  });

  it('checks a thousand-object release with one journal view and one linear cache pass', async () => {
    const objectCount = 1_024;
    const objects: ContentObjectRecord[] = Array.from({ length: objectCount }, (_, index) => ({
      sha256: index.toString(16).padStart(64, '0'),
      bytes: 1,
      verifiedAt: 1,
    }));
    const mappings: AssetContentMapping[] = objects.map((object, index) => ({
      releaseId: 'release-thousand',
      path: `objects/${index}.bin`,
      assetSha256: object.sha256,
      assetBytes: object.bytes,
      representationId: 'whole',
      representationKind: 'whole',
      spans: [
        {
          objectSha256: object.sha256,
          objectOffset: 0,
          assetOffset: 0,
          bytes: object.bytes,
        },
      ],
    }));
    const transaction: ReleaseTransactionRecord = {
      releaseId: 'release-thousand',
      manifestSha256: fixedDigest('a'),
      storageSetSha256: fixedDigest('b'),
      planFingerprint: fixedDigest('c'),
      state: 'active',
      previousActiveReleaseId: null,
      requiredObjects: objects.map(({ sha256, bytes }) => ({ sha256, bytes })),
      plannedMappings: mappings,
      committedMappings: mappings,
      updatedAt: 1,
    };
    const getReleaseReadView = vi.fn(async () => ({ transaction, objects }));
    const listObjects = vi.fn(async () => {
      throw new Error('health checks must not scan the object table per object');
    });
    const journal = {
      getReleaseReadView,
      listObjects,
    } as unknown as CanonicalResourceJournal;
    const cache = memoryCacheStorage();
    for (const object of objects) {
      cache.entries.set(canonicalContentCacheKey(object.sha256, 'https://onlyoffice.example.test'), {
        bytes: Uint8Array.of(1),
        headers: new Headers({
          'content-length': String(object.bytes),
          [CANONICAL_CONTENT_DIGEST_HEADER]: object.sha256,
          [CANONICAL_CONTENT_BYTES_HEADER]: String(object.bytes),
        }),
      });
    }
    const store = new CanonicalResourceStore({
      cacheStorage: cache.cacheStorage,
      journal,
      fetch: vi.fn() as unknown as typeof fetch,
      objectUrl,
      cacheKeyOrigin: 'https://onlyoffice.example.test',
    });

    await expect(store.checkHealth({ releaseId: transaction.releaseId, probe: true })).resolves.toMatchObject({
      ready: true,
      missingObjects: [],
      probeSucceeded: true,
    });
    expect(getReleaseReadView).toHaveBeenCalledTimes(1);
    expect(listObjects).not.toHaveBeenCalled();
    expect(cache.open).toHaveBeenCalledTimes(1);
    expect(cache.match).toHaveBeenCalledTimes(objectCount);
  });

  it('rolls B back to A without GC deleting either release and reactivates B without network traffic', async () => {
    const bytesA = body('release-a-object');
    const bytesB = body('release-b-object');
    const releaseA = release('release-a', [{ path: 'asset.bin', bytes: bytesA }]);
    const releaseB = release('release-b', [{ path: 'asset.bin', bytes: bytesB }], fixedDigest('e'), fixedDigest('f'));
    const cache = memoryCacheStorage();
    const journal = new MemoryCanonicalResourceJournal();
    const fetchImplementation = vi.fn(async (input: RequestInfo | URL) => {
      const objectDigest = new URL(String(input)).pathname.split('/').at(-1)!;
      const bytes = objectDigest === releaseA.objects[0].sha256 ? bytesA : bytesB;
      return strictStreamingResponse(bytes, [5]);
    }) as unknown as typeof fetch;
    const store = createStore({ cache, journal, fetch: fetchImplementation });

    await store.installAndActivate(releaseA);
    await store.installAndActivate(releaseB);
    expect(fetchImplementation).toHaveBeenCalledTimes(2);

    await store.rollbackActivation(releaseB.releaseId, 'storage');
    expect(await journal.getActiveRelease()).toMatchObject({
      releaseId: releaseA.releaseId,
      state: 'active',
    });
    expect(await journal.getTransaction(releaseB.releaseId)).toMatchObject({
      state: 'failed',
      failureCode: 'storage',
    });

    await store.installAndActivate(releaseB);
    expect(fetchImplementation).toHaveBeenCalledTimes(2);
    expect(await journal.getActiveRelease()).toMatchObject({
      releaseId: releaseB.releaseId,
      state: 'active',
    });
    expect(cache.entries.size).toBe(2);
    cache.delete.mockClear();
    expect(store.planGarbageCollection()).toMatchObject({
      blocked: true,
      blockedReason: 'execution-disabled',
      deleteReleaseIds: [],
      deleteObjectSha256: [],
    });
    expect(cache.delete).not.toHaveBeenCalled();
  });
});
