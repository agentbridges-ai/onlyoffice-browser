import { sha256 } from '@noble/hashes/sha2.js';
import { describe, expect, it, vi } from 'vitest';
import {
  CANONICAL_CONTENT_CACHE_NAME,
  MemoryCanonicalResourceJournal,
  canonicalContentCacheKey,
} from '../../src/lib/canonical-resource-store';
import { CanonicalResourceInstaller } from '../../src/lib/canonical-resource-installer';
import {
  computeStorageSetSha256,
  type ReleaseAssetV5,
  type ReleaseManifestV5,
  type RequiredReleaseIdentity,
  type ResourceProfile,
} from '../../src/lib/release-resources';

function digest(bytes: Uint8Array | string): string {
  const input = typeof bytes === 'string' ? new TextEncoder().encode(bytes) : bytes;
  return Array.from(sha256(input), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

const profiles: ResourceProfile[] = ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'];

function release(
  releaseId: string,
  content: Array<{ path: string; value: string; profile?: ResourceProfile }>,
): { manifest: ReleaseManifestV5; text: string; sha256: string; objects: Map<string, Uint8Array> } {
  const objects = new Map<string, Uint8Array>();
  let packageOffset = 16;
  const packedContent = new Uint8Array(content.reduce((sum, entry) => sum + bytes(entry.value).byteLength, 0));
  let packedOffset = 0;
  const assets: ReleaseAssetV5[] = content.map((entry, index) => {
    const objectBytes = bytes(entry.value);
    packedContent.set(objectBytes, packedOffset);
    packedOffset += objectBytes.byteLength;
    const sha256 = digest(objectBytes);
    objects.set(sha256, objectBytes);
    const asset: ReleaseAssetV5 = {
      path: entry.path,
      bytes: objectBytes.byteLength,
      mime: entry.path.endsWith('.js') ? 'text/javascript' : 'application/octet-stream',
      sha256,
      profile: entry.profile || 'base',
      chunk: `logical-${index}`,
      packageOffset,
      representations: {
        whole: { sha256, bytes: objectBytes.byteLength },
      },
    };
    packageOffset += objectBytes.byteLength;
    return asset;
  });
  const packageBytes = packageOffset;
  const packageHeader = bytes('OOBPACK1TESTHEAD');
  expect(packageHeader.byteLength).toBe(16);
  const packagePayload = new Uint8Array(packageBytes);
  packagePayload.set(packageHeader);
  packagePayload.set(packedContent, packageHeader.byteLength);
  const packageDigest = digest(packagePayload);
  const headerDigest = digest(packageHeader);
  const contentDigest = digest(packedContent);
  objects.set(headerDigest, packageHeader);
  objects.set(contentDigest, packedContent);
  const pack = {
    format: 'onlyoffice-pack-v1' as const,
    path: 'office-resources.oobpack' as const,
    bytes: packageBytes,
    sha256: packageDigest,
    headerBytes: 16,
    segmentBytes: packageBytes,
    segments: [
      {
        id: headerDigest,
        offset: 0,
        bytes: packageHeader.byteLength,
        sha256: headerDigest,
      },
      {
        id: contentDigest,
        offset: packageHeader.byteLength,
        bytes: packedContent.byteLength,
        sha256: contentDigest,
      },
    ],
  };
  const profileEntries = Object.fromEntries(
    profiles.map((profile) => [
      profile,
      assets.filter((asset) => asset.profile === profile).map((asset) => asset.path),
    ]),
  ) as Record<ResourceProfile, string[]>;
  const manifest: ReleaseManifestV5 = {
    version: 5,
    releaseId,
    packageVersion: releaseId,
    hostBuildId: `host-${releaseId}`,
    shellRevision: `shell-${releaseId}`,
    runtimeManifestSha256: digest(`runtime:${releaseId}`),
    fontManifestSha256: digest(`fonts:${releaseId}`),
    x2t: {
      version: '9.3.0+1',
      commit: '0123456789abcdef',
      sha256: digest(`x2t:${releaseId}`),
    },
    profiles: profileEntries,
    chunks: profiles.map((profile) => ({
      id: `chunk-${profile}`,
      profile,
      bytes: assets.filter((asset) => asset.profile === profile).reduce((sum, asset) => sum + asset.bytes, 0),
      paths: profileEntries[profile],
    })),
    package: pack,
    assets,
    contentProtocol: {
      version: 1,
      digest: 'sha256',
      cacheKeyFormat: 'canonical-sha256-v1',
      storageSetSha256: computeStorageSetSha256(pack, assets),
      fastcdcPolicyId: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0',
    },
  };
  const text = JSON.stringify(manifest);
  return { manifest, text, sha256: digest(text), objects };
}

function releaseWithIndependentObjects(
  releaseId: string,
  content: Array<{ path: string; value: string; profile?: ResourceProfile }>,
): ReturnType<typeof release> {
  const generated = release(releaseId, content);
  const header = generated.manifest.package.segments[0];
  generated.manifest.package.segments = [
    header,
    ...generated.manifest.assets.map((asset) => ({
      id: asset.sha256,
      offset: asset.packageOffset,
      bytes: asset.bytes,
      sha256: asset.sha256,
    })),
  ];
  generated.manifest.contentProtocol.storageSetSha256 = computeStorageSetSha256(
    generated.manifest.package,
    generated.manifest.assets,
  );
  generated.text = JSON.stringify(generated.manifest);
  generated.sha256 = digest(generated.text);
  return generated;
}

function releaseWithFastCdc(releaseId: string, path: string, chunks: string[]): ReturnType<typeof release> {
  const value = chunks.join('');
  const generated = release(releaseId, [{ path, value }]);
  let offset = 0;
  generated.manifest.assets[0].representations.fastcdc = {
    algorithm: 'fastcdc-v2020',
    minBytes: 65_536,
    averageBytes: 262_144,
    maxBytes: 1_048_576,
    normalization: 1,
    seed: 0,
    chunks: chunks.map((chunk) => {
      const chunkBytes = bytes(chunk);
      const descriptor = {
        offset,
        bytes: chunkBytes.byteLength,
        sha256: digest(chunkBytes),
      };
      generated.objects.set(descriptor.sha256, chunkBytes);
      offset += chunkBytes.byteLength;
      return descriptor;
    }),
  };
  generated.manifest.contentProtocol.storageSetSha256 = computeStorageSetSha256(
    generated.manifest.package,
    generated.manifest.assets,
  );
  generated.text = JSON.stringify(generated.manifest);
  generated.sha256 = digest(generated.text);
  return generated;
}

type MemoryCache = {
  cacheStorage: CacheStorage;
  entries: Map<string, { bytes: Uint8Array; headers: Headers }>;
  delete: (key: string) => Promise<boolean>;
};

function requestKey(input: RequestInfo | URL): string {
  return input instanceof Request ? input.url : String(input);
}

function memoryCache(): MemoryCache {
  const entries = new Map<string, { bytes: Uint8Array; headers: Headers }>();
  const cache = {
    async put(input: RequestInfo | URL, response: Response) {
      const value = new Uint8Array(await response.arrayBuffer());
      entries.set(requestKey(input), { bytes: value, headers: new Headers(response.headers) });
    },
    async match(input: RequestInfo | URL) {
      const entry = entries.get(requestKey(input));
      if (!entry) return undefined;
      return new Response(entry.bytes.slice().buffer, {
        status: 200,
        headers: new Headers(entry.headers),
      });
    },
    async keys() {
      return [...entries.keys()].map((key) => new Request(key));
    },
    async delete(input: RequestInfo | URL) {
      return entries.delete(requestKey(input));
    },
  };
  return {
    entries,
    delete: (key) => cache.delete(key),
    cacheStorage: {
      async open(name: string) {
        expect(name).toBe(CANONICAL_CONTENT_CACHE_NAME);
        return cache as unknown as Cache;
      },
      async match() {
        throw new Error('the canonical installer must not use the global cache matcher');
      },
    } as unknown as CacheStorage,
  };
}

function streamingResponse(value: Uint8Array): Response {
  let offset = 0;
  const split = Math.max(1, Math.floor(value.byteLength / 2));
  return new Response(
    new ReadableStream<Uint8Array>({
      pull(controller) {
        if (offset >= value.byteLength) {
          controller.close();
          return;
        }
        const next = Math.min(value.byteLength, offset + split);
        controller.enqueue(value.slice(offset, next));
        offset = next;
      },
    }),
    {
      status: 200,
      headers: {
        'Content-Length': String(value.byteLength),
        'Content-Type': 'application/octet-stream',
      },
    },
  );
}

function server(initial: ReturnType<typeof release>) {
  let current = initial;
  const releases = new Map([[initial.manifest.releaseId, initial]]);
  const objectRequests: string[] = [];
  const manifestRequests: string[] = [];
  const fetch = vi.fn(async (input: RequestInfo | URL) => {
    const url = new URL(String(input));
    if (url.pathname === '/channels/stable-v5.json') {
      return Response.json({
        version: 1,
        releaseId: current.manifest.releaseId,
        manifestUrl: `/releases/${current.manifest.releaseId}/manifest.json`,
        manifestSha256: current.sha256,
      });
    }
    const manifestMatch = /^\/releases\/([^/]+)\/manifest\.json$/.exec(url.pathname);
    if (manifestMatch) {
      const releaseId = decodeURIComponent(manifestMatch[1]);
      manifestRequests.push(releaseId);
      const selected = releases.get(releaseId);
      return selected ? new Response(selected.text, { status: 200 }) : new Response(null, { status: 404 });
    }
    const objectMatch = /^\/objects\/([^/]+)\/sha256\/([a-f0-9]{64})$/.exec(url.pathname);
    if (objectMatch) {
      const releaseId = decodeURIComponent(objectMatch[1]);
      objectRequests.push(`${releaseId}:${objectMatch[2]}`);
      const object = releases.get(releaseId)?.objects.get(objectMatch[2]);
      return object ? streamingResponse(object) : new Response(null, { status: 404 });
    }
    return new Response(null, { status: 404 });
  }) as unknown as typeof globalThis.fetch;
  return {
    fetch,
    objectRequests,
    manifestRequests,
    setCurrent(next: ReturnType<typeof release>) {
      releases.set(next.manifest.releaseId, next);
      current = next;
    },
  };
}

function requiredIdentity(target: ReturnType<typeof release>): RequiredReleaseIdentity {
  return {
    releaseId: target.manifest.releaseId,
    manifestSha256: target.sha256,
    packageVersion: target.manifest.packageVersion,
    hostBuildId: target.manifest.hostBuildId,
  };
}

function installerHarness(initial: ReturnType<typeof release>, requiredReleaseIdentity?: RequiredReleaseIdentity) {
  const network = server(initial);
  const cache = memoryCache();
  const journal = new MemoryCanonicalResourceJournal({ now: () => 100 });
  const installer = new CanonicalResourceInstaller({
    assetBaseUrl: 'https://onlyoffice.example.test/',
    fetch: network.fetch,
    cacheStorage: cache.cacheStorage,
    journal,
    retryDelaysMs: [],
    timeoutMs: 1_000,
    now: () => 100,
    requiredReleaseIdentity,
  });
  return { installer, network, cache, journal };
}

describe('canonical resource installer', () => {
  it('uses stable-v5 and installs the complete release as one incrementally planned task', async () => {
    const target = release('release-a', [
      { path: 'sdkjs/word/word.js', value: 'word-runtime' },
      { path: 'fonts/aptos.ttf', value: 'aptos-font', profile: 'fonts-basic' },
    ]);
    const { installer, network, cache } = installerHarness(target);
    const phases: string[] = [];
    installer.subscribeInstaller((snapshot) => phases.push(snapshot.phase));

    await installer.initialize();
    const plan = await installer.plan({ scope: 'recommended' });

    expect(plan).toMatchObject({
      releaseId: 'release-a',
      scope: 'recommended',
      profiles,
      downloadBytes: bytes('word-runtime').byteLength + bytes('aptos-font').byteLength,
      reusedBytes: 0,
    });
    expect(plan.totalBytes).toBe(plan.downloadBytes);

    await installer.apply(plan);
    await installer.checkHealth();

    expect(network.objectRequests).toHaveLength(1);
    expect(network.objectRequests[0]).toBe(
      `release-a:${digest(new Uint8Array([...bytes('word-runtime'), ...bytes('aptos-font')]))}`,
    );
    expect(cache.entries).toHaveLength(1);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      availableRelease: 'release-a',
      availablePackageVersion: 'release-a',
      readiness: 'ready',
      phase: 'idle',
      downloadedBytes: plan.downloadBytes,
      verifiedBytes: plan.downloadBytes,
      errorCode: null,
    });
    expect(installer.getInstalledPaths()).toEqual(['fonts/aptos.ttf', 'sdkjs/word/word.js']);
    expect(phases).toContain('downloading');
    expect(phases).toContain('verifying');
    expect(phases).toContain('activating');
  });

  it('streams a cold package into planned FastCDC CAS chunks when release evidence enables them', async () => {
    const target = releaseWithFastCdc('release-fastcdc-cold', 'wasm/runtime.bin', [
      'fastcdc-chunk-one',
      'fastcdc-chunk-two',
    ]);
    const { installer, network, cache, journal } = installerHarness(target);

    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);

    const chunks = target.manifest.assets[0].representations.fastcdc!.chunks;
    expect(network.objectRequests).toEqual([`release-fastcdc-cold:${target.manifest.package.segments[1].sha256}`]);
    expect(cache.entries).toHaveLength(chunks.length);
    expect([...cache.entries.keys()].sort()).toEqual(
      chunks.map((chunk) => canonicalContentCacheKey(chunk.sha256, 'https://onlyoffice.example.test')).sort(),
    );
    expect(
      cache.entries.has(canonicalContentCacheKey(target.manifest.assets[0].sha256, 'https://onlyoffice.example.test')),
    ).toBe(false);
    expect((await journal.listAssetMappings(target.manifest.releaseId))[0]).toMatchObject({
      representationKind: 'fastcdc',
      spans: chunks.map((chunk) => ({
        objectSha256: chunk.sha256,
        objectOffset: 0,
        assetOffset: chunk.offset,
        bytes: chunk.bytes,
      })),
    });
  });

  it('retries a package after a post-200 stream reset without refetching verified chunks', async () => {
    const target = releaseWithFastCdc('release-fastcdc-stream-retry', 'wasm/runtime.bin', [
      'fastcdc-chunk-one',
      'fastcdc-chunk-two',
    ]);
    const network = server(target);
    const cache = memoryCache();
    const journal = new MemoryCanonicalResourceJournal({ now: () => 100 });
    const failedSegment = target.manifest.package.segments[1].sha256;
    let failOnce = true;
    const fetch = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const response = await network.fetch(input, init);
      const url = new URL(String(input));
      if (!failOnce || !url.pathname.endsWith(`/sha256/${failedSegment}`) || !response.body) return response;
      failOnce = false;
      const reader = response.body.getReader();
      let first = true;
      const body = new ReadableStream<Uint8Array>({
        async pull(controller) {
          const next = await reader.read();
          if (next.done) {
            controller.close();
            return;
          }
          if (!first) {
            await reader.cancel();
            controller.error(new TypeError('ERR_HTTP2_PROTOCOL_ERROR'));
            return;
          }
          first = false;
          controller.enqueue(next.value);
        },
        async cancel(reason) {
          await reader.cancel(reason);
        },
      });
      return new Response(body, { status: response.status, headers: response.headers });
    }) as unknown as typeof globalThis.fetch;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      retryDelaysMs: [0],
      timeoutMs: 1_000,
      now: () => 100,
    });

    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);

    expect(network.objectRequests.filter((request) => request.endsWith(failedSegment))).toHaveLength(2);
    expect(cache.entries).toHaveLength(target.manifest.assets[0].representations.fastcdc!.chunks.length);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      phase: 'idle',
      errorCode: null,
      failedResources: [],
    });
  });

  it('wakes a stalled package reader on timeout and retries only the interrupted segment', async () => {
    const target = releaseWithFastCdc('release-fastcdc-stalled-reader', 'wasm/runtime.bin', [
      'fastcdc-chunk-one',
      'fastcdc-chunk-two',
    ]);
    const network = server(target);
    const cache = memoryCache();
    const journal = new MemoryCanonicalResourceJournal({ now: () => 100 });
    const stalledSegment = target.manifest.package.segments[1].sha256;
    let stallOnce = true;
    const fetch = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const response = await network.fetch(input, init);
      const url = new URL(String(input));
      if (!stallOnce || !url.pathname.endsWith(`/sha256/${stalledSegment}`) || !response.body) return response;
      stallOnce = false;
      const reader = response.body.getReader();
      let first = true;
      const body = new ReadableStream<Uint8Array>({
        async pull(controller) {
          if (!first) {
            await new Promise<void>(() => undefined);
            return;
          }
          first = false;
          const next = await reader.read();
          if (next.done) {
            controller.close();
            return;
          }
          controller.enqueue(next.value);
        },
        async cancel(reason) {
          await reader.cancel(reason);
        },
      });
      return new Response(body, { status: response.status, headers: response.headers });
    }) as unknown as typeof globalThis.fetch;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      retryDelaysMs: [0],
      timeoutMs: 20,
      now: () => 100,
    });

    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);

    expect(network.objectRequests.filter((request) => request.endsWith(stalledSegment))).toHaveLength(2);
    expect(cache.entries).toHaveLength(target.manifest.assets[0].representations.fastcdc!.chunks.length);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      phase: 'idle',
      errorCode: null,
      failedResources: [],
    });
  });

  it('reuses unchanged content across releases and downloads only the changed object', async () => {
    const releaseA = release('release-a', [
      { path: 'shared.js', value: 'shared' },
      { path: 'changed.js', value: 'before' },
    ]);
    const releaseB = release('release-b', [
      { path: 'shared-renamed.js', value: 'shared' },
      { path: 'changed.js', value: 'after' },
    ]);
    const { installer, network, cache } = installerHarness(releaseA);

    const planA = await installer.plan({ scope: 'all' });
    await installer.apply(planA);
    network.setCurrent(releaseB);
    await installer.checkForUpdates();
    const planB = await installer.plan({ scope: 'fonts' });

    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      availableRelease: 'release-b',
      availablePackageVersion: 'release-b',
      readiness: 'update-available',
    });
    expect(planB.profiles).toEqual(profiles);
    expect(planB.downloadBytes).toBe(bytes('after').byteLength);
    expect(planB.reusedBytes).toBe(bytes('shared').byteLength + bytes('before').byteLength);

    await installer.apply(planB);

    expect(network.objectRequests).toHaveLength(2);
    expect(network.objectRequests.filter((request) => request.endsWith(digest(bytes('shared'))))).toHaveLength(0);
    expect(cache.entries).toHaveLength(2);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-b',
      availableRelease: 'release-b',
      readiness: 'ready',
    });
    expect(installer.getInstalledPaths()).toEqual(['changed.js', 'shared-renamed.js']);
  });

  it('resumes a failed cold package ingest from verified whole objects without refetching completed segments', async () => {
    const target = releaseWithIndependentObjects('release-cold-resume', [
      { path: 'one.js', value: 'resume-one' },
      { path: 'two.js', value: 'resume-two' },
    ]);
    const network = server(target);
    const cache = memoryCache();
    const journal = new MemoryCanonicalResourceJournal({ now: () => 100 });
    const failedDigest = target.manifest.assets[1].sha256;
    let failOnce = true;
    const fetch = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input));
      if (failOnce && url.pathname.endsWith(`/sha256/${failedDigest}`)) {
        failOnce = false;
        network.objectRequests.push(`release-cold-resume:${failedDigest}`);
        return new Response(null, { status: 503 });
      }
      return network.fetch(input, init);
    }) as unknown as typeof globalThis.fetch;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      retryDelaysMs: [],
      timeoutMs: 1_000,
      now: () => 100,
    });

    const firstPlan = await installer.plan({ scope: 'all' });
    await expect(installer.apply(firstPlan)).rejects.toMatchObject({ code: 'network' });
    const completedDigest = target.manifest.assets[0].sha256;
    expect(cache.entries.has(canonicalContentCacheKey(completedDigest, 'https://onlyoffice.example.test'))).toBe(true);
    expect(cache.entries.has(canonicalContentCacheKey(failedDigest, 'https://onlyoffice.example.test'))).toBe(false);

    const resumedPlan = await installer.plan({ scope: 'all' });
    expect(resumedPlan).toMatchObject({
      downloadBytes: target.manifest.assets[1].bytes,
      reusedBytes: target.manifest.assets[0].bytes,
    });
    await installer.apply(resumedPlan);

    expect(network.objectRequests).toEqual([
      `release-cold-resume:${completedDigest}`,
      `release-cold-resume:${failedDigest}`,
      `release-cold-resume:${failedDigest}`,
    ]);
    expect(network.objectRequests.filter((request) => request.endsWith(completedDigest))).toHaveLength(1);
    expect(network.objectRequests.filter((request) => request.endsWith(failedDigest))).toHaveLength(2);
    expect(cache.entries).toHaveLength(2);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-cold-resume',
      readiness: 'ready',
      phase: 'idle',
    });
  });

  it('reports a missing canonical object and repair fetches only that object', async () => {
    const target = release('release-repair', [
      { path: 'one.js', value: 'object-one' },
      { path: 'two.js', value: 'object-two' },
    ]);
    const { installer, network, cache } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const missingDigest = digest(new Uint8Array([...bytes('object-one'), ...bytes('object-two')]));
    await cache.delete(canonicalContentCacheKey(missingDigest, 'https://onlyoffice.example.test'));
    network.objectRequests.length = 0;

    await installer.checkHealth({ deep: true });

    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'repair-needed',
      errorCode: 'storage',
      failedResources: [
        {
          path: `objects/sha256/${missingDigest}`,
          code: 'storage',
          attempts: 1,
        },
      ],
    });

    await installer.repair({ scope: 'required' });

    expect(network.objectRequests).toEqual([`release-repair:${missingDigest}`]);
    expect(cache.entries).toHaveLength(1);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-repair',
      readiness: 'ready',
      errorCode: null,
    });
  });

  it('deeply verifies trusted-looking cached bodies and repairs only the corrupted object', async () => {
    const target = releaseWithIndependentObjects('release-deep-repair', [
      { path: 'one.js', value: 'healthy-one' },
      { path: 'two.js', value: 'corrupt-two' },
    ]);
    const { installer, network, cache } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const corruptedDigest = target.manifest.assets[1].sha256;
    const corruptedKey = canonicalContentCacheKey(corruptedDigest, 'https://onlyoffice.example.test');
    const cached = cache.entries.get(corruptedKey)!;
    cached.bytes[0] ^= 0xff;
    network.objectRequests.length = 0;
    const repairFailures: Array<{ path: string; code: string }> = [];
    installer.subscribeInstaller((snapshot) => {
      if (snapshot.phase === 'repairing') {
        repairFailures.push(...snapshot.failedResources.map(({ path, code }) => ({ path, code })));
      }
    });

    // The lightweight health check deliberately trusts the verified journal
    // and response metadata. A user-requested repair performs the deep read.
    await installer.checkHealth();
    expect(installer.getInstallerSnapshot().readiness).toBe('ready');
    await installer.repair({ scope: 'required' });

    expect(network.objectRequests).toEqual([`release-deep-repair:${corruptedDigest}`]);
    expect(repairFailures).toContainEqual({
      path: `objects/sha256/${corruptedDigest}`,
      code: 'integrity',
    });
    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-deep-repair',
      readiness: 'ready',
      phase: 'idle',
      failedResources: [],
      errorCode: null,
    });
  });

  it('reports the exact corrupted object when its targeted repair download fails', async () => {
    const target = releaseWithIndependentObjects('release-deep-repair-failure', [
      { path: 'one.js', value: 'healthy-one' },
      { path: 'two.js', value: 'corrupt-two' },
    ]);
    const { installer, network, cache } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const corruptedDigest = target.manifest.assets[1].sha256;
    const corruptedKey = canonicalContentCacheKey(corruptedDigest, 'https://onlyoffice.example.test');
    cache.entries.get(corruptedKey)!.bytes[0] ^= 0xff;
    vi.mocked(network.fetch).mockClear();
    vi.mocked(network.fetch).mockResolvedValue(new Response(null, { status: 503 }));

    await expect(installer.repair({ scope: 'required' })).rejects.toMatchObject({ code: 'network' });

    expect(network.fetch).toHaveBeenCalledTimes(1);
    expect(new URL(String(vi.mocked(network.fetch).mock.calls[0][0])).pathname).toBe(
      `/objects/release-deep-repair-failure/sha256/${corruptedDigest}`,
    );
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'error',
      phase: 'idle',
      errorCode: 'network',
      failedResources: [
        {
          path: `objects/sha256/${corruptedDigest}`,
          code: 'network',
          attempts: 1,
        },
      ],
    });
  });

  it('performs a healthy deep repair with zero network access after restart', async () => {
    const target = releaseWithIndependentObjects('release-healthy-repair', [
      { path: 'one.js', value: 'healthy-one' },
      { path: 'two.js', value: 'healthy-two' },
    ]);
    const { installer, network, cache, journal } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    network.objectRequests.length = 0;
    network.manifestRequests.length = 0;
    const restarted = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch: network.fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      retryDelaysMs: [],
      requiredReleaseIdentity: requiredIdentity(target),
    });
    vi.mocked(network.fetch).mockClear();

    await restarted.repair({ scope: 'installed' });

    expect(network.fetch).not.toHaveBeenCalled();
    expect(network.objectRequests).toEqual([]);
    expect(network.manifestRequests).toEqual([]);
    expect(restarted.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-healthy-repair',
      readiness: 'ready',
      phase: 'idle',
      errorCode: null,
    });
  });

  it('pauses and resumes a deep cache verification without downloading again', async () => {
    const target = release('release-pause-repair', [{ path: 'one.js', value: 'pause-deep-verification' }]);
    const { installer, network, cache } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    network.objectRequests.length = 0;
    const object = target.manifest.package.segments[1];
    const key = canonicalContentCacheKey(object.sha256, 'https://onlyoffice.example.test');
    const cacheApi = await cache.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    const originalMatch = cacheApi.match.bind(cacheApi);
    let readStarted!: () => void;
    const started = new Promise<void>((resolve) => {
      readStarted = resolve;
    });
    let releasePull: (() => void) | undefined;
    vi.spyOn(cacheApi, 'match').mockImplementationOnce(async () => {
      const entry = cache.entries.get(key)!;
      return new Response(
        new ReadableStream<Uint8Array>({
          pull() {
            readStarted();
            return new Promise<void>((resolve) => {
              releasePull = resolve;
            });
          },
          cancel() {
            releasePull?.();
          },
        }),
        { status: 200, headers: new Headers(entry.headers) },
      );
    });

    const repairing = installer.repair({ scope: 'required' });
    await started;
    await expect(installer.repair({ scope: 'all' })).rejects.toMatchObject({ code: 'incompatible' });
    installer.pause();
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'paused',
      phase: 'paused',
      canResume: true,
    });
    const resuming = installer.resume();
    await Promise.all([repairing, resuming]);

    expect(cacheApi.match).toHaveBeenCalled();
    expect(await originalMatch(key)).toBeDefined();
    expect(network.objectRequests).toEqual([]);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      phase: 'idle',
      canResume: false,
    });
  });

  it('cancels a deep verification without deleting or redownloading the interrupted object', async () => {
    const target = release('release-cancel-repair', [{ path: 'one.js', value: 'cancel-deep-verification' }]);
    const { installer, network, cache } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    network.objectRequests.length = 0;
    const object = target.manifest.package.segments[1];
    const key = canonicalContentCacheKey(object.sha256, 'https://onlyoffice.example.test');
    const cacheApi = await cache.cacheStorage.open(CANONICAL_CONTENT_CACHE_NAME);
    let readStarted!: () => void;
    const started = new Promise<void>((resolve) => {
      readStarted = resolve;
    });
    let releasePull: (() => void) | undefined;
    vi.spyOn(cacheApi, 'match').mockImplementationOnce(async () => {
      const entry = cache.entries.get(key)!;
      return new Response(
        new ReadableStream<Uint8Array>({
          pull() {
            readStarted();
            return new Promise<void>((resolve) => {
              releasePull = resolve;
            });
          },
          cancel() {
            releasePull?.();
          },
        }),
        { status: 200, headers: new Headers(entry.headers) },
      );
    });

    const repairing = installer.repair({ scope: 'required' });
    await started;
    installer.cancel();

    await expect(repairing).rejects.toMatchObject({ code: 'aborted' });
    expect(cache.entries.has(key)).toBe(true);
    expect(network.objectRequests).toEqual([]);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'repair-needed',
      phase: 'idle',
      errorCode: 'aborted',
      canRetry: true,
    });
  });

  it('aborts an in-flight stream on pause and resumes from the journaled transaction', async () => {
    const target = release('release-pause', [{ path: 'large.js', value: 'pause-and-resume-stream' }]);
    const network = server(target);
    const cache = memoryCache();
    const journal = new MemoryCanonicalResourceJournal();
    let firstObjectAttempt = true;
    const fetch = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input));
      const match = /^\/objects\/([^/]+)\/sha256\/([a-f0-9]{64})$/.exec(url.pathname);
      if (!match || !firstObjectAttempt) return network.fetch(input, init);
      firstObjectAttempt = false;
      network.objectRequests.push(`${match[1]}:${match[2]}`);
      const value = target.objects.get(match[2])!;
      let offset = 0;
      return new Response(
        new ReadableStream<Uint8Array>({
          pull(controller) {
            if (offset === 0) {
              controller.enqueue(value.slice(0, 1));
              offset = 1;
              return;
            }
            return new Promise<void>((resolve) => {
              const abort = () => {
                controller.error(new DOMException('The operation was aborted', 'AbortError'));
                resolve();
              };
              if (init?.signal?.aborted) {
                abort();
              } else {
                init?.signal?.addEventListener('abort', abort, { once: true });
              }
            });
          },
        }),
        {
          status: 200,
          headers: {
            'Content-Length': String(value.byteLength),
            'Content-Type': 'text/javascript',
          },
        },
      );
    }) as unknown as typeof globalThis.fetch;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      retryDelaysMs: [],
      timeoutMs: 1_000,
    });
    const plan = await installer.plan({ scope: 'all' });
    let resolveProgress!: () => void;
    const progressSeen = new Promise<void>((resolve) => {
      resolveProgress = resolve;
    });
    const unsubscribe = installer.subscribeInstaller((snapshot) => {
      if (snapshot.downloadedBytes > 0) resolveProgress();
    });

    const applying = installer.apply(plan);
    await progressSeen;
    installer.pause();
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'paused',
      phase: 'paused',
      canResume: true,
    });

    const resuming = installer.resume();
    await Promise.all([applying, resuming]);
    unsubscribe();

    expect(network.objectRequests).toHaveLength(2);
    expect(new Set(network.objectRequests)).toEqual(new Set([`release-pause:${target.manifest.assets[0].sha256}`]));
    expect(cache.entries).toHaveLength(1);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-pause',
      readiness: 'ready',
      phase: 'idle',
      canResume: false,
    });
  });

  it('restores a healthy active release while offline without requiring the stable channel', async () => {
    const target = release('release-offline', [{ path: 'offline.js', value: 'offline-runtime' }]);
    const { installer, cache, journal } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const offlineFetch = vi.fn(async () => {
      throw new TypeError('offline');
    }) as unknown as typeof globalThis.fetch;
    let online = true;
    const restarted = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch: offlineFetch,
      cacheStorage: cache.cacheStorage,
      journal,
      retryDelaysMs: [],
      requiredReleaseIdentity: requiredIdentity(target),
      online: () => online,
    });

    await restarted.initialize();

    expect(offlineFetch).not.toHaveBeenCalled();
    expect(restarted.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-offline',
      targetRelease: 'release-offline',
      availableRelease: null,
      readiness: 'ready',
      errorCode: null,
    });
    expect(restarted.getInstalledPaths()).toEqual(['offline.js']);

    online = false;
    await restarted.checkForUpdates();
    expect(offlineFetch).not.toHaveBeenCalled();
    expect(restarted.getInstallerSnapshot().readiness).toBe('ready');
  });

  it('rejects a peer snapshot that disagrees with the transactional active release', async () => {
    const target = release('release-active', [{ path: 'active.js', value: 'active-runtime' }]);
    const network = server(target);
    const cache = memoryCache();
    const journal = new MemoryCanonicalResourceJournal({ now: () => 100 });
    const events = new EventTarget();
    const broadcast = {
      addEventListener: events.addEventListener.bind(events),
      postMessage: vi.fn(),
    } as unknown as BroadcastChannel;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch: network.fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      broadcast,
      retryDelaysMs: [],
    });
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const stalePeerSnapshot = {
      ...installer.getInstallerSnapshot(),
      installedRelease: 'release-stale-peer',
      targetRelease: 'release-stale-peer',
      availableRelease: 'release-stale-peer',
    };

    events.dispatchEvent(
      new MessageEvent('message', {
        data: {
          type: 'canonical-resource-snapshot-v1',
          sourceId: 'peer-installer',
          snapshot: stalePeerSnapshot,
          requiredReleaseIdentity: null,
        },
      }),
    );
    await Promise.resolve();
    await Promise.resolve();

    expect(installer.getInstallerSnapshot().installedRelease).toBe('release-active');
    expect(installer.getInstalledPaths()).toEqual(['active.js']);

    events.dispatchEvent(
      new MessageEvent('message', {
        data: {
          type: 'canonical-resource-snapshot-v1',
          sourceId: 'stale-availability-peer',
          snapshot: {
            ...installer.getInstallerSnapshot(),
            targetRelease: 'release-newer',
            availableRelease: 'release-newer',
            readiness: 'update-available',
          },
          requiredReleaseIdentity: null,
        },
      }),
    );
    await Promise.resolve();
    await Promise.resolve();

    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-active',
      targetRelease: 'release-active',
      availableRelease: 'release-active',
      readiness: 'ready',
    });
  });

  it('keeps the Piwork-authorized target ready when stable points at a different upstream release', async () => {
    const releaseA = release('release-authorized', [{ path: 'shared.js', value: 'authorized-content' }]);
    const releaseB = release('release-unknown-stable', [{ path: 'shared.js', value: 'upstream-content' }]);
    const { installer, network } = installerHarness(releaseA, requiredIdentity(releaseA));

    await installer.initialize();
    const installPlan = await installer.plan({ scope: 'all' });
    await installer.apply(installPlan);
    network.setCurrent(releaseB);
    network.manifestRequests.length = 0;

    await installer.checkForUpdates();
    const stillAuthorized = await installer.plan({ scope: 'all' });

    expect(installer.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-authorized',
      targetRelease: 'release-authorized',
      availableRelease: 'release-unknown-stable',
      availablePackageVersion: 'release-unknown-stable',
      readiness: 'ready',
      errorCode: null,
    });
    expect(stillAuthorized).toMatchObject({
      releaseId: 'release-authorized',
      downloadBytes: 0,
    });
    expect(network.manifestRequests).toContain('release-unknown-stable');
    expect(network.manifestRequests.at(-1)).toBe('release-authorized');
  });

  it('rejects a tampered exact manifest before planning or activation', async () => {
    const target = release('release-tampered', [{ path: 'one.js', value: 'object-one' }]);
    const cache = memoryCache();
    const journal = new MemoryCanonicalResourceJournal();
    const fetch = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (url.pathname === `/releases/${target.manifest.releaseId}/manifest.json`) {
        return new Response(`${target.text}\n`, { status: 200 });
      }
      return new Response(null, { status: 404 });
    }) as unknown as typeof globalThis.fetch;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      requiredReleaseIdentity: requiredIdentity(target),
      retryDelaysMs: [],
    });

    await expect(installer.plan({ scope: 'all' })).rejects.toMatchObject({ code: 'manifest' });
    expect(installer.getInstallerSnapshot()).toMatchObject({
      targetRelease: target.manifest.releaseId,
      readiness: 'error',
      errorCode: 'manifest',
    });
    expect(await journal.getActiveRelease()).toBeNull();
  });

  it.each([
    ['packageVersion', { packageVersion: '0.0.0-wrong' }],
    ['hostBuildId', { hostBuildId: 'host-wrong' }],
  ] as const)('rejects a manifest whose %s is outside the authorized identity', async (_field, override) => {
    const target = release('release-identity-fields', [{ path: 'one.js', value: 'object-one' }]);
    const { network, cache, journal } = installerHarness(target);
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch: network.fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      requiredReleaseIdentity: {
        ...requiredIdentity(target),
        ...override,
      },
      retryDelaysMs: [],
    });

    await expect(installer.plan({ scope: 'all' })).rejects.toMatchObject({ code: 'incompatible' });
    expect(await journal.getActiveRelease()).toBeNull();
  });

  it('does not report ready when the active release id has a different authorized manifest digest', async () => {
    const target = release('release-same-id', [{ path: 'one.js', value: 'object-one' }]);
    const { installer, network, cache, journal } = installerHarness(target);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const pinned = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch: network.fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      requiredReleaseIdentity: {
        ...requiredIdentity(target),
        manifestSha256: 'ff'.repeat(32),
      },
      retryDelaysMs: [],
    });

    await expect(pinned.initialize()).rejects.toMatchObject({
      code: 'incompatible',
      path: 'release/identity',
    });

    expect(pinned.getInstallerSnapshot()).toMatchObject({
      installedRelease: target.manifest.releaseId,
      targetRelease: target.manifest.releaseId,
      readiness: 'error',
      errorCode: 'incompatible',
      canRetry: true,
    });
    await expect(pinned.repair({ scope: 'required' })).rejects.toMatchObject({
      code: 'incompatible',
      path: 'release/identity',
    });
  });

  it('fails closed when an older active release exists but the newly authorized target is unavailable', async () => {
    const releaseA = release('release-old-active', [{ path: 'one.js', value: 'object-one' }]);
    const releaseB = release('release-required-missing', [{ path: 'one.js', value: 'object-two' }]);
    const { installer, network, cache, journal } = installerHarness(releaseA);
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    const pinned = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch: network.fetch,
      cacheStorage: cache.cacheStorage,
      journal,
      requiredReleaseIdentity: requiredIdentity(releaseB),
      retryDelaysMs: [],
    });

    await expect(pinned.initialize()).rejects.toMatchObject({ code: 'network' });
    expect(pinned.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-old-active',
      targetRelease: 'release-required-missing',
      readiness: 'error',
      errorCode: 'network',
      canRetry: true,
    });
  });

  it('rejects a v5 channel that does not pin the manifest digest', async () => {
    const target = release('release-unpinned', [{ path: 'one.js', value: 'object-one' }]);
    const fetch = vi.fn(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (url.pathname === '/channels/stable-v5.json') {
        return Response.json({
          version: 1,
          releaseId: target.manifest.releaseId,
          manifestUrl: `/releases/${target.manifest.releaseId}/manifest.json`,
        });
      }
      return new Response(target.text, { status: 200 });
    }) as unknown as typeof globalThis.fetch;
    const installer = new CanonicalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test/',
      fetch,
      cacheStorage: memoryCache().cacheStorage,
      journal: new MemoryCanonicalResourceJournal(),
      retryDelaysMs: [],
    });

    await expect(installer.plan({ scope: 'all' })).rejects.toMatchObject({
      code: 'manifest',
    });
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'error',
      errorCode: 'manifest',
    });
    expect(fetch).toHaveBeenCalledTimes(1);
  });
});
