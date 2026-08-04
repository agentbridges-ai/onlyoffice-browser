import { sha256 } from '@noble/hashes/sha2.js';
import { describe, expect, it, vi } from 'vitest';
import {
  MemoryInstallationJournal,
  ReleaseRepository,
  ResourcePlanner,
  TransactionalResourceInstaller,
  computeStorageSetSha256,
  parseRequiredReleaseIdentity,
  parseReleaseManifest,
  requiredReleaseIdentitiesEqual,
  type ReleaseAsset,
  type ReleaseAssetV5,
  type ReleaseManifestV3,
  type ReleaseManifestV5,
} from '../../src/lib/release-resources';

function digest(bytes: Uint8Array | string): string {
  return Array.from(sha256(typeof bytes === 'string' ? new TextEncoder().encode(bytes) : bytes), (byte) =>
    byte.toString(16).padStart(2, '0'),
  ).join('');
}

const body = new Uint8Array([1, 2, 3, 4]);
const packBody = Uint8Array.from({ length: 20 }, (_, index) => index);
const asset: ReleaseAsset = {
  path: 'sdkjs/word/word.js',
  bytes: body.byteLength,
  mime: 'text/javascript',
  sha256: digest(body),
  profile: 'word',
  chunk: 'word-001',
};

function manifest(releaseId = 'v0.4.0-test'): ReleaseManifestV3 {
  return {
    version: 3,
    releaseId,
    packageVersion: '0.4.0',
    hostBuildId: digest('host'),
    shellRevision: digest('shell'),
    runtimeManifestSha256: digest('runtime'),
    fontManifestSha256: digest('fonts'),
    x2t: { version: '9.3.0+1', commit: 'abc123', sha256: digest('x2t') },
    profiles: {
      base: [],
      word: [asset.path],
      cell: [],
      slide: [],
      'fonts-basic': [],
      'fonts-office-compat': [],
    },
    chunks: [{ id: 'word-001', profile: 'word', bytes: asset.bytes, paths: [asset.path] }],
    assets: [asset],
  };
}

function resourceFetch(release = manifest()) {
  return vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    if (url.pathname === '/channels/stable.json') {
      return Response.json({ version: 1, releaseId: release.releaseId });
    }
    if (url.pathname === `/releases/${release.releaseId}/manifest.json`) {
      return Response.json(release);
    }
    if (url.pathname === `/r/${release.releaseId}/${asset.path}`) {
      expect(init?.cache).not.toBe('only-if-cached');
      return new Response(body);
    }
    return new Response(null, { status: 404 });
  });
}

function packageManifest(releaseId = 'v0.5.0-pack'): ReleaseManifestV3 {
  const release = manifest(releaseId);
  return {
    ...release,
    version: 4,
    packageVersion: '0.5.0',
    package: {
      format: 'onlyoffice-pack-v1',
      path: 'office-resources.oobpack',
      bytes: packBody.byteLength,
      sha256: digest(packBody),
      headerBytes: 16,
      segmentBytes: packBody.byteLength,
      segments: [
        {
          id: digest(packBody),
          offset: 0,
          bytes: packBody.byteLength,
          sha256: digest(packBody),
        },
      ],
    },
    assets: [{ ...asset, packageOffset: 16 }],
  };
}

function contentManifest(releaseId = 'v0.6.0-content'): ReleaseManifestV5 {
  const release = packageManifest(releaseId);
  const assets: ReleaseAssetV5[] = [
    {
      ...asset,
      packageOffset: 16,
      representations: {
        whole: {
          sha256: asset.sha256,
          bytes: asset.bytes,
        },
      },
    },
  ];
  const packageDescriptor = release.package!;
  return {
    ...release,
    version: 5,
    package: packageDescriptor,
    assets,
    contentProtocol: {
      version: 1,
      digest: 'sha256',
      cacheKeyFormat: 'canonical-sha256-v1',
      storageSetSha256: computeStorageSetSha256(packageDescriptor, assets),
      fastcdcPolicyId: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0',
    },
  };
}

function packageFetch(release = packageManifest()) {
  return vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input));
    if (url.pathname === '/channels/stable.json') {
      return Response.json({ version: 1, releaseId: release.releaseId });
    }
    if (url.pathname === `/releases/${release.releaseId}/manifest.json`) {
      return Response.json(release);
    }
    if (url.pathname.startsWith('/segments/sha256/')) {
      expect(init?.cache).not.toBe('only-if-cached');
      const segment = release.package?.segments.find(
        (candidate) => candidate.sha256 === url.pathname.split('/').at(-1),
      );
      if (!segment) return new Response(null, { status: 404 });
      const bytes = packBody.slice(segment.offset, segment.offset + segment.bytes);
      return new Response(bytes, {
        headers: { 'Content-Length': String(bytes.byteLength) },
      });
    }
    return new Response(null, { status: 404 });
  });
}

function memoryCacheStorage() {
  const entries = new Map<string, Response>();
  const cache = {
    put: vi.fn(async (input: RequestInfo | URL, response: Response) => {
      entries.set(String(input), response.clone());
    }),
    match: vi.fn(async (input: RequestInfo | URL) => entries.get(String(input))?.clone()),
    keys: vi.fn(async () => [...entries.keys()].map((key) => new Request(key))),
    delete: vi.fn(async (input: RequestInfo | URL) => entries.delete(String(input))),
  };
  const storage = {
    open: vi.fn(async () => cache as unknown as Cache),
    match: vi.fn(async (input: RequestInfo | URL) => entries.get(String(input))?.clone()),
    keys: vi.fn(async () => ['onlyoffice-release-staging-v0.5.0-pack']),
    delete: vi.fn(async () => true),
  } as unknown as CacheStorage;
  return { storage, entries, cache };
}

describe('Release Manifest v3', () => {
  it('strictly parses the immutable release identity used by Piwork', () => {
    const identity = {
      releaseId: 'v0.5.7-prod.1',
      manifestSha256: 'ab'.repeat(32),
      packageVersion: '0.5.7',
      hostBuildId: 'office-host-v0.5.7-r1',
    };
    expect(parseRequiredReleaseIdentity(identity)).toEqual(identity);
    expect(requiredReleaseIdentitiesEqual(identity, { ...identity })).toBe(true);
    expect(parseRequiredReleaseIdentity({ ...identity, manifestSha256: 'AB'.repeat(32) })).toBeNull();
    expect(parseRequiredReleaseIdentity({ ...identity, unexpected: true })).toBeNull();
    expect(parseRequiredReleaseIdentity({ ...identity, releaseId: '../stable' })).toBeNull();
    expect(parseRequiredReleaseIdentity(Object.assign(Object.create({}), identity))).toBeNull();
  });

  it('rejects truncated digests and unsafe paths', () => {
    expect(parseReleaseManifest(manifest())).toMatchObject({ version: 3, releaseId: 'v0.4.0-test' });
    expect(() =>
      parseReleaseManifest({
        ...manifest(),
        assets: [{ ...asset, path: '../secret' }],
      }),
    ).toThrowError('manifest');
    expect(() =>
      parseReleaseManifest({
        ...manifest(),
        assets: [{ ...asset, sha256: asset.sha256.slice(0, 16) }],
      }),
    ).toThrowError('manifest');
  });

  it('plans only changed resources between releases', () => {
    const installed = new Map([
      [
        asset.path,
        {
          releaseId: 'v0.4.0-test',
          path: asset.path,
          sha256: asset.sha256,
          bytes: asset.bytes,
          verifiedAt: 1,
        },
      ],
    ]);
    const unchanged = new ResourcePlanner(manifest(), installed).create({
      scope: 'document',
      documentType: 'word',
    });
    expect(unchanged.plan).toMatchObject({ downloadBytes: 0, reusedBytes: 4 });

    const changedManifest = manifest('v0.4.1-test');
    changedManifest.assets = [{ ...asset, sha256: digest('changed') }];
    const changed = new ResourcePlanner(changedManifest, installed).create({
      scope: 'document',
      documentType: 'word',
    });
    expect(changed.plan).toMatchObject({ downloadBytes: 4, reusedBytes: 0 });
    expect(changed.assets).toHaveLength(1);
  });
});

describe('Release Manifest v4 Office Pack', () => {
  it('requires a complete contiguous package descriptor and safe entry offsets', () => {
    expect(parseReleaseManifest(packageManifest())).toMatchObject({
      version: 4,
      package: { format: 'onlyoffice-pack-v1', bytes: 20 },
    });
    expect(() =>
      parseReleaseManifest({
        ...packageManifest(),
        package: { ...packageManifest().package!, segments: [] },
      }),
    ).toThrowError('manifest');
    expect(() =>
      parseReleaseManifest({
        ...packageManifest(),
        assets: [{ ...asset, packageOffset: 19 }],
      }),
    ).toThrowError('manifest');
  });

  it('installs every component and font through one immutable package request', async () => {
    const journal = new MemoryInstallationJournal();
    const fetchMock = packageFetch();
    const create = async () => {
      const installer = new TransactionalResourceInstaller({
        assetBaseUrl: 'https://onlyoffice.example.test',
        fetch: fetchMock as unknown as typeof fetch,
        journal,
        storageMode: 'http-cache',
        retryDelaysMs: [],
      });
      await installer.initialize();
      return installer;
    };

    const first = await create();
    const plan = await first.plan({ scope: 'document', documentType: 'word' });
    expect(plan).toMatchObject({
      downloadBytes: packBody.byteLength,
      reusedBytes: 0,
      profiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
    });
    await first.apply(plan);
    expect(first.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      downloadedBytes: packBody.byteLength,
      verifiedBytes: packBody.byteLength,
      installedProfiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
    });

    const restarted = await create();
    const reused = await restarted.plan({ scope: 'recommended' });
    expect(reused).toMatchObject({
      downloadBytes: 0,
      reusedBytes: packBody.byteLength,
    });
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/segments/sha256/'))).toHaveLength(1);
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/r/'))).toHaveLength(0);
  });

  it('keeps verified package segments in Cache Storage for the standalone origin', async () => {
    const cachePut = vi.fn();
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: packageFetch() as unknown as typeof fetch,
      cacheStorage: {
        open: vi.fn(async () => ({ put: cachePut }) as unknown as Cache),
        keys: vi.fn(async () => []),
        delete: vi.fn(async () => true),
      } as unknown as CacheStorage,
      journal: new MemoryInstallationJournal(),
      storageMode: 'cache-storage',
      retryDelaysMs: [],
    });
    await installer.initialize();
    expect(installer.getInstallerSnapshot().storageMode).toBe('cache-storage');
    const plan = await installer.plan({ scope: 'all' });
    await installer.apply(plan);
    expect(cachePut).toHaveBeenCalledOnce();
    expect(String(cachePut.mock.calls[0]?.[0])).toBe(
      `https://onlyoffice.example.test/segments/sha256/${digest(packBody)}`,
    );
  });

  it('warms only the shared HTTP cache for cross-origin Piwork integration', async () => {
    const cachePut = vi.fn();
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: packageFetch() as unknown as typeof fetch,
      cacheStorage: {
        open: vi.fn(async () => ({ put: cachePut }) as unknown as Cache),
        keys: vi.fn(async () => []),
        delete: vi.fn(async () => true),
      } as unknown as CacheStorage,
      journal: new MemoryInstallationJournal(),
      storageMode: 'http-cache',
      retryDelaysMs: [],
    });
    await installer.initialize();
    await installer.apply(await installer.plan({ scope: 'all' }));
    expect(installer.getInstallerSnapshot().storageMode).toBe('http-cache');
    expect(cachePut).not.toHaveBeenCalled();
  });

  it('downloads and verifies one immutable cacheable response per package segment', async () => {
    const release = packageManifest('v0.5.0-segmented');
    release.package = {
      ...release.package!,
      segmentBytes: 10,
      segments: [
        {
          id: digest(packBody.slice(0, 10)),
          offset: 0,
          bytes: 10,
          sha256: digest(packBody.slice(0, 10)),
        },
        {
          id: digest(packBody.slice(10, 20)),
          offset: 10,
          bytes: 10,
          sha256: digest(packBody.slice(10, 20)),
        },
      ],
    };
    const fetchMock = packageFetch(release);
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: fetchMock as unknown as typeof fetch,
      journal: new MemoryInstallationJournal(),
      storageMode: 'http-cache',
      retryDelaysMs: [],
    });
    await installer.initialize();
    await installer.apply(await installer.plan({ scope: 'all' }));
    expect(
      fetchMock.mock.calls
        .map(([input]) => new URL(String(input)))
        .filter((url) => url.pathname.startsWith('/segments/sha256/'))
        .map((url) => url.pathname.split('/').at(-1)),
    ).toEqual([digest(packBody.slice(0, 10)), digest(packBody.slice(10, 20))]);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      downloadedBytes: packBody.byteLength,
      verifiedBytes: packBody.byteLength,
    });
  });

  it('cancels and retries a package segment whose body stops yielding bytes', async () => {
    const release = packageManifest('v0.5.0-stalled-segment');
    const fallback = packageFetch(release);
    let segmentAttempts = 0;
    const fetchMock = vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input));
      if (!url.pathname.startsWith('/segments/sha256/')) return fallback(input, init);
      segmentAttempts += 1;
      if (segmentAttempts === 1) {
        const stream = new ReadableStream<Uint8Array>({
          start(controller) {
            controller.enqueue(packBody.slice(0, 1));
          },
        });
        return new Response(stream, {
          headers: { 'Content-Length': String(packBody.byteLength) },
        });
      }
      return fallback(input, init);
    });
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: fetchMock as unknown as typeof fetch,
      journal: new MemoryInstallationJournal(),
      storageMode: 'http-cache',
      timeoutMs: 20,
      retryDelaysMs: [1],
    });
    await installer.initialize();
    await installer.apply(await installer.plan({ scope: 'all' }));
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      downloadedBytes: packBody.byteLength,
      verifiedBytes: packBody.byteLength,
    });
    expect(segmentAttempts).toBe(2);
  });

  it.each(['cache-storage', 'http-cache'] as const)(
    'reuses unchanged content-addressed segments across releases in %s mode',
    async (storageMode) => {
      const nextPackBody = packBody.slice();
      nextPackBody.fill(42, 10);
      const segmentedRelease = (releaseId: string, bytes: Uint8Array) => {
        const release = packageManifest(releaseId);
        release.package = {
          ...release.package!,
          sha256: digest(bytes),
          segmentBytes: 10,
          segments: [
            {
              id: digest(bytes.slice(0, 10)),
              offset: 0,
              bytes: 10,
              sha256: digest(bytes.slice(0, 10)),
            },
            {
              id: digest(bytes.slice(10, 20)),
              offset: 10,
              bytes: 10,
              sha256: digest(bytes.slice(10, 20)),
            },
          ],
        };
        return release;
      };
      const firstRelease = segmentedRelease('v0.5.0-segmented', packBody);
      const nextRelease = segmentedRelease('v0.5.1-segmented', nextPackBody);
      let currentRelease = firstRelease;
      const segmentBodies = new Map([
        [digest(packBody.slice(0, 10)), packBody.slice(0, 10)],
        [digest(packBody.slice(10, 20)), packBody.slice(10, 20)],
        [digest(nextPackBody.slice(10, 20)), nextPackBody.slice(10, 20)],
      ]);
      const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
        const url = new URL(String(input));
        if (url.pathname === '/channels/stable.json') {
          return Response.json({ version: 1, releaseId: currentRelease.releaseId });
        }
        if (url.pathname === `/releases/${currentRelease.releaseId}/manifest.json`) {
          return Response.json(currentRelease);
        }
        const segment = segmentBodies.get(url.pathname.split('/').at(-1) || '');
        return segment
          ? new Response(segment, { headers: { 'Content-Length': String(segment.byteLength) } })
          : new Response(null, { status: 404 });
      });
      const journal = new MemoryInstallationJournal();
      const cached = memoryCacheStorage();
      const create = async () => {
        const installer = new TransactionalResourceInstaller({
          assetBaseUrl: 'https://onlyoffice.example.test',
          fetch: fetchMock as unknown as typeof fetch,
          cacheStorage: storageMode === 'cache-storage' ? cached.storage : undefined,
          journal,
          storageMode,
          retryDelaysMs: [],
        });
        await installer.initialize();
        return installer;
      };

      const first = await create();
      await first.apply(await first.plan({ scope: 'all' }));
      currentRelease = nextRelease;
      const upgraded = await create();
      const plan = await upgraded.plan({ scope: 'all' });
      expect(plan).toMatchObject({ downloadBytes: 10, reusedBytes: 10 });
      await upgraded.apply(plan);

      const sharedDigest = digest(packBody.slice(0, 10));
      const segmentRequests = fetchMock.mock.calls
        .map(([input]) => new URL(String(input)))
        .filter((url) => url.pathname.startsWith('/segments/sha256/'));
      expect(segmentRequests.filter((url) => url.pathname.endsWith(sharedDigest))).toHaveLength(1);
      expect(segmentRequests).toHaveLength(3);
      expect(upgraded.getInstallerSnapshot()).toMatchObject({
        installedRelease: nextRelease.releaseId,
        readiness: 'ready',
      });
    },
  );

  it.each(['cache-storage', 'http-cache'] as const)(
    'reuses unchanged content-addressed segments across releases in %s mode',
    async (storageMode) => {
      const nextPackBody = packBody.slice();
      nextPackBody.fill(42, 10);
      const segmentedRelease = (releaseId: string, bytes: Uint8Array) => {
        const release = packageManifest(releaseId);
        release.package = {
          ...release.package!,
          sha256: digest(bytes),
          segmentBytes: 10,
          segments: [
            {
              id: digest(bytes.slice(0, 10)),
              offset: 0,
              bytes: 10,
              sha256: digest(bytes.slice(0, 10)),
            },
            {
              id: digest(bytes.slice(10, 20)),
              offset: 10,
              bytes: 10,
              sha256: digest(bytes.slice(10, 20)),
            },
          ],
        };
        return release;
      };
      const firstRelease = segmentedRelease('v0.5.0-segmented', packBody);
      const nextRelease = segmentedRelease('v0.5.1-segmented', nextPackBody);
      let currentRelease = firstRelease;
      const segmentBodies = new Map([
        [digest(packBody.slice(0, 10)), packBody.slice(0, 10)],
        [digest(packBody.slice(10, 20)), packBody.slice(10, 20)],
        [digest(nextPackBody.slice(10, 20)), nextPackBody.slice(10, 20)],
      ]);
      const fetchMock = vi.fn(async (input: RequestInfo | URL) => {
        const url = new URL(String(input));
        if (url.pathname === '/channels/stable.json') {
          return Response.json({ version: 1, releaseId: currentRelease.releaseId });
        }
        if (url.pathname === `/releases/${currentRelease.releaseId}/manifest.json`) {
          return Response.json(currentRelease);
        }
        const segment = segmentBodies.get(url.pathname.split('/').at(-1) || '');
        return segment
          ? new Response(segment, { headers: { 'Content-Length': String(segment.byteLength) } })
          : new Response(null, { status: 404 });
      });
      const journal = new MemoryInstallationJournal();
      const cached = memoryCacheStorage();
      const create = async () => {
        const installer = new TransactionalResourceInstaller({
          assetBaseUrl: 'https://onlyoffice.example.test',
          fetch: fetchMock as unknown as typeof fetch,
          cacheStorage: storageMode === 'cache-storage' ? cached.storage : undefined,
          journal,
          storageMode,
          retryDelaysMs: [],
        });
        await installer.initialize();
        return installer;
      };

      const first = await create();
      await first.apply(await first.plan({ scope: 'all' }));
      currentRelease = nextRelease;
      const upgraded = await create();
      const plan = await upgraded.plan({ scope: 'all' });
      expect(plan).toMatchObject({ downloadBytes: 10, reusedBytes: 10 });
      await upgraded.apply(plan);

      const sharedDigest = digest(packBody.slice(0, 10));
      const segmentRequests = fetchMock.mock.calls
        .map(([input]) => new URL(String(input)))
        .filter((url) => url.pathname.startsWith('/segments/sha256/'));
      expect(segmentRequests.filter((url) => url.pathname.endsWith(sharedDigest))).toHaveLength(1);
      expect(segmentRequests).toHaveLength(3);
      expect(upgraded.getInstallerSnapshot()).toMatchObject({
        installedRelease: nextRelease.releaseId,
        readiness: 'ready',
      });
    },
  );

  it('detects an evicted standalone segment and repairs only the missing segment', async () => {
    const release = packageManifest('v0.5.0-segmented');
    release.package = {
      ...release.package!,
      segmentBytes: 10,
      segments: [
        {
          id: digest(packBody.slice(0, 10)),
          offset: 0,
          bytes: 10,
          sha256: digest(packBody.slice(0, 10)),
        },
        {
          id: digest(packBody.slice(10, 20)),
          offset: 10,
          bytes: 10,
          sha256: digest(packBody.slice(10, 20)),
        },
      ],
    };
    const fetchMock = packageFetch(release);
    const journal = new MemoryInstallationJournal();
    const cached = memoryCacheStorage();
    const create = async () => {
      const installer = new TransactionalResourceInstaller({
        assetBaseUrl: 'https://onlyoffice.example.test',
        fetch: fetchMock as unknown as typeof fetch,
        cacheStorage: cached.storage,
        journal,
        storageMode: 'cache-storage',
        retryDelaysMs: [],
      });
      await installer.initialize();
      return installer;
    };

    const first = await create();
    await first.apply(await first.plan({ scope: 'all' }));
    expect(cached.entries.size).toBe(2);
    cached.entries.delete(`https://onlyoffice.example.test/segments/sha256/${digest(packBody.slice(10, 20))}`);

    const restarted = await create();
    expect(restarted.getInstallerSnapshot().readiness).toBe('repair-needed');
    await restarted.checkHealth();
    expect(restarted.getInstallerSnapshot()).toMatchObject({
      readiness: 'repair-needed',
      failedResources: [
        {
          path: `office-resources.oobpack?segment=${digest(packBody.slice(10, 20))}`,
          code: 'storage',
        },
      ],
    });
    await restarted.repair({ scope: 'all' });
    expect(restarted.getInstallerSnapshot().readiness).toBe('ready');
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/segments/sha256/'))).toHaveLength(3);
  });

  it('fails closed when the all-in-one package digest is wrong', async () => {
    const release = packageManifest();
    const fetchMock = packageFetch(release);
    fetchMock.mockImplementation(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (url.pathname === '/channels/stable.json') {
        return Response.json({ version: 1, releaseId: release.releaseId });
      }
      if (url.pathname.includes('/manifest.json')) return Response.json(release);
      return new Response(new Uint8Array(packBody.byteLength).fill(9), {
        headers: { 'Content-Length': String(packBody.byteLength) },
      });
    });
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: fetchMock as unknown as typeof fetch,
      journal: new MemoryInstallationJournal(),
      storageMode: 'http-cache',
      retryDelaysMs: [],
    });
    await installer.initialize();
    const plan = await installer.plan({ scope: 'all' });
    await expect(installer.apply(plan)).rejects.toMatchObject({
      code: 'integrity',
      path: 'office-resources.oobpack',
    });
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'error',
      errorCode: 'integrity',
      canRetry: true,
      failedResources: [{ path: 'office-resources.oobpack', code: 'integrity' }],
    });
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/segments/sha256/'))).toHaveLength(1);
  });
});

describe('Release Manifest v5 content protocol', () => {
  it('polls only the tiny channel pointer when the cached immutable manifest identity is unchanged', async () => {
    const release = contentManifest('v0.6.0-cached-manifest');
    const manifestBytes = new TextEncoder().encode(JSON.stringify(release));
    const manifestSha256 = digest(manifestBytes);
    const requests: string[] = [];
    const repository = new ReleaseRepository(
      'https://onlyoffice.example.test/',
      vi.fn(async (input: RequestInfo | URL) => {
        const url = new URL(String(input));
        requests.push(url.pathname);
        if (url.pathname === '/channels/stable-v5.json') {
          return Response.json({
            version: 1,
            releaseId: release.releaseId,
            manifestUrl: `/releases/${release.releaseId}/manifest.json`,
            manifestSha256,
          });
        }
        return new Response(manifestBytes);
      }) as unknown as typeof fetch,
    );

    const first = await repository.currentV5();
    const second = await repository.currentV5(first);

    expect(second).toEqual(first);
    expect(requests).toEqual([
      '/channels/stable-v5.json',
      `/releases/${release.releaseId}/manifest.json`,
      '/channels/stable-v5.json',
    ]);
  });

  it('pins the SHA-256 of the exact manifest response bytes', async () => {
    const release = contentManifest('v0.6.0-raw-manifest');
    const encoded = new TextEncoder().encode(JSON.stringify(release));
    const raw = new Uint8Array(encoded.byteLength + 3);
    raw.set([0xef, 0xbb, 0xbf]);
    raw.set(encoded, 3);
    const repository = new ReleaseRepository(
      'https://onlyoffice.example.test/',
      vi.fn(async () => new Response(raw)) as unknown as typeof fetch,
    );

    await expect(repository.releaseV5(release.releaseId, digest(raw))).resolves.toMatchObject({
      manifest: { releaseId: release.releaseId },
      manifestSha256: digest(raw),
    });
    await expect(repository.releaseV5(release.releaseId, digest(encoded))).rejects.toMatchObject({
      code: 'manifest',
    });
    await expect(repository.releaseV5('different-release-id', digest(raw))).rejects.toMatchObject({
      code: 'incompatible',
    });
  });

  it('accepts a deterministic whole-file CAS representation', () => {
    const release = contentManifest();
    expect(parseReleaseManifest(release)).toMatchObject({
      version: 5,
      releaseId: 'v0.6.0-content',
      contentProtocol: {
        cacheKeyFormat: 'canonical-sha256-v1',
        storageSetSha256: computeStorageSetSha256(release.package, release.assets),
      },
    });
  });

  it('rejects a storage-set digest or whole representation that does not match the asset', () => {
    const release = contentManifest();
    expect(() =>
      parseReleaseManifest({
        ...release,
        contentProtocol: { ...release.contentProtocol, storageSetSha256: digest('tampered') },
      }),
    ).toThrowError('manifest');
    expect(() =>
      parseReleaseManifest({
        ...release,
        assets: [
          {
            ...release.assets[0],
            representations: {
              whole: { ...release.assets[0].representations.whole, bytes: asset.bytes + 1 },
            },
          },
        ],
      }),
    ).toThrowError('manifest');
  });

  it('rejects unsafe encoded paths, non-contiguous FastCDC chunks, and digest-size conflicts', () => {
    const release = contentManifest();
    expect(() =>
      parseReleaseManifest({
        ...release,
        assets: [{ ...release.assets[0], path: '%2e%2e/secret' }],
      }),
    ).toThrowError('manifest');

    const withFastCdc = (chunks: Array<{ offset: number; bytes: number; sha256: string }>) => ({
      ...release.assets[0],
      representations: {
        ...release.assets[0].representations,
        fastcdc: {
          algorithm: 'fastcdc-v2020',
          minBytes: 65_536,
          averageBytes: 262_144,
          maxBytes: 1_048_576,
          normalization: 1,
          seed: 0,
          chunks,
        },
      },
    });
    expect(() =>
      parseReleaseManifest({
        ...release,
        assets: [
          withFastCdc([
            { offset: 0, bytes: 2, sha256: digest('first') },
            { offset: 3, bytes: 2, sha256: digest('second') },
          ]),
        ],
      }),
    ).toThrowError('manifest');
    expect(() =>
      parseReleaseManifest({
        ...release,
        assets: [
          withFastCdc([
            {
              offset: 0,
              bytes: asset.bytes,
              sha256: release.package.segments[0].sha256,
            },
          ]),
        ],
      }),
    ).toThrowError('manifest');
  });
});

describe('TransactionalResourceInstaller', () => {
  it('accepts same-release snapshots without rebroadcasting and ignores stale releases', async () => {
    const broadcast = new EventTarget() as BroadcastChannel;
    broadcast.postMessage = vi.fn();
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: resourceFetch() as unknown as typeof fetch,
      journal: new MemoryInstallationJournal(),
      storageMode: 'http-cache',
      broadcast,
    });
    await installer.initialize();
    const listener = vi.fn();
    installer.subscribeInstaller(listener);
    vi.mocked(broadcast.postMessage).mockClear();

    const current = installer.getInstallerSnapshot();
    broadcast.dispatchEvent(
      new MessageEvent('message', {
        data: { type: 'resource-snapshot', snapshot: { ...current, phase: 'paused', readiness: 'paused' } },
      }),
    );
    expect(listener).toHaveBeenCalledOnce();
    expect(installer.getInstallerSnapshot()).toMatchObject({ phase: 'paused', readiness: 'paused' });
    expect(broadcast.postMessage).not.toHaveBeenCalled();

    listener.mockClear();
    broadcast.dispatchEvent(
      new MessageEvent('message', {
        data: {
          type: 'resource-snapshot',
          snapshot: { ...current, targetRelease: 'v0.3.0-stale', availableRelease: 'v0.3.0-stale' },
        },
      }),
    );
    expect(listener).not.toHaveBeenCalled();
    expect(installer.getInstallerSnapshot()).toMatchObject({ targetRelease: current.targetRelease });
    expect(broadcast.postMessage).not.toHaveBeenCalled();
  });

  it('commits each verified resource and reuses it after a restarted installer', async () => {
    const journal = new MemoryInstallationJournal();
    const fetchMock = resourceFetch();
    const create = async () => {
      const installer = new TransactionalResourceInstaller({
        assetBaseUrl: 'https://onlyoffice.example.test',
        fetch: fetchMock as unknown as typeof fetch,
        journal,
        storageMode: 'http-cache',
        retryDelaysMs: [],
      });
      await installer.initialize();
      return installer;
    };
    const first = await create();
    const firstPlan = await first.plan({ scope: 'document', documentType: 'word' });
    expect(firstPlan.downloadBytes).toBe(4);
    await first.apply(firstPlan);
    expect(first.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'v0.4.0-test',
      readiness: 'ready',
      downloadedBytes: 4,
      verifiedBytes: 4,
    });

    const restarted = await create();
    const resumedPlan = await restarted.plan({ scope: 'document', documentType: 'word' });
    expect(resumedPlan).toMatchObject({ downloadBytes: 0, reusedBytes: 4 });
    const assetRequests = fetchMock.mock.calls.filter(([input]) => String(input).includes('/r/'));
    expect(assetRequests).toHaveLength(1);
  });

  it('reports integrity failures with a retryable structured error', async () => {
    const fetchMock = resourceFetch();
    fetchMock.mockImplementation(async (input: RequestInfo | URL) => {
      const url = new URL(String(input));
      if (url.pathname === '/channels/stable.json') {
        return Response.json({ version: 1, releaseId: manifest().releaseId });
      }
      if (url.pathname.includes('/manifest.json')) return Response.json(manifest());
      return new Response(new Uint8Array([9, 9, 9, 9]));
    });
    const installer = new TransactionalResourceInstaller({
      assetBaseUrl: 'https://onlyoffice.example.test',
      fetch: fetchMock as unknown as typeof fetch,
      journal: new MemoryInstallationJournal(),
      storageMode: 'http-cache',
      retryDelaysMs: [],
    });
    await installer.initialize();
    const plan = await installer.plan({ scope: 'document', documentType: 'word' });

    await expect(installer.apply(plan)).rejects.toMatchObject({ code: 'integrity', path: asset.path });
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'error',
      errorCode: 'integrity',
      canRetry: true,
      downloadedBytes: 0,
      verifiedBytes: 0,
      failedResources: [{ path: asset.path, code: 'integrity' }],
    });
  });
});
