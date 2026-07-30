import { sha256 } from '@noble/hashes/sha2.js';
import { describe, expect, it, vi } from 'vitest';
import {
  MemoryInstallationJournal,
  ResourcePlanner,
  TransactionalResourceInstaller,
  parseReleaseManifest,
  type ReleaseAsset,
  type ReleaseManifestV3,
} from '../../src/lib/release-resources';

function digest(bytes: Uint8Array | string): string {
  return Array.from(sha256(typeof bytes === 'string' ? new TextEncoder().encode(bytes) : bytes), (byte) =>
    byte.toString(16).padStart(2, '0'),
  ).join('');
}

const body = new Uint8Array([1, 2, 3, 4]);
const packBody = new Uint8Array(20).fill(7);
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
          id: 'segment-001',
          offset: 0,
          bytes: packBody.byteLength,
          sha256: digest(packBody),
        },
      ],
    },
    assets: [{ ...asset, packageOffset: 16 }],
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
    if (url.pathname === `/p/${release.releaseId}/office-resources.oobpack`) {
      expect(init?.cache).not.toBe('only-if-cached');
      const segment = release.package?.segments.find((candidate) => candidate.id === url.searchParams.get('segment'));
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
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/p/'))).toHaveLength(1);
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
    expect(String(cachePut.mock.calls[0]?.[0])).toContain('segment=segment-001');
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
        { id: 'segment-001', offset: 0, bytes: 10, sha256: digest(packBody.slice(0, 10)) },
        { id: 'segment-002', offset: 10, bytes: 10, sha256: digest(packBody.slice(10, 20)) },
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
        .filter((url) => url.pathname.startsWith('/p/'))
        .map((url) => url.searchParams.get('segment')),
    ).toEqual(['segment-001', 'segment-002']);
    expect(installer.getInstallerSnapshot()).toMatchObject({
      readiness: 'ready',
      downloadedBytes: packBody.byteLength,
      verifiedBytes: packBody.byteLength,
    });
  });

  it('detects an evicted standalone segment and repairs only the missing segment', async () => {
    const release = packageManifest('v0.5.0-segmented');
    release.package = {
      ...release.package!,
      segmentBytes: 10,
      segments: [
        { id: 'segment-001', offset: 0, bytes: 10, sha256: digest(packBody.slice(0, 10)) },
        { id: 'segment-002', offset: 10, bytes: 10, sha256: digest(packBody.slice(10, 20)) },
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
    cached.entries.delete(
      'https://onlyoffice.example.test/p/v0.5.0-segmented/office-resources.oobpack?segment=segment-002',
    );

    const restarted = await create();
    expect(restarted.getInstallerSnapshot().readiness).toBe('repair-needed');
    await restarted.checkHealth();
    expect(restarted.getInstallerSnapshot()).toMatchObject({
      readiness: 'repair-needed',
      failedResources: [
        {
          path: 'office-resources.oobpack?segment=segment-002',
          code: 'storage',
        },
      ],
    });
    await restarted.repair({ scope: 'all' });
    expect(restarted.getInstallerSnapshot().readiness).toBe('ready');
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/p/'))).toHaveLength(3);
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
    expect(fetchMock.mock.calls.filter(([input]) => String(input).includes('/p/'))).toHaveLength(1);
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
