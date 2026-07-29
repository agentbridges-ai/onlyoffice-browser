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

describe('TransactionalResourceInstaller', () => {
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
