import { beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createOfficeRuntimeResourceManager,
  resolveOfficeResourceInstallerRuntimeMode,
  type OfficeRuntimeResourceSnapshot,
} from '../../src/lib/runtime-resources';
import type { OfficeRuntimeResourceInstaller, ResourceInstallerSnapshot } from '../../src/lib/release-resources';

const revision = '9f64a747e1b97f13';

function response(body: BodyInit | null, init: ResponseInit = {}): Response {
  return new Response(body, { status: 200, ...init });
}

function resourceFetch() {
  return vi.fn(async (input: RequestInfo | URL, init?: RequestInit) => {
    const url = new URL(String(input), window.location.origin);
    if (url.pathname === '/onlyoffice-runtime-assets.json' && init?.method === 'HEAD') {
      return response(null, { headers: { 'X-OnlyOffice-Asset-Version': 'resource-v1' } });
    }
    if (url.pathname === '/onlyoffice-runtime-assets.json') {
      return response(
        JSON.stringify({
          version: 2,
          generatedAt: '2026-07-28T00:00:00.000Z',
          assets: [{ path: 'sdkjs/word/word.js', bytes: 4, pack: 'word', revision }],
        }),
      );
    }
    if (url.pathname === '/onlyoffice-browser-font-assets.json') {
      return response(
        JSON.stringify({
          defaultFonts: ['fonts/dengxian.ttf'],
          builtInFonts: [],
          fontFamilies: [
            { name: 'DengXian', paths: ['fonts/dengxian.ttf'] },
            { name: 'Microsoft YaHei', paths: ['fonts/yahei.ttf'] },
          ],
          assets: [
            {
              path: 'fonts/dengxian.ttf',
              bytes: 4,
              revision,
              families: ['DengXian'],
            },
            {
              path: 'fonts/yahei.ttf',
              bytes: 4,
              revision,
              families: ['Microsoft YaHei'],
            },
          ],
        }),
      );
    }
    return response(new Uint8Array([1, 2, 3, 4]));
  });
}

describe('OfficeRuntimeResourceManager', () => {
  beforeEach(() => {
    window.localStorage.clear();
  });

  it('publishes verified optional fonts and keeps the required default non-removable', async () => {
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: resourceFetch() as unknown as typeof fetch,
      cacheStorage: undefined,
    });
    const snapshots: OfficeRuntimeResourceSnapshot[] = [];
    manager.subscribe((snapshot) => snapshots.push(snapshot));

    expect(manager.getSnapshot().fonts).toEqual([
      expect.objectContaining({ name: 'DengXian', removable: false }),
      expect.objectContaining({ name: 'Microsoft YaHei', downloaded: false, removable: true }),
    ]);
    expect(manager.getSnapshot()).toMatchObject({
      packageVersion: '0.5.13',
      assetVersion: 'resource-v1',
      readiness: 'needs-download',
      packs: expect.arrayContaining([expect.objectContaining({ id: 'word', ready: false })]),
    });
    expect(manager.getSnapshot()).toBe(manager.getSnapshot());

    await manager.downloadFontFamily('microsoft yahei');

    expect(manager.getVerifiedFontPaths()).toContain('fonts/yahei.ttf');
    expect(manager.getSnapshot()).toMatchObject({
      operation: null,
      error: null,
      fonts: [
        expect.objectContaining({ name: 'DengXian', removable: false }),
        expect.objectContaining({ name: 'Microsoft YaHei', downloaded: true, removable: true }),
      ],
    });
    expect(snapshots.some((snapshot) => snapshot.operation === 'download-font')).toBe(true);
    expect(snapshots.at(-1)).toBe(manager.getSnapshot());

    await manager.uninstallFontFamily('microsoft yahei');
    expect(manager.getVerifiedFontPaths()).not.toContain('fonts/yahei.ttf');
  });

  it('prepares only the packs needed by the active document', async () => {
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });

    await manager.prepareForDocumentType('word');

    expect(manager.getSnapshot().packs.find((pack) => pack.id === 'word')).toMatchObject({ ready: true });
    expect(manager.getSnapshot().operation).toBeNull();
  });

  it('coalesces identical operations and serializes different mutations', async () => {
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });

    const first = manager.downloadFontFamily('microsoft yahei');
    const duplicate = manager.downloadFontFamily('microsoft yahei');
    const remove = manager.uninstallFontFamily('microsoft yahei');

    expect(duplicate).toBe(first);
    await first;
    await remove;
    expect(manager.getSnapshot().fonts.find((font) => font.name === 'Microsoft YaHei')).toMatchObject({
      downloaded: false,
    });
  });

  it('loads and verifies resources from an embedding-provided canonical asset origin', async () => {
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
      assetBaseUrl: 'https://onlyoffice.example.test/runtime/',
    });

    await manager.downloadFontFamily('microsoft yahei');

    expect(
      fetchMock.mock.calls.some(
        ([input]) => String(input) === 'https://onlyoffice.example.test/onlyoffice-runtime-assets.json',
      ),
    ).toBe(true);
    expect(fetchMock.mock.calls.find(([input]) => String(input).includes('/fonts/yahei.ttf'))?.[1]).toMatchObject({
      mode: 'cors',
      credentials: 'omit',
    });
  });

  it('surfaces integrity failures without losing the last usable catalog', async () => {
    const fetchMock = resourceFetch();
    fetchMock.mockImplementation(async (input: RequestInfo | URL, init?: RequestInit) => {
      const url = new URL(String(input), window.location.origin);
      if (url.pathname.endsWith('/word.js') && init?.cache !== undefined) {
        return response(new Uint8Array([9, 9, 9, 9]));
      }
      return resourceFetch()(input, init);
    });
    const manager = await createOfficeRuntimeResourceManager({
      storage: window.localStorage,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });

    const result = await manager.loadAll();

    expect(result.phase).toBe('error');
    expect(manager.getSnapshot().progress.failedFiles).toBe(1);
    expect(manager.getSnapshot().fonts).toHaveLength(2);
  });

  it('keeps font mutations on the canonical all-in-one installer without fetching resource bytes in the parent', async () => {
    let installerListener: (snapshot: ResourceInstallerSnapshot) => void = () => undefined;
    const installerSnapshot: ResourceInstallerSnapshot = {
      installedRelease: null,
      targetRelease: 'release-a',
      availableRelease: 'release-a',
      availablePackageVersion: '0.5.7',
      readiness: 'needs-download',
      phase: 'idle',
      storageMode: 'cache-storage',
      currentChunk: null,
      currentChunkIndex: 0,
      currentChunkCount: 0,
      downloadedBytes: 0,
      downloadBytes: 400,
      verifiedBytes: 0,
      verifyBytes: 400,
      bytesPerSecond: 0,
      failedResources: [],
      canPause: false,
      canResume: false,
      canRetry: false,
      errorCode: null,
      installedProfiles: [],
    };
    const releaseInstaller: OfficeRuntimeResourceInstaller = {
      plan: vi.fn(async () => ({
        planId: 'fonts-plan',
        releaseId: 'release-a',
        scope: 'fonts' as const,
        profiles: ['fonts-basic', 'fonts-office-compat'],
        totalBytes: 400,
        downloadBytes: 400,
        reusedBytes: 0,
      })),
      apply: vi.fn(async () => undefined),
      checkForUpdates: vi.fn(async () => undefined),
      checkHealth: vi.fn(async () => undefined),
      repair: vi.fn(async () => undefined),
      pause: vi.fn(),
      resume: vi.fn(async () => undefined),
      cancel: vi.fn(),
      getInstallerSnapshot: () => installerSnapshot,
      subscribeInstaller: (listener) => {
        installerListener = listener;
        return () => undefined;
      },
      getInstalledPaths: () => [],
    };
    const fetchMock = resourceFetch();
    const manager = await createOfficeRuntimeResourceManager({
      releaseInstaller,
      fetch: fetchMock as unknown as typeof fetch,
      cacheStorage: undefined,
    });
    expect(fetchMock, 'canonical v5 must not load legacy runtime or font manifests').not.toHaveBeenCalled();

    expect(manager.remainingBytes()).toBe(400);
    installerSnapshot.installedRelease = 'release-a';
    installerSnapshot.readiness = 'ready';
    installerSnapshot.installedProfiles = ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'];
    installerSnapshot.downloadBytes = 0;
    installerSnapshot.verifyBytes = 0;
    installerListener(installerSnapshot);

    await manager.prepareForDocumentType('word');
    expect(
      releaseInstaller.plan,
      'ready all-in-one resources must not be planned or fetched again',
    ).not.toHaveBeenCalled();

    await manager.downloadFontFamily('microsoft yahei');
    await manager.uninstallFontFamily('microsoft yahei');

    expect(releaseInstaller.plan).toHaveBeenCalledWith({ scope: 'fonts' });
    expect(releaseInstaller.apply).toHaveBeenCalledOnce();
    expect(
      fetchMock.mock.calls.some(
        ([input]) => new URL(String(input), window.location.origin).pathname === '/fonts/yahei.ttf',
      ),
    ).toBe(false);

    installerSnapshot.installedRelease = null;
    installerSnapshot.readiness = 'error';
    installerSnapshot.errorCode = 'network';
    installerSnapshot.failedResources = [{ path: 'fonts/missing.ttf', code: 'network', attempts: 3 }];
    installerListener(installerSnapshot);

    expect(manager.getSnapshot()).toMatchObject({
      readiness: 'error',
      error: { code: 'network', path: 'fonts/missing.ttf' },
      installedRelease: null,
    });
  });

  it('automatically completes a candidate with the retained release still usable when the candidate fails', async () => {
    let installerListener: (snapshot: ResourceInstallerSnapshot) => void = () => undefined;
    const installerSnapshot: ResourceInstallerSnapshot = {
      installedRelease: 'release-a',
      targetRelease: 'release-b',
      availableRelease: 'release-b',
      availablePackageVersion: '0.5.8',
      readiness: 'update-available',
      phase: 'idle',
      storageMode: 'cache-storage',
      currentChunk: null,
      currentChunkIndex: 0,
      currentChunkCount: 0,
      downloadedBytes: 0,
      downloadBytes: 40,
      verifiedBytes: 0,
      verifyBytes: 40,
      bytesPerSecond: 0,
      failedResources: [],
      canPause: false,
      canResume: false,
      canRetry: false,
      errorCode: null,
      installedProfiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
    };
    const releaseInstaller: OfficeRuntimeResourceInstaller = {
      plan: vi.fn(async () => ({
        planId: 'candidate-plan',
        releaseId: 'release-b',
        scope: 'all' as const,
        profiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
        totalBytes: 40,
        downloadBytes: 40,
        reusedBytes: 60,
      })),
      apply: vi.fn(async () => {
        // This is the installer contract after a failed candidate prewarm:
        // A stays active and B remains available for a later retry.
        installerSnapshot.failedResources = [{ path: 'objects/sha256/candidate', code: 'network', attempts: 1 }];
        installerSnapshot.errorCode = 'network';
        installerSnapshot.canRetry = true;
        installerListener(installerSnapshot);
        throw Object.assign(new Error('candidate download failed'), { code: 'network' as const });
      }),
      checkForUpdates: vi.fn(async () => undefined),
      checkHealth: vi.fn(async () => undefined),
      repair: vi.fn(async () => undefined),
      pause: vi.fn(),
      resume: vi.fn(async () => undefined),
      cancel: vi.fn(),
      getInstallerSnapshot: () => installerSnapshot,
      subscribeInstaller: (listener) => {
        installerListener = listener;
        return () => undefined;
      },
      getInstalledPaths: () => [],
    };
    const manager = await createOfficeRuntimeResourceManager({ releaseInstaller });

    await expect(manager.maintain()).rejects.toMatchObject({ code: 'network' });

    expect(releaseInstaller.checkForUpdates).toHaveBeenCalledOnce();
    expect(releaseInstaller.plan).toHaveBeenCalledWith({ scope: 'all' });
    expect(releaseInstaller.apply).toHaveBeenCalledOnce();
    expect(manager.getSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      targetRelease: 'release-b',
      readiness: 'update-available',
      error: { code: 'network', path: 'objects/sha256/candidate' },
    });
  });

  it('automatically repairs an installed release instead of redownloading its complete package', async () => {
    let installerListener: (snapshot: ResourceInstallerSnapshot) => void = () => undefined;
    const installerSnapshot: ResourceInstallerSnapshot = {
      installedRelease: 'release-a',
      targetRelease: 'release-a',
      availableRelease: 'release-a',
      availablePackageVersion: '0.5.7',
      readiness: 'repair-needed',
      phase: 'idle',
      storageMode: 'cache-storage',
      currentChunk: null,
      currentChunkIndex: 0,
      currentChunkCount: 0,
      downloadedBytes: 0,
      downloadBytes: 100,
      verifiedBytes: 0,
      verifyBytes: 100,
      bytesPerSecond: 0,
      failedResources: [{ path: 'objects/sha256/bad', code: 'integrity', attempts: 1 }],
      canPause: false,
      canResume: false,
      canRetry: true,
      errorCode: 'integrity',
      installedProfiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
    };
    const releaseInstaller: OfficeRuntimeResourceInstaller = {
      plan: vi.fn(),
      apply: vi.fn(),
      checkForUpdates: vi.fn(async () => undefined),
      checkHealth: vi.fn(async () => undefined),
      repair: vi.fn(async () => {
        installerSnapshot.readiness = 'ready';
        installerSnapshot.failedResources = [];
        installerSnapshot.errorCode = null;
        installerSnapshot.canRetry = false;
        installerListener(installerSnapshot);
      }),
      pause: vi.fn(),
      resume: vi.fn(async () => undefined),
      cancel: vi.fn(),
      getInstallerSnapshot: () => installerSnapshot,
      subscribeInstaller: (listener) => {
        installerListener = listener;
        return () => undefined;
      },
      getInstalledPaths: () => [],
    };
    const manager = await createOfficeRuntimeResourceManager({ releaseInstaller });

    await manager.maintain();

    expect(releaseInstaller.repair).toHaveBeenCalledWith({ scope: 'installed' });
    expect(releaseInstaller.plan).not.toHaveBeenCalled();
    expect(releaseInstaller.apply).not.toHaveBeenCalled();
    expect(manager.getSnapshot()).toMatchObject({ installedRelease: 'release-a', readiness: 'ready', error: null });
  });
});

describe('Office resource installer runtime selection', () => {
  it('uses one canonical store in production and requires explicit localhost test mode', () => {
    expect(
      resolveOfficeResourceInstallerRuntimeMode({
        pageOrigin: 'https://onlyoffice.getpi.work',
      }),
    ).toBe('canonical');
    expect(
      resolveOfficeResourceInstallerRuntimeMode({
        pageOrigin: 'https://piwork.getpi.work',
      }),
    ).toBe('remote');
    expect(
      resolveOfficeResourceInstallerRuntimeMode({
        pageOrigin: 'https://piwork.getpi.work.evil.example',
      }),
    ).toBe('legacy');

    expect(
      resolveOfficeResourceInstallerRuntimeMode({
        pageOrigin: 'http://piwork.localhost:8787',
        canonicalOrigin: 'http://onlyoffice.localhost:8787',
      }),
    ).toBe('legacy');
    expect(
      resolveOfficeResourceInstallerRuntimeMode({
        pageOrigin: 'http://piwork.localhost:8787',
        canonicalOrigin: 'http://onlyoffice.localhost:8787',
        allowLocalTestMode: true,
      }),
    ).toBe('remote');
    expect(
      resolveOfficeResourceInstallerRuntimeMode({
        pageOrigin: 'http://onlyoffice.localhost:8787',
        canonicalOrigin: 'http://onlyoffice.localhost:8787',
        allowLocalTestMode: true,
      }),
    ).toBe('canonical');
  });
});
