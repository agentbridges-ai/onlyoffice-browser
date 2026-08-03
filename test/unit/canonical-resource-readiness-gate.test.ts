import { describe, expect, it, vi } from 'vitest';
import { OFFICE_EDITOR_ORIGIN_SLOTS } from '../../src/lib/office-origin-pool';
import {
  CanonicalResourceReadinessGate,
  type CanonicalReleaseReadinessProbe,
} from '../../src/lib/canonical-resource-readiness-gate';
import { EDITOR_SHELL_MAX_TOTAL_BYTES } from '../../src/lib/editor-shell-cache';
import {
  ResourceInstallerError,
  type OfficeRuntimeResourceInstaller,
  type ResourceErrorCode,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
} from '../../src/lib/release-resources';

const readySnapshot = (releaseId = 'release-a'): ResourceInstallerSnapshot => ({
  installedRelease: releaseId,
  targetRelease: releaseId,
  availableRelease: releaseId,
  availablePackageVersion: '0.5.7',
  readiness: 'ready',
  phase: 'idle',
  storageMode: 'cache-storage',
  currentChunk: null,
  currentChunkIndex: 0,
  currentChunkCount: 0,
  downloadedBytes: 100,
  downloadBytes: 100,
  verifiedBytes: 100,
  verifyBytes: 100,
  bytesPerSecond: 0,
  failedResources: [],
  canPause: false,
  canResume: false,
  canRetry: false,
  errorCode: null,
  installedProfiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
});

const plan: ResourcePlan = {
  planId: 'plan-a',
  releaseId: 'release-a',
  scope: 'all',
  profiles: ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'],
  totalBytes: 100,
  downloadBytes: 100,
  reusedBytes: 0,
};

function primeResults(releaseId: string) {
  return OFFICE_EDITOR_ORIGIN_SLOTS.map((slot) => ({
    origin: `https://${slot}.getpi.work`,
    releaseId,
    sessionId: `prime-${slot}`,
    brokerReady: true as const,
    occupied: false,
    serviceWorkerVersion: 'sw-test',
    cachedPaths: ['office-host.html', 'editor-shell-prime.html', 'assets/officeHost.js'],
    cachedBytes: 100,
    storageMode: 'cache' as const,
  }));
}

class FakeInstaller implements OfficeRuntimeResourceInstaller {
  snapshot = readySnapshot();
  readonly listeners = new Set<(snapshot: ResourceInstallerSnapshot) => void>();

  initialize = vi.fn(async () => undefined);
  plan = vi.fn(async () => plan);
  apply = vi.fn(async () => undefined);
  checkForUpdates = vi.fn(async () => undefined);
  checkHealth = vi.fn(async () => undefined);
  repair = vi.fn(async () => undefined);
  pause = vi.fn();
  resume = vi.fn(async () => undefined);
  cancel = vi.fn();
  rollbackActivation = vi.fn(async (releaseId: string, failure: { code: ResourceErrorCode; path: string }) => {
    this.publish({
      ...readySnapshot('release-a'),
      targetRelease: releaseId,
      availableRelease: releaseId,
      readiness: 'update-available',
      failedResources: [
        {
          path: failure.path,
          code: failure.code,
          attempts: 1,
        },
      ],
      canRetry: true,
      errorCode: failure.code,
    });
  });

  getInstallerSnapshot(): ResourceInstallerSnapshot {
    return this.snapshot;
  }

  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  getInstalledPaths(): string[] {
    return ['fonts/Aptos.ttf'];
  }

  publish(snapshot: ResourceInstallerSnapshot): void {
    this.snapshot = snapshot;
    for (const listener of this.listeners) listener(snapshot);
  }
}

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((nextResolve) => {
    resolve = nextResolve;
  });
  return { promise, resolve };
}

describe('CanonicalResourceReadinessGate', () => {
  it('does not publish ready until all 12 editor origins have primed and probed the exact release', async () => {
    const installer = new FakeInstaller();
    const pending = deferred<ReturnType<typeof primeResults>>();
    const probe: CanonicalReleaseReadinessProbe = vi.fn(() => pending.promise);
    const gate = new CanonicalResourceReadinessGate({ installer, probeRelease: probe });

    const initialization = gate.initialize();
    await vi.waitFor(() => {
      expect(gate.getInstallerSnapshot()).toMatchObject({
        readiness: 'updating',
        phase: 'activating',
      });
    });
    expect(probe).toHaveBeenCalledWith('release-a', expect.any(Function));

    pending.resolve(primeResults('release-a'));
    await initialization;
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      readiness: 'ready',
      phase: 'idle',
      errorCode: null,
    });
  });

  it('uses the agreed three-origin acceptance set in the local Cloudflare matrix', async () => {
    const installer = new FakeInstaller();
    const localResults = primeResults('release-a')
      .slice(0, 3)
      .map((result) => ({
        ...result,
        origin: result.origin.replace('https://', 'http://host-').replace('.getpi.work', '.office.localhost:8787'),
      }));
    const probe: CanonicalReleaseReadinessProbe = vi.fn(async () => localResults);
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: probe,
      canonicalOrigin: 'http://onlyoffice.localhost:8787',
      localTestMode: true,
    });

    await gate.initialize();
    expect(gate.getInstallerSnapshot()).toMatchObject({ readiness: 'ready', installedRelease: 'release-a' });
    expect(localResults.map(({ origin }) => origin)).toEqual([
      'http://host-aries.office.localhost:8787',
      'http://host-taurus.office.localhost:8787',
      'http://host-gemini.office.localhost:8787',
    ]);
  });

  it('fails closed with a structured repair state when any origin or Broker probe is missing', async () => {
    const installer = new FakeInstaller();
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: async (releaseId) => primeResults(releaseId).slice(0, 11),
    });

    await expect(gate.initialize()).rejects.toMatchObject({
      code: 'storage',
      path: 'editor-origins/incomplete',
    });
    expect(gate.getInstallerSnapshot()).toMatchObject({
      readiness: 'repair-needed',
      errorCode: 'storage',
      failedResources: [
        {
          path: 'editor-origins/incomplete',
          code: 'storage',
          attempts: 1,
        },
      ],
      canRetry: true,
    });
  });

  it('rejects 12 successful-looking results unless they are the fixed constellation origins', async () => {
    const installer = new FakeInstaller();
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: async (releaseId) => {
        const results = primeResults(releaseId);
        results[0] = { ...results[0], origin: 'https://evil.example' };
        return results;
      },
    });
    await expect(gate.initialize()).rejects.toMatchObject({
      code: 'storage',
      path: 'editor-origins/incomplete',
    });
    expect(gate.getInstallerSnapshot().readiness).toBe('repair-needed');
  });

  it('rejects an origin result whose shell proof exceeds the bounded manifest total', async () => {
    const installer = new FakeInstaller();
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: async (releaseId) => {
        const results = primeResults(releaseId);
        results[0] = {
          ...results[0],
          cachedBytes: EDITOR_SHELL_MAX_TOTAL_BYTES + 1,
        };
        return results;
      },
    });
    await expect(gate.initialize()).rejects.toMatchObject({
      code: 'storage',
      path: 'editor-origins/https://aries.getpi.work',
    });
    expect(gate.getInstallerSnapshot().readiness).toBe('repair-needed');
  });

  it('allows older pinned editors to occupy origins when every shell is ready and one free origin completes a read', async () => {
    const installer = new FakeInstaller();
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: async (releaseId) =>
        primeResults(releaseId).map((result, index) =>
          index === 0
            ? result
            : {
                ...result,
                brokerReady: false,
                occupied: true,
              },
        ),
    });

    await expect(gate.initialize()).resolves.toBeUndefined();
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      readiness: 'ready',
    });
  });

  it('fails closed when all origins are occupied and no routed Broker read can be proven', async () => {
    const installer = new FakeInstaller();
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: async (releaseId) =>
        primeResults(releaseId).map((result) => ({
          ...result,
          brokerReady: false,
          occupied: true,
        })),
    });

    await expect(gate.initialize()).rejects.toMatchObject({
      code: 'storage',
      path: 'editor-origins/no-live-broker-probe',
    });
  });

  it('keeps routine health checks fast while explicit update checks revalidate all origins', async () => {
    const installer = new FakeInstaller();
    const probe = vi.fn(async (releaseId: string) => primeResults(releaseId));
    const gate = new CanonicalResourceReadinessGate({ installer, probeRelease: probe });
    await gate.initialize();
    await gate.checkHealth();
    await gate.checkForUpdates();
    expect(probe).toHaveBeenCalledTimes(2);
    expect(probe).toHaveBeenNthCalledWith(1, 'release-a', expect.any(Function));
    expect(probe).toHaveBeenNthCalledWith(2, 'release-a', expect.any(Function));

    installer.publish(readySnapshot('release-b'));
    await gate.checkHealth();
    expect(probe).toHaveBeenNthCalledWith(3, 'release-b', expect.any(Function));
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-b',
      readiness: 'ready',
    });
  });

  it('turns a same-release shell or Service Worker probe failure into repair-needed', async () => {
    const installer = new FakeInstaller();
    let healthy = true;
    const probe: CanonicalReleaseReadinessProbe = vi.fn(async (releaseId) => {
      if (!healthy) {
        throw new ResourceInstallerError('storage', 'editor-origins/https://aries.getpi.work');
      }
      return primeResults(releaseId);
    });
    const gate = new CanonicalResourceReadinessGate({ installer, probeRelease: probe });
    await gate.initialize();
    expect(gate.getInstallerSnapshot().readiness).toBe('ready');

    healthy = false;
    await expect(gate.checkForUpdates()).rejects.toMatchObject({
      code: 'storage',
      path: 'editor-origins/https://aries.getpi.work',
    });
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      readiness: 'repair-needed',
      errorCode: 'storage',
      failedResources: [
        {
          path: 'editor-origins/https://aries.getpi.work',
          code: 'storage',
        },
      ],
      canRetry: true,
    });
  });

  it('publishes a readiness heartbeat after every completed origin probe', async () => {
    const installer = new FakeInstaller();
    const probe: CanonicalReleaseReadinessProbe = async (releaseId, onProgress) => {
      const results = primeResults(releaseId);
      for (const result of results) onProgress?.(result);
      return results;
    };
    const gate = new CanonicalResourceReadinessGate({ installer, probeRelease: probe });
    const snapshots: ResourceInstallerSnapshot[] = [];
    gate.subscribeInstaller((snapshot) => snapshots.push(snapshot));

    await gate.initialize();

    expect(
      snapshots.filter((snapshot) => snapshot.readiness === 'updating' && snapshot.phase === 'activating'),
    ).toHaveLength(13);
    expect(gate.getInstallerSnapshot().readiness).toBe('ready');
  });

  it('rolls a failed B readiness probe back to validated A and resumes B without invalidating A', async () => {
    const installer = new FakeInstaller();
    let failB = true;
    const probe = vi.fn(async (releaseId: string) =>
      releaseId === 'release-b' && failB ? primeResults(releaseId).slice(0, 11) : primeResults(releaseId),
    );
    const gate = new CanonicalResourceReadinessGate({ installer, probeRelease: probe });
    await gate.initialize();
    installer.apply.mockImplementationOnce(async () => {
      installer.publish(readySnapshot('release-b'));
    });

    await expect(
      gate.apply({
        ...plan,
        planId: 'plan-b',
        releaseId: 'release-b',
      }),
    ).rejects.toMatchObject({
      code: 'storage',
      path: 'editor-origins/incomplete',
    });

    expect(installer.rollbackActivation).toHaveBeenCalledWith('release-b', {
      code: 'storage',
      path: 'editor-origins/incomplete',
    });
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-a',
      targetRelease: 'release-b',
      availableRelease: 'release-b',
      readiness: 'update-available',
      errorCode: 'storage',
      canRetry: true,
      failedResources: [
        {
          path: 'editor-origins/incomplete',
          code: 'storage',
        },
      ],
    });

    failB = false;
    installer.resume.mockImplementationOnce(async () => {
      installer.publish(readySnapshot('release-b'));
    });
    await gate.resume();
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: 'release-b',
      readiness: 'ready',
      errorCode: null,
    });
  });

  it('rolls a failed first activation back to no active release and never exposes ready', async () => {
    const installer = new FakeInstaller();
    installer.snapshot = {
      ...readySnapshot(),
      installedRelease: null,
      installedProfiles: [],
      readiness: 'needs-download',
    };
    installer.apply.mockImplementationOnce(async () => {
      installer.publish(readySnapshot('release-b'));
    });
    installer.rollbackActivation.mockImplementationOnce(async (releaseId, failure) => {
      installer.publish({
        ...readySnapshot(releaseId),
        installedRelease: null,
        installedProfiles: [],
        readiness: 'error',
        failedResources: [
          {
            path: failure.path,
            code: failure.code,
            attempts: 1,
          },
        ],
        canRetry: true,
        errorCode: failure.code,
      });
    });
    const gate = new CanonicalResourceReadinessGate({
      installer,
      probeRelease: async (releaseId) => primeResults(releaseId).slice(0, 11),
    });
    await gate.initialize();

    await expect(
      gate.apply({
        ...plan,
        planId: 'plan-b',
        releaseId: 'release-b',
      }),
    ).rejects.toMatchObject({
      code: 'storage',
    });
    expect(gate.getInstallerSnapshot()).toMatchObject({
      installedRelease: null,
      readiness: 'error',
      errorCode: 'storage',
      canRetry: true,
    });
  });
});
