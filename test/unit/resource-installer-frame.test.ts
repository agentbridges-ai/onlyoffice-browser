import { afterEach, describe, expect, it, vi } from 'vitest';
import type { RequiredReleaseIdentity, ResourceInstallerSnapshot, ResourcePlan } from '../../src/lib/release-resources';
import {
  RESOURCE_INSTALLER_CANONICAL_ORIGIN,
  RESOURCE_INSTALLER_FRAME_PROTOCOL,
  ResourceInstallerCapabilityGate,
  createResourceInstallerFrameUrl,
  parseResourceInstallerChallengeMessage,
  parseResourceInstallerInitMessage,
  parseResourceInstallerReconnectMessage,
  parseResourceInstallerRequestMessage,
  parseResourceInstallerServerMessage,
  resolveResourceInstallerClientIdentity,
  resolveResourceInstallerFrameIdentity,
} from '../../src/lib/resource-installer-frame-protocol';
import {
  ResourceInstallerFrameClient,
  ResourceInstallerFrameRpcError,
} from '../../src/lib/resource-installer-frame-client';
import {
  ResourceInstallerFrameSession,
  startResourceInstallerFrame,
  type ResourceInstallerFrameManager,
  type ResourceInstallerFramePort,
} from '../../src/resource-installer';
import type { ResourceInstallerRequestMessage } from '../../src/lib/resource-installer-frame-protocol';

const readySnapshot: ResourceInstallerSnapshot = {
  installedRelease: 'release-a',
  targetRelease: 'release-a',
  availableRelease: null,
  availablePackageVersion: null,
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
  installedProfiles: ['base', 'word', 'fonts-basic'],
};

const requiredReleaseIdentity: RequiredReleaseIdentity = {
  releaseId: 'release-b',
  manifestSha256: 'ab'.repeat(32),
  packageVersion: '0.5.7',
  hostBuildId: 'host-release-b',
};

const plan: ResourcePlan = {
  planId: 'plan-a',
  releaseId: 'release-b',
  scope: 'recommended',
  profiles: ['base', 'fonts-basic'],
  totalBytes: 1_000,
  downloadBytes: 300,
  reusedBytes: 700,
};

class FakePort implements ResourceInstallerFramePort {
  onmessage: ((event: MessageEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  readonly posted: unknown[] = [];
  peer: FakePort | null = null;
  started = false;
  closed = false;

  close(): void {
    this.closed = true;
  }

  postMessage(message: unknown): void {
    if (this.closed) throw new Error('closed');
    this.posted.push(message);
    if (this.peer) queueMicrotask(() => this.peer?.onmessage?.({ data: message } as MessageEvent));
  }

  start(): void {
    this.started = true;
  }

  emit(message: unknown): void {
    this.onmessage?.({ data: message } as MessageEvent);
  }
}

function portPair(): { port1: FakePort; port2: FakePort } {
  const port1 = new FakePort();
  const port2 = new FakePort();
  port1.peer = port2;
  port2.peer = port1;
  return { port1, port2 };
}

function deferred<T>() {
  let resolve!: (value: T | PromiseLike<T>) => void;
  let reject!: (error: unknown) => void;
  const promise = new Promise<T>((nextResolve, nextReject) => {
    resolve = nextResolve;
    reject = nextReject;
  });
  return { promise, resolve, reject };
}

function manager(overrides: Partial<ResourceInstallerFrameManager> = {}): ResourceInstallerFrameManager {
  return {
    getInstallerSnapshot: () => readySnapshot,
    subscribeInstaller: () => () => undefined,
    plan: async () => plan,
    apply: async () => undefined,
    checkForUpdates: async () => undefined,
    checkHealth: async () => undefined,
    repair: async () => undefined,
    pause: () => undefined,
    resume: async () => undefined,
    cancel: () => undefined,
    ...overrides,
  };
}

function frameClientHarness(options: {
  requestTimeoutMs?: number;
  snapshot?: ResourceInstallerSnapshot;
  onRequest?: (request: ResourceInstallerRequestMessage, port: FakePort) => void;
}) {
  const windowListeners = new Map<string, Set<(event: MessageEvent) => void>>();
  let snapshot = options.snapshot ?? readySnapshot;
  let currentClientPort: FakePort | null = null;
  let currentServerPort: FakePort | null = null;
  let challengeCount = 0;
  let reconnectCount = 0;
  let requestCount = 0;

  const fakeWindow = {
    location: { origin: 'https://piwork.getpi.work' },
    setTimeout: window.setTimeout.bind(window),
    clearTimeout: window.clearTimeout.bind(window),
    addEventListener(type: string, listener: (event: MessageEvent) => void) {
      const listeners = windowListeners.get(type) ?? new Set();
      listeners.add(listener);
      windowListeners.set(type, listeners);
    },
    removeEventListener(type: string, listener: (event: MessageEvent) => void) {
      windowListeners.get(type)?.delete(listener);
    },
  } as unknown as Window;

  const emitChallenge = () => {
    challengeCount += 1;
    const challenge = {
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'CHALLENGE',
      capability: {
        token: challengeCount.toString(16).padStart(64, '0'),
        parentOrigin: 'https://piwork.getpi.work',
        issuedAtMs: 1_000,
        expiresAtMs: 31_000,
      },
    };
    for (const listener of windowListeners.get('message') ?? []) {
      listener({
        data: challenge,
        origin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        source: contentWindow as unknown as Window,
      } as MessageEvent);
    }
  };

  const contentWindow = {
    postMessage(message: unknown, targetOrigin: string, transfer: Transferable[] = []) {
      expect(targetOrigin).toBe(RESOURCE_INSTALLER_CANONICAL_ORIGIN);
      if (typeof message === 'object' && message !== null && 'type' in message && message.type === 'RECONNECT') {
        reconnectCount += 1;
        emitChallenge();
        return;
      }
      const init = parseResourceInstallerInitMessage(message);
      expect(init).not.toBeNull();
      const port = transfer[0] as unknown as FakePort;
      currentServerPort = port;
      port.onmessage = (event) => {
        const request = parseResourceInstallerRequestMessage(event.data);
        expect(request).not.toBeNull();
        requestCount += 1;
        if (request?.command === 'SNAPSHOT') {
          port.postMessage({
            protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
            type: 'RESULT',
            id: request.id,
            command: 'SNAPSHOT',
            snapshot,
          });
          return;
        }
        if (request) options.onRequest?.(request, port);
      };
      port.start();
      port.postMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'READY',
        id: init!.id,
        command: 'INIT',
        snapshot,
      });
    },
  };
  const frame = {
    hidden: false,
    tabIndex: 0,
    title: '',
    src: '',
    contentWindow,
    referrerPolicy: '',
    setAttribute: vi.fn(),
    remove: vi.fn(),
  } as unknown as HTMLIFrameElement;
  const fakeDocument = {
    createElement: () => frame,
    body: {
      append: vi.fn(() => emitChallenge()),
    },
  } as unknown as Document;
  let uuid = 0;
  const client = new ResourceInstallerFrameClient({
    requiredReleaseIdentity,
    window: fakeWindow,
    document: fakeDocument,
    now: () => 2_000,
    requestTimeoutMs: options.requestTimeoutMs,
    connectTimeoutMs: options.requestTimeoutMs,
    randomUUID: () => `request-${++uuid}`,
    createMessageChannel: () => {
      const pair = portPair();
      currentClientPort = pair.port1;
      return pair as unknown as MessageChannel;
    },
  });

  return {
    client,
    get clientPort() {
      return currentClientPort;
    },
    get serverPort() {
      return currentServerPort;
    },
    get reconnectCount() {
      return reconnectCount;
    },
    get requestCount() {
      return requestCount;
    },
    setSnapshot(next: ResourceInstallerSnapshot) {
      snapshot = next;
    },
    event(next: ResourceInstallerSnapshot) {
      snapshot = next;
      currentServerPort?.postMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'EVENT',
        event: 'SNAPSHOT',
        snapshot,
      });
    },
  };
}

afterEach(() => {
  vi.useRealTimers();
});

describe('Resource Installer frame identity', () => {
  it('accepts only the two production parent origins', () => {
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        documentReferrer: 'https://piwork.getpi.work/settings',
        search: '',
      }),
    ).toEqual({
      canonicalOrigin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
      physicalParentOrigin: 'https://piwork.getpi.work',
      logicalParentOrigin: 'https://piwork.getpi.work',
      localTestMode: false,
    });
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        documentReferrer: 'https://onlyoffice.getpi.work/',
        search: '',
      })?.logicalParentOrigin,
    ).toBe('https://onlyoffice.getpi.work');
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        documentReferrer: 'https://piwork.getpi.work.evil.example/',
        search: '',
      }),
    ).toBeNull();
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        documentReferrer: 'https://aries.getpi.work/',
        search: '',
      }),
    ).toBeNull();
    expect(
      parseResourceInstallerReconnectMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'RECONNECT',
      }),
    ).not.toBeNull();
    expect(
      parseResourceInstallerReconnectMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'RECONNECT',
        capability: 'replay',
      }),
    ).toBeNull();
  });

  it('requires explicit localhost test mode and an identical protocol and port', () => {
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://piwork.localhost:8787/settings',
        search: '?testMode=1',
      }),
    ).toEqual({
      canonicalOrigin: 'http://onlyoffice.localhost:8787',
      physicalParentOrigin: 'http://piwork.localhost:8787',
      logicalParentOrigin: 'https://piwork.getpi.work',
      localTestMode: true,
    });
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://piwork.localhost:8787/settings',
        search: '',
      }),
    ).toBeNull();
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://piwork.localhost:9999/settings',
        search: '?testMode=1',
      }),
    ).toBeNull();
    expect(
      resolveResourceInstallerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://piwork.localhost:8787/settings',
        search: '?testMode=1&debug=1',
      }),
    ).toBeNull();
  });

  it('builds only the fixed canonical frame path', () => {
    const identity = resolveResourceInstallerClientIdentity({
      parentOrigin: 'https://piwork.getpi.work',
    });
    expect(identity).not.toBeNull();
    expect(createResourceInstallerFrameUrl(identity!).href).toBe(
      'https://onlyoffice.getpi.work/resource-installer.html',
    );
  });

  it('uses a progress idle timeout instead of cancelling a job after 30 seconds of continuous activity', async () => {
    vi.useFakeTimers();
    const applyGate = deferred<void>();
    let snapshot: ResourceInstallerSnapshot = {
      ...readySnapshot,
      installedRelease: null,
      targetRelease: 'release-b',
      readiness: 'updating' as const,
      phase: 'downloading' as const,
      downloadedBytes: 0,
      downloadBytes: 100,
      verifiedBytes: 0,
      verifyBytes: 100,
      canPause: true,
    };
    const subscription: { notify: ((next: ResourceInstallerSnapshot) => void) | null } = { notify: null };
    const cancel = vi.fn();
    const port = new FakePort();
    const session = new ResourceInstallerFrameSession({
      port,
      requestTimeoutMs: 30_000,
      manager: manager({
        getInstallerSnapshot: () => snapshot,
        subscribeInstaller(listener) {
          subscription.notify = listener;
          return () => {
            subscription.notify = null;
          };
        },
        apply: () => applyGate.promise,
        cancel,
      }),
    });
    session.start('init', true);
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'apply-progress',
      command: 'APPLY',
      plan,
    });

    for (let downloadedBytes = 20; downloadedBytes <= 80; downloadedBytes += 20) {
      await vi.advanceTimersByTimeAsync(20_000);
      snapshot = { ...snapshot, downloadedBytes };
      subscription.notify?.(snapshot);
    }
    expect(port.posted).not.toContainEqual(
      expect.objectContaining({ type: 'ERROR', id: 'apply-progress', error: { code: 'timeout', retryable: true } }),
    );
    expect(cancel).not.toHaveBeenCalled();

    snapshot = {
      ...readySnapshot,
      installedRelease: 'release-b',
      targetRelease: 'release-b',
    };
    subscription.notify?.(snapshot);
    applyGate.resolve();
    await vi.runAllTimersAsync();
    await Promise.resolve();
    expect(port.posted).toContainEqual(
      expect.objectContaining({ type: 'RESULT', id: 'apply-progress', command: 'APPLY' }),
    );
    session.destroy();
  });

  it.each(['CHECK_HEALTH', 'CHECK_UPDATES'] as const)(
    'keeps %s alive while readiness probes publish unchanged heartbeat snapshots',
    async (command) => {
      vi.useFakeTimers();
      const requestId = `probe-${command.toLowerCase().replace('_', '-')}`;
      const operationGate = deferred<void>();
      const subscription: { notify: ((next: ResourceInstallerSnapshot) => void) | null } = { notify: null };
      const port = new FakePort();
      const session = new ResourceInstallerFrameSession({
        port,
        requestTimeoutMs: 30_000,
        manager: manager({
          subscribeInstaller(listener) {
            subscription.notify = listener;
            return () => {
              subscription.notify = null;
            };
          },
          checkHealth: () => operationGate.promise,
          checkForUpdates: () => operationGate.promise,
        }),
      });
      session.start('init', true);
      port.emit({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: requestId,
        command,
      });

      for (let batch = 0; batch < 4; batch += 1) {
        await vi.advanceTimersByTimeAsync(10_000);
        subscription.notify?.(readySnapshot);
      }
      expect(port.posted).not.toContainEqual(
        expect.objectContaining({
          type: 'ERROR',
          id: requestId,
          error: { code: 'timeout', retryable: true },
        }),
      );

      operationGate.resolve();
      for (let turn = 0; turn < 4; turn += 1) await Promise.resolve();
      expect(port.posted).toContainEqual(
        expect.objectContaining({
          type: 'RESULT',
          id: requestId,
          command,
        }),
      );
      session.destroy();
    },
  );

  it('times out and cancels an active job after 30 seconds without a snapshot heartbeat', async () => {
    vi.useFakeTimers();
    const applyGate = deferred<void>();
    const cancel = vi.fn();
    const port = new FakePort();
    const session = new ResourceInstallerFrameSession({
      port,
      requestTimeoutMs: 30_000,
      manager: manager({
        apply: () => applyGate.promise,
        cancel,
      }),
    });
    session.start('init', false);
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'apply-stalled',
      command: 'APPLY',
      plan,
    });

    await vi.advanceTimersByTimeAsync(30_001);
    expect(cancel).toHaveBeenCalledOnce();
    expect(port.posted).toContainEqual({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'ERROR',
      id: 'apply-stalled',
      command: 'APPLY',
      error: { code: 'timeout', retryable: true },
    });
    applyGate.resolve();
    session.destroy();
  });

  it('keeps the original APPLY pending while paused for over 30 seconds and lets RESUME bypass the mutation queue', async () => {
    vi.useFakeTimers();
    const applyGate = deferred<void>();
    let snapshot: ResourceInstallerSnapshot = {
      ...readySnapshot,
      installedRelease: null,
      targetRelease: 'release-b',
      readiness: 'updating',
      phase: 'downloading',
      downloadedBytes: 10,
      downloadBytes: 100,
      verifiedBytes: 0,
      verifyBytes: 100,
      canPause: true,
    };
    const subscription: { notify: ((next: ResourceInstallerSnapshot) => void) | null } = { notify: null };
    const calls: string[] = [];
    const port = new FakePort();
    const session = new ResourceInstallerFrameSession({
      port,
      requestTimeoutMs: 30_000,
      manager: manager({
        getInstallerSnapshot: () => snapshot,
        subscribeInstaller(listener) {
          subscription.notify = listener;
          return () => {
            subscription.notify = null;
          };
        },
        async apply() {
          calls.push('apply:start');
          await applyGate.promise;
          calls.push('apply:end');
        },
        pause() {
          calls.push('pause');
          snapshot = {
            ...snapshot,
            readiness: 'paused',
            phase: 'paused',
            canPause: false,
            canResume: true,
          };
          subscription.notify?.(snapshot);
        },
        async resume() {
          calls.push('resume');
          snapshot = {
            ...snapshot,
            readiness: 'updating',
            phase: 'downloading',
            canPause: true,
            canResume: false,
          };
          subscription.notify?.(snapshot);
          applyGate.resolve();
          await applyGate.promise;
        },
      }),
    });
    session.start('init', true);
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'apply-paused',
      command: 'APPLY',
      plan,
    });
    await Promise.resolve();
    await Promise.resolve();
    await Promise.resolve();
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'pause-long',
      command: 'PAUSE',
    });
    await Promise.resolve();
    await vi.advanceTimersByTimeAsync(60_000);
    expect(port.posted).not.toContainEqual(expect.objectContaining({ type: 'ERROR', id: 'apply-paused' }));

    snapshot = {
      ...readySnapshot,
      installedRelease: 'release-b',
      targetRelease: 'release-b',
    };
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'resume-long',
      command: 'RESUME',
    });
    subscription.notify?.(snapshot);
    await vi.runAllTimersAsync();
    await Promise.resolve();
    await Promise.resolve();
    expect(calls).toEqual(['apply:start', 'pause', 'resume', 'apply:end']);
    expect(port.posted).toContainEqual(expect.objectContaining({ type: 'RESULT', id: 'apply-paused' }));
    expect(port.posted).toContainEqual(expect.objectContaining({ type: 'RESULT', id: 'resume-long' }));
    session.destroy();
  });
});

describe('Resource Installer capability and strict protocol', () => {
  it('binds a short-lived capability to the parent and burns it after one use', () => {
    const gate = new ResourceInstallerCapabilityGate({
      parentOrigin: 'https://piwork.getpi.work',
      now: () => 1_000,
      randomFill(bytes) {
        bytes.fill(0xab);
        return bytes;
      },
    });
    const capability = gate.issue();
    const init = {
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'init-1',
      command: 'INIT',
      capability: {
        token: capability.token,
        parentOrigin: capability.parentOrigin,
      },
      requiredReleaseIdentity,
      subscribe: true,
    };
    expect(gate.consume(init, 'https://piwork.getpi.work', 1_001)).toMatchObject({ ok: true });
    expect(gate.consume(init, 'https://piwork.getpi.work', 1_002)).toEqual({
      ok: false,
      code: 'invalid',
    });
  });

  it('rejects expired challenges, extra fields, malformed plans, and binary payloads', () => {
    const challenge = {
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'CHALLENGE',
      capability: {
        token: 'ab'.repeat(32),
        parentOrigin: 'https://piwork.getpi.work',
        issuedAtMs: 1_000,
        expiresAtMs: 31_000,
      },
    };
    expect(parseResourceInstallerChallengeMessage(challenge, 'https://piwork.getpi.work', 2_000)).toEqual(challenge);
    expect(parseResourceInstallerChallengeMessage(challenge, 'https://piwork.getpi.work', 31_000)).toBeNull();

    expect(
      parseResourceInstallerInitMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: 'init',
        command: 'INIT',
        capability: { token: 'ab'.repeat(32), parentOrigin: 'https://piwork.getpi.work' },
        subscribe: true,
        unexpected: true,
      }),
    ).toBeNull();
    expect(
      parseResourceInstallerInitMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: 'init',
        command: 'INIT',
        capability: { token: 'ab'.repeat(32), parentOrigin: 'https://piwork.getpi.work' },
        subscribe: true,
      }),
    ).toBeNull();
    expect(
      parseResourceInstallerInitMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: 'init',
        command: 'INIT',
        capability: { token: 'ab'.repeat(32), parentOrigin: 'https://piwork.getpi.work' },
        requiredReleaseIdentity: {
          ...requiredReleaseIdentity,
          packageVersion: '../stable',
        },
        subscribe: true,
      }),
    ).toBeNull();
    expect(
      parseResourceInstallerInitMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: 'init',
        command: 'INIT',
        capability: { token: 'ab'.repeat(32), parentOrigin: 'https://piwork.getpi.work' },
        requiredReleaseIdentity: {
          ...requiredReleaseIdentity,
          stable: true,
        },
        subscribe: true,
      }),
    ).toBeNull();
    expect(
      parseResourceInstallerRequestMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: 'apply',
        command: 'APPLY',
        plan: { ...plan, downloadBytes: 301 },
      }),
    ).toBeNull();
    expect(
      parseResourceInstallerRequestMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: 'apply',
        command: 'APPLY',
        plan,
        bytes: new ArrayBuffer(8),
      }),
    ).toBeNull();
  });

  it('strictly validates typed results, errors, and snapshots', () => {
    expect(
      parseResourceInstallerServerMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'READY',
        id: 'init',
        command: 'INIT',
        snapshot: readySnapshot,
      }),
    ).not.toBeNull();
    expect(
      parseResourceInstallerServerMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'RESULT',
        id: 'plan',
        command: 'PLAN',
        plan,
      }),
    ).not.toBeNull();
    expect(
      parseResourceInstallerServerMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'ERROR',
        id: 'apply',
        command: 'APPLY',
        error: { code: 'quota', retryable: true },
      }),
    ).not.toBeNull();
    expect(
      parseResourceInstallerServerMessage({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'EVENT',
        event: 'SNAPSHOT',
        snapshot: { ...readySnapshot, downloadedBytes: 101 },
      }),
    ).toBeNull();
  });
});

describe('Resource Installer frame session', () => {
  it('returns READY, plans without resource bytes, and publishes subscribed snapshots', async () => {
    const port = new FakePort();
    const subscription: { notify: ((snapshot: ResourceInstallerSnapshot) => void) | null } = {
      notify: null,
    };
    const session = new ResourceInstallerFrameSession({
      port,
      manager: manager({
        subscribeInstaller(listener) {
          subscription.notify = listener;
          return () => {
            subscription.notify = null;
          };
        },
      }),
    });
    session.start('init', true);
    expect(port.posted[0]).toMatchObject({ type: 'READY', snapshot: readySnapshot });

    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'plan',
      command: 'PLAN',
      request: { scope: 'recommended' },
    });
    await vi.waitFor(() => {
      expect(port.posted).toContainEqual({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'RESULT',
        id: 'plan',
        command: 'PLAN',
        plan,
      });
    });
    expect(port.posted.some((message) => message instanceof ArrayBuffer)).toBe(false);

    subscription.notify?.({ ...readySnapshot, readiness: 'update-available', availableRelease: 'release-b' });
    expect(port.posted.at(-1)).toMatchObject({
      type: 'EVENT',
      event: 'SNAPSHOT',
      snapshot: { availableRelease: 'release-b' },
    });
    session.destroy();
    expect(subscription.notify).toBeNull();
  });

  it('serializes state mutations while allowing PAUSE to interrupt an APPLY', async () => {
    const applyGate = deferred<void>();
    const calls: string[] = [];
    const port = new FakePort();
    const session = new ResourceInstallerFrameSession({
      port,
      manager: manager({
        async apply() {
          calls.push('apply:start');
          await applyGate.promise;
          calls.push('apply:end');
        },
        async repair() {
          calls.push('repair');
        },
        pause() {
          calls.push('pause');
        },
      }),
    });
    session.start('init', false);
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'apply',
      command: 'APPLY',
      plan,
    });
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'repair',
      command: 'REPAIR',
      scope: 'installed',
    });
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'pause',
      command: 'PAUSE',
    });
    await vi.waitFor(() => {
      expect(calls).toContain('apply:start');
      expect(calls).toContain('pause');
      expect(calls).not.toContain('repair');
    });
    expect(port.posted).toContainEqual(expect.objectContaining({ type: 'RESULT', id: 'pause' }));

    applyGate.resolve();
    await vi.waitFor(() => {
      expect(calls.at(-2)).toBe('apply:end');
      expect(calls.at(-1)).toBe('repair');
    });
    session.destroy();
  });

  it('returns a structured timeout and never starts a queued mutation after its deadline', async () => {
    vi.useFakeTimers();
    const first = deferred<void>();
    const calls: string[] = [];
    const port = new FakePort();
    const session = new ResourceInstallerFrameSession({
      port,
      requestTimeoutMs: 100,
      manager: manager({
        async apply() {
          calls.push('apply');
          await first.promise;
        },
        async repair() {
          calls.push('repair');
        },
      }),
    });
    session.start('init', false);
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'apply',
      command: 'APPLY',
      plan,
    });
    port.emit({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'repair',
      command: 'REPAIR',
      scope: 'installed',
    });
    await vi.advanceTimersByTimeAsync(101);
    expect(port.posted).toContainEqual({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'ERROR',
      id: 'repair',
      command: 'REPAIR',
      error: { code: 'timeout', retryable: true },
    });
    first.resolve();
    await Promise.resolve();
    await Promise.resolve();
    expect(calls).toEqual(['apply']);
    session.destroy();
  });
});

describe('Resource Installer frame bootstrap', () => {
  it('checks the exact parent source/origin before consuming the one-time capability', async () => {
    const listeners = new Map<string, Set<(event: MessageEvent) => void>>();
    const challenges: Array<{ message: unknown; targetOrigin: string }> = [];
    const parentWindow = {
      postMessage(message: unknown, targetOrigin: string) {
        challenges.push({ message, targetOrigin });
      },
    };
    const fakeWindow = {
      parent: parentWindow,
      location: {
        origin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        search: '',
      },
      setTimeout: window.setTimeout.bind(window),
      clearTimeout: window.clearTimeout.bind(window),
      addEventListener(type: string, listener: (event: MessageEvent) => void) {
        const entries = listeners.get(type) ?? new Set();
        entries.add(listener);
        listeners.set(type, entries);
      },
      removeEventListener(type: string, listener: (event: MessageEvent) => void) {
        listeners.get(type)?.delete(listener);
      },
    } as unknown as Window;
    const fakeDocument = {
      referrer: 'https://piwork.getpi.work/settings',
    } as Document;
    const cancel = vi.fn();
    const managed = manager({ cancel });
    const createManager = vi.fn(() => managed);
    const destroy = startResourceInstallerFrame({
      window: fakeWindow,
      document: fakeDocument,
      createManager,
      now: () => 1_000,
      randomFill(bytes) {
        bytes.fill(0xab);
        return bytes;
      },
    });
    expect(challenges).toHaveLength(1);
    expect(challenges[0]?.targetOrigin).toBe('https://piwork.getpi.work');
    const challenge = parseResourceInstallerChallengeMessage(
      challenges[0]?.message,
      'https://piwork.getpi.work',
      1_001,
    );
    expect(challenge).not.toBeNull();
    const init = {
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: 'init',
      command: 'INIT',
      capability: {
        token: challenge!.capability.token,
        parentOrigin: 'https://piwork.getpi.work',
      },
      requiredReleaseIdentity,
      subscribe: true,
    };

    const ignoredPort = new FakePort();
    for (const listener of listeners.get('message') ?? []) {
      listener({
        data: init,
        origin: 'https://piwork.getpi.work',
        source: {} as Window,
        ports: [ignoredPort as unknown as MessagePort],
      } as unknown as MessageEvent);
    }
    expect(createManager).not.toHaveBeenCalled();
    expect(ignoredPort.closed).toBe(false);

    const connectedPort = new FakePort();
    for (const listener of listeners.get('message') ?? []) {
      listener({
        data: init,
        origin: 'https://piwork.getpi.work',
        source: parentWindow as unknown as Window,
        ports: [connectedPort as unknown as MessagePort],
      } as unknown as MessageEvent);
    }
    await vi.waitFor(() => expect(createManager).toHaveBeenCalledOnce());
    expect(createManager).toHaveBeenCalledWith(requiredReleaseIdentity, expect.any(Function));
    await vi.waitFor(() => {
      expect(connectedPort.posted).toContainEqual(
        expect.objectContaining({ type: 'READY', id: 'init', command: 'INIT' }),
      );
    });

    for (const listener of listeners.get('message') ?? []) {
      listener({
        data: {
          protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
          type: 'RECONNECT',
        },
        origin: 'https://piwork.getpi.work',
        source: parentWindow as unknown as Window,
        ports: [],
      } as unknown as MessageEvent);
    }
    expect(connectedPort.closed).toBe(true);
    expect(cancel).not.toHaveBeenCalled();
    expect(challenges).toHaveLength(2);
    const reconnectChallenge = parseResourceInstallerChallengeMessage(
      challenges[1]?.message,
      'https://piwork.getpi.work',
      1_001,
    );
    expect(reconnectChallenge).not.toBeNull();
    const tamperedPort = new FakePort();
    for (const listener of listeners.get('message') ?? []) {
      listener({
        data: {
          ...init,
          id: 'init-tampered',
          capability: {
            token: reconnectChallenge!.capability.token,
            parentOrigin: 'https://piwork.getpi.work',
          },
          requiredReleaseIdentity: {
            ...requiredReleaseIdentity,
            releaseId: 'release-c',
          },
        },
        origin: 'https://piwork.getpi.work',
        source: parentWindow as unknown as Window,
        ports: [tamperedPort as unknown as MessagePort],
      } as unknown as MessageEvent);
    }
    await vi.waitFor(() => {
      expect(tamperedPort.posted).toContainEqual({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'ERROR',
        id: 'init-tampered',
        command: 'INIT',
        error: {
          code: 'incompatible',
          retryable: false,
          path: 'release/identity',
        },
      });
    });
    expect(createManager).toHaveBeenCalledOnce();
    expect(challenges).toHaveLength(3);
    const retryChallenge = parseResourceInstallerChallengeMessage(
      challenges[2]?.message,
      'https://piwork.getpi.work',
      1_001,
    );
    expect(retryChallenge).not.toBeNull();
    const reconnectedPort = new FakePort();
    for (const listener of listeners.get('message') ?? []) {
      listener({
        data: {
          ...init,
          id: 'init-reconnected',
          capability: {
            token: retryChallenge!.capability.token,
            parentOrigin: 'https://piwork.getpi.work',
          },
        },
        origin: 'https://piwork.getpi.work',
        source: parentWindow as unknown as Window,
        ports: [reconnectedPort as unknown as MessagePort],
      } as unknown as MessageEvent);
    }
    await vi.waitFor(() => {
      expect(reconnectedPort.posted).toContainEqual(
        expect.objectContaining({ type: 'READY', id: 'init-reconnected', command: 'INIT' }),
      );
    });
    expect(createManager).toHaveBeenCalledOnce();
    destroy();
    expect(reconnectedPort.closed).toBe(true);
    expect(cancel).toHaveBeenCalledOnce();
  });

  it('keeps INIT alive while four readiness batches report progress over 40 seconds', async () => {
    vi.useFakeTimers();
    const listeners = new Map<string, Set<(event: MessageEvent) => void>>();
    const challenges: Array<{ message: unknown; targetOrigin: string }> = [];
    const parentWindow = {
      postMessage(message: unknown, targetOrigin: string) {
        challenges.push({ message, targetOrigin });
      },
    };
    const fakeWindow = {
      parent: parentWindow,
      location: {
        origin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        search: '',
      },
      setTimeout: window.setTimeout.bind(window),
      clearTimeout: window.clearTimeout.bind(window),
      addEventListener(type: string, listener: (event: MessageEvent) => void) {
        const entries = listeners.get(type) ?? new Set();
        entries.add(listener);
        listeners.set(type, entries);
      },
      removeEventListener(type: string, listener: (event: MessageEvent) => void) {
        listeners.get(type)?.delete(listener);
      },
    } as unknown as Window;
    const fakeDocument = {
      referrer: 'https://piwork.getpi.work/settings',
    } as Document;
    const creation = deferred<ResourceInstallerFrameManager>();
    let reportProgress: (() => void) | undefined;
    const createManager = vi.fn((_identity: RequiredReleaseIdentity, nextReportProgress?: () => void) => {
      reportProgress = nextReportProgress;
      return creation.promise;
    });
    const managed = manager();
    const destroy = startResourceInstallerFrame({
      window: fakeWindow,
      document: fakeDocument,
      createManager,
      requestTimeoutMs: 30_000,
      now: () => 1_000,
      randomFill(bytes) {
        bytes.fill(0xab);
        return bytes;
      },
    });
    const challenge = parseResourceInstallerChallengeMessage(
      challenges[0]?.message,
      'https://piwork.getpi.work',
      1_001,
    );
    expect(challenge).not.toBeNull();
    const port = new FakePort();
    for (const listener of listeners.get('message') ?? []) {
      listener({
        data: {
          protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
          type: 'REQUEST',
          id: 'init-progress',
          command: 'INIT',
          capability: {
            token: challenge!.capability.token,
            parentOrigin: 'https://piwork.getpi.work',
          },
          requiredReleaseIdentity,
          subscribe: true,
        },
        origin: 'https://piwork.getpi.work',
        source: parentWindow as unknown as Window,
        ports: [port as unknown as MessagePort],
      } as unknown as MessageEvent);
    }
    await Promise.resolve();
    await Promise.resolve();
    expect(createManager).toHaveBeenCalledWith(requiredReleaseIdentity, expect.any(Function));

    for (let batch = 1; batch <= 4; batch += 1) {
      await vi.advanceTimersByTimeAsync(10_000);
      reportProgress?.();
      expect(port.posted).not.toContainEqual(
        expect.objectContaining({ type: 'ERROR', id: 'init-progress', error: { code: 'timeout' } }),
      );
    }

    creation.resolve(managed);
    for (let turn = 0; turn < 6; turn += 1) await Promise.resolve();
    expect(
      port.posted.some(
        (message) =>
          typeof message === 'object' &&
          message !== null &&
          'type' in message &&
          message.type === 'READY' &&
          'id' in message &&
          message.id === 'init-progress' &&
          'command' in message &&
          message.command === 'INIT',
      ),
    ).toBe(true);
    destroy();
  });
});

describe('Resource Installer frame client', () => {
  it('fails closed before creating a remote frame when the authorized release identity is missing', async () => {
    const createElement = vi.fn();
    const client = new ResourceInstallerFrameClient({
      window: {
        location: { origin: 'https://piwork.getpi.work' },
      } as unknown as Window,
      document: {
        body: {},
        createElement,
      } as unknown as Document,
    });

    await expect(client.connect()).rejects.toMatchObject({
      code: 'identity',
      retryable: false,
    });
    expect(createElement).not.toHaveBeenCalled();
    client.destroy();
  });

  it('handshakes only with the exact iframe source/origin and uses the exact targetOrigin', async () => {
    const windowListeners = new Map<string, Set<(event: MessageEvent) => void>>();
    const fakeWindow = {
      location: { origin: 'https://piwork.getpi.work' },
      setTimeout: window.setTimeout.bind(window),
      clearTimeout: window.clearTimeout.bind(window),
      addEventListener(type: string, listener: (event: MessageEvent) => void) {
        const listeners = windowListeners.get(type) ?? new Set();
        listeners.add(listener);
        windowListeners.set(type, listeners);
      },
      removeEventListener(type: string, listener: (event: MessageEvent) => void) {
        windowListeners.get(type)?.delete(listener);
      },
    } as unknown as Window;
    let remotePort: FakePort | null = null;
    const postTargets: string[] = [];
    const contentWindow = {
      postMessage(message: unknown, targetOrigin: string, transfer: Transferable[]) {
        postTargets.push(targetOrigin);
        const init = parseResourceInstallerInitMessage(message);
        expect(init).not.toBeNull();
        expect(init?.requiredReleaseIdentity).toEqual(requiredReleaseIdentity);
        remotePort = transfer[0] as unknown as FakePort;
        remotePort.onmessage = (event) => {
          const request = parseResourceInstallerRequestMessage(event.data);
          if (request?.command === 'SNAPSHOT') {
            remotePort?.postMessage({
              protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
              type: 'RESULT',
              id: request.id,
              command: 'SNAPSHOT',
              snapshot: readySnapshot,
            });
          }
        };
        remotePort.start();
        remotePort.postMessage({
          protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
          type: 'READY',
          id: init!.id,
          command: 'INIT',
          snapshot: readySnapshot,
        });
      },
    };
    const frame = {
      hidden: false,
      tabIndex: 0,
      title: '',
      src: '',
      contentWindow,
      referrerPolicy: '',
      setAttribute: vi.fn(),
      remove: vi.fn(),
    } as unknown as HTMLIFrameElement;
    const fakeDocument = {
      createElement: () => frame,
      body: { append: vi.fn() },
    } as unknown as Document;
    const pair = portPair();
    let uuid = 0;
    const client = new ResourceInstallerFrameClient({
      requiredReleaseIdentity,
      window: fakeWindow,
      document: fakeDocument,
      now: () => 2_000,
      randomUUID: () => `request-${++uuid}`,
      createMessageChannel: () => pair as unknown as MessageChannel,
    });
    const connected = client.connect();
    expect(frame.src).toBe('https://onlyoffice.getpi.work/resource-installer.html');

    const challenge = {
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'CHALLENGE',
      capability: {
        token: 'ab'.repeat(32),
        parentOrigin: 'https://piwork.getpi.work',
        issuedAtMs: 1_000,
        expiresAtMs: 31_000,
      },
    };
    for (const listener of windowListeners.get('message') ?? []) {
      listener({
        data: challenge,
        origin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        source: {} as Window,
      } as MessageEvent);
    }
    expect(postTargets).toEqual([]);
    for (const listener of windowListeners.get('message') ?? []) {
      listener({
        data: challenge,
        origin: RESOURCE_INSTALLER_CANONICAL_ORIGIN,
        source: contentWindow as unknown as Window,
      } as MessageEvent);
    }
    await expect(connected).resolves.toEqual(readySnapshot);
    expect(postTargets).toEqual([RESOURCE_INSTALLER_CANONICAL_ORIGIN]);
    await expect(client.refreshSnapshot()).resolves.toEqual(readySnapshot);
    client.destroy();
    expect(frame.remove).toHaveBeenCalledOnce();
  });

  it('renews the client idle deadline from valid progress snapshots beyond 30 seconds', async () => {
    vi.useFakeTimers();
    let applyRequest: Extract<ResourceInstallerRequestMessage, { command: 'APPLY' }> | null = null;
    const downloading: ResourceInstallerSnapshot = {
      ...readySnapshot,
      installedRelease: null,
      targetRelease: 'release-b',
      readiness: 'updating',
      phase: 'downloading',
      downloadedBytes: 0,
      downloadBytes: 100,
      verifiedBytes: 0,
      verifyBytes: 100,
      canPause: true,
    };
    const harness = frameClientHarness({
      requestTimeoutMs: 30_000,
      snapshot: downloading,
      onRequest(request) {
        if (request.command === 'APPLY') applyRequest = request;
      },
    });
    await harness.client.connect();
    const applying = harness.client.apply(plan);
    await vi.advanceTimersByTimeAsync(0);
    expect(applyRequest).not.toBeNull();

    for (let downloadedBytes = 20; downloadedBytes <= 80; downloadedBytes += 20) {
      await vi.advanceTimersByTimeAsync(20_000);
      harness.event({ ...downloading, downloadedBytes });
      await vi.advanceTimersByTimeAsync(0);
    }

    const complete = {
      ...readySnapshot,
      installedRelease: 'release-b',
      targetRelease: 'release-b',
    };
    harness.serverPort?.postMessage({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'RESULT',
      id: applyRequest!.id,
      command: 'APPLY',
      snapshot: complete,
    });
    await expect(applying).resolves.toBeUndefined();
    harness.client.destroy();
  });

  it('returns a structured client timeout after 30 seconds without progress', async () => {
    vi.useFakeTimers();
    const harness = frameClientHarness({
      requestTimeoutMs: 30_000,
      snapshot: {
        ...readySnapshot,
        installedRelease: null,
        targetRelease: 'release-b',
        readiness: 'updating',
        phase: 'downloading',
        downloadedBytes: 0,
        downloadBytes: 100,
        verifiedBytes: 0,
        verifyBytes: 100,
        canPause: true,
      },
    });
    await harness.client.connect();
    const applying = harness.client.apply(plan);
    const timedOut = expect(applying).rejects.toMatchObject({ code: 'timeout', retryable: true });
    await vi.advanceTimersByTimeAsync(30_001);
    await timedOut;
    harness.client.destroy();
  });

  it('keeps the client APPLY promise alive while paused and sends RESUME outside its mutation queue', async () => {
    vi.useFakeTimers();
    let applyRequest: Extract<ResourceInstallerRequestMessage, { command: 'APPLY' }> | null = null;
    let harness!: ReturnType<typeof frameClientHarness>;
    const downloading: ResourceInstallerSnapshot = {
      ...readySnapshot,
      installedRelease: null,
      targetRelease: 'release-b',
      readiness: 'updating',
      phase: 'downloading',
      downloadedBytes: 10,
      downloadBytes: 100,
      verifiedBytes: 0,
      verifyBytes: 100,
      canPause: true,
    };
    harness = frameClientHarness({
      requestTimeoutMs: 30_000,
      snapshot: downloading,
      onRequest(request, port) {
        if (request.command === 'APPLY') {
          applyRequest = request;
          return;
        }
        if (request.command === 'PAUSE') {
          const paused: ResourceInstallerSnapshot = {
            ...downloading,
            readiness: 'paused',
            phase: 'paused',
            canPause: false,
            canResume: true,
          };
          harness.event(paused);
          port.postMessage({
            protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
            type: 'RESULT',
            id: request.id,
            command: 'PAUSE',
            snapshot: paused,
          });
          return;
        }
        if (request.command === 'RESUME') {
          harness.event(downloading);
          const complete = {
            ...readySnapshot,
            installedRelease: 'release-b',
            targetRelease: 'release-b',
          };
          port.postMessage({
            protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
            type: 'RESULT',
            id: applyRequest!.id,
            command: 'APPLY',
            snapshot: complete,
          });
          port.postMessage({
            protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
            type: 'RESULT',
            id: request.id,
            command: 'RESUME',
            snapshot: complete,
          });
        }
      },
    });
    await harness.client.connect();
    const applying = harness.client.apply(plan);
    await vi.advanceTimersByTimeAsync(0);
    await harness.client.pause();
    await vi.advanceTimersByTimeAsync(60_000);

    const resuming = harness.client.resume();
    await vi.advanceTimersByTimeAsync(0);
    await expect(Promise.all([applying, resuming])).resolves.toEqual([undefined, undefined]);
    harness.client.destroy();
  });

  it('reconnects a broken port and completes the original APPLY without sending APPLY twice', async () => {
    vi.useFakeTimers();
    let applyRequests = 0;
    const downloading: ResourceInstallerSnapshot = {
      ...readySnapshot,
      installedRelease: null,
      targetRelease: 'release-b',
      readiness: 'updating',
      phase: 'downloading',
      downloadedBytes: 40,
      downloadBytes: 100,
      verifiedBytes: 40,
      verifyBytes: 100,
      canPause: true,
    };
    const harness = frameClientHarness({
      requestTimeoutMs: 30_000,
      snapshot: downloading,
      onRequest(request) {
        if (request.command === 'APPLY') applyRequests += 1;
      },
    });
    await harness.client.connect();
    const applying = harness.client.apply(plan);
    await vi.advanceTimersByTimeAsync(0);
    expect(applyRequests).toBe(1);

    harness.clientPort?.onmessageerror?.({ data: null } as MessageEvent);
    await vi.advanceTimersByTimeAsync(0);
    expect(harness.reconnectCount).toBe(1);
    expect(applyRequests).toBe(1);

    harness.event({
      ...readySnapshot,
      installedRelease: 'release-b',
      targetRelease: 'release-b',
    });
    await vi.advanceTimersByTimeAsync(0);
    await expect(applying).resolves.toBeUndefined();
    expect(applyRequests).toBe(1);
    harness.client.destroy();
  });

  it('surfaces structured remote errors', () => {
    const error = new ResourceInstallerFrameRpcError({ code: 'quota', retryable: true, path: 'fonts/Aptos.ttf' });
    expect(error).toMatchObject({
      code: 'quota',
      retryable: true,
      path: 'fonts/Aptos.ttf',
    });
  });
});
