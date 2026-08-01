import { describe, expect, it, vi } from 'vitest';
import {
  CANONICAL_RESOURCE_BROKER_MAX_RESERVED_BYTES,
  CanonicalResourceBrokerService,
  parseCanonicalResourceBrokerClientUrl,
} from '../../src/lib/canonical-resource-broker-service';
import type { AssetContentMapping, ReleaseTransactionRecord } from '../../src/lib/release-content-model';
import { RESOURCE_BROKER_PROTOCOL, type ResourceBrokerServerMessage } from '../../src/lib/resource-broker-protocol';

const digest = (character: string) => character.repeat(64);

class FakePort {
  onmessage: ((event: MessageEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  readonly messages: ResourceBrokerServerMessage[] = [];
  closed = false;

  close(): void {
    this.closed = true;
  }

  postMessage(message: unknown): void {
    this.messages.push(message as ResourceBrokerServerMessage);
  }

  start(): void {}

  send(message: unknown): void {
    this.onmessage?.({ data: message } as MessageEvent);
  }
}

const mapping: AssetContentMapping = {
  releaseId: 'release-a',
  path: 'sdkjs/app.js',
  assetSha256: digest('a'),
  assetBytes: 6,
  mime: 'text/javascript',
  representationId: 'whole',
  representationKind: 'whole',
  spans: [{ objectSha256: digest('a'), objectOffset: 0, assetOffset: 0, bytes: 6 }],
};

const transaction: ReleaseTransactionRecord = {
  releaseId: 'release-a',
  manifestSha256: digest('b'),
  storageSetSha256: digest('c'),
  planFingerprint: digest('d'),
  state: 'active',
  previousActiveReleaseId: null,
  requiredObjects: [{ sha256: digest('a'), bytes: 6 }],
  plannedMappings: [mapping],
  committedMappings: [mapping],
  updatedAt: 1,
};

function createService(
  overrides: {
    state?: ReleaseTransactionRecord['state'];
    body?: string;
    healthReady?: boolean;
    onProbe?: () => void;
    probeGate?: Promise<void>;
    now?: () => number;
    maxActiveReads?: number;
  } = {},
) {
  const body = overrides.body ?? 'abcdef';
  const state = overrides.state ?? 'active';
  const record = { sha256: digest('a'), bytes: 6, verifiedAt: 1 };
  return new CanonicalResourceBrokerService({
    journal: {
      async getAssetReadView(_releaseId, path) {
        return {
          release: {
            releaseId: 'release-a',
            manifestSha256: digest('b'),
            storageSetSha256: digest('c'),
            state,
            updatedAt: 1,
          },
          mapping: path === mapping.path ? mapping : null,
          objects: path === mapping.path ? [record] : [],
        };
      },
    },
    store: {
      async probeRelease() {
        overrides.onProbe?.();
        await overrides.probeGate;
        return {
          releaseId: 'release-a',
          state,
          ready: overrides.healthReady ?? true,
          probeSucceeded: overrides.healthReady ?? true,
          probePath: mapping.path,
          probeAssetBytes: mapping.assetBytes,
          probeAssetSha256: mapping.assetSha256,
        };
      },
      async matchObject() {
        return new Response(body, {
          headers: {
            'content-length': String(body.length),
          },
        });
      },
    },
    now: overrides.now,
    maxActiveReads: overrides.maxActiveReads,
  });
}

describe('parseCanonicalResourceBrokerClientUrl', () => {
  it('binds browser timer methods before storing them on the canonical service', async () => {
    const originalSetTimeout = globalThis.setTimeout;
    const originalClearTimeout = globalThis.clearTimeout;
    const checkedSetTimeout = function (
      this: typeof globalThis,
      handler: TimerHandler,
      timeout?: number,
      ...args: unknown[]
    ) {
      if (this !== globalThis) throw new TypeError('Illegal invocation');
      return originalSetTimeout(handler, timeout, ...args);
    } as typeof globalThis.setTimeout;
    const checkedClearTimeout = function (this: typeof globalThis, timer: ReturnType<typeof globalThis.setTimeout>) {
      if (this !== globalThis) throw new TypeError('Illegal invocation');
      return originalClearTimeout(timer);
    } as typeof globalThis.clearTimeout;
    vi.stubGlobal('setTimeout', checkedSetTimeout);
    vi.stubGlobal('clearTimeout', checkedClearTimeout);
    try {
      const service = createService();
      const port = new FakePort();
      const handled = service.handle(
        {
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'READ',
          id: 'read-browser-timer',
          releaseId: 'release-a',
          sessionId: 'session-a',
          path: mapping.path,
          windowBytes: 256 * 1024,
        },
        port,
      );
      await Promise.resolve();
      port.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'CANCEL', id: 'read-browser-timer' });
      await handled;
      expect(port.messages.some((message) => message.type === 'CANCELLED')).toBe(true);
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('accepts only the exact canonical broker client path and identity', () => {
    expect(
      parseCanonicalResourceBrokerClientUrl(
        'https://onlyoffice.getpi.work/resource-broker.html?releaseId=release-a&sessionId=session-a&parentOrigin=https%3A%2F%2Fpiwork.getpi.work',
        'https://onlyoffice.getpi.work',
      ),
    ).toEqual({
      releaseId: 'release-a',
      sessionId: 'session-a',
      parentOrigin: 'https://piwork.getpi.work',
      localMatrix: false,
    });
    expect(
      parseCanonicalResourceBrokerClientUrl(
        'https://onlyoffice.getpi.work/resource-broker.html?releaseId=release-a&sessionId=session-a&parentOrigin=https%3A%2F%2Fevil.example',
        'https://onlyoffice.getpi.work',
      ),
    ).toBeNull();
    expect(
      parseCanonicalResourceBrokerClientUrl(
        'https://onlyoffice.getpi.work/resource-broker.html?releaseId=release-a&sessionId=session-a&parentOrigin=https%3A%2F%2Fpiwork.getpi.work&extra=1',
        'https://onlyoffice.getpi.work',
      ),
    ).toBeNull();
  });

  it('serves a path from a thousand-object release with one journal read view', async () => {
    const objectCount = 1_024;
    const objects = Array.from({ length: objectCount }, (_, index) => ({
      sha256: index.toString(16).padStart(64, '0'),
      bytes: 1,
      verifiedAt: 1,
    }));
    const mappings: AssetContentMapping[] = objects.map((object, index) => ({
      releaseId: 'release-large',
      path: `sdkjs/assets/${index}.bin`,
      assetSha256: object.sha256,
      assetBytes: 1,
      representationId: 'whole',
      representationKind: 'whole',
      spans: [
        {
          objectSha256: object.sha256,
          objectOffset: 0,
          assetOffset: 0,
          bytes: 1,
        },
      ],
    }));
    const largeTransaction: ReleaseTransactionRecord = {
      ...transaction,
      releaseId: 'release-large',
      requiredObjects: objects.map(({ sha256, bytes }) => ({ sha256, bytes })),
      plannedMappings: mappings,
      committedMappings: mappings,
    };
    const getAssetReadView = vi.fn(async () => ({
      release: {
        releaseId: largeTransaction.releaseId,
        manifestSha256: largeTransaction.manifestSha256,
        storageSetSha256: largeTransaction.storageSetSha256,
        state: largeTransaction.state,
        updatedAt: largeTransaction.updatedAt,
      },
      mapping: mappings.at(-1)!,
      objects,
    }));
    const matchObject = vi.fn(async () => new Response(Uint8Array.of(42)));
    const service = new CanonicalResourceBrokerService({
      journal: { getAssetReadView },
      store: {
        async probeRelease() {
          throw new Error('not used');
        },
        matchObject,
      },
    });
    const port = new FakePort();
    const handling = service.handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-large',
        releaseId: 'release-large',
        sessionId: 'session-large',
        path: `sdkjs/assets/${objectCount - 1}.bin`,
        windowBytes: 1,
      },
      port,
    );
    await vi.waitFor(() => expect(port.messages[0]?.type).toBe('HEADERS'));
    port.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'PULL', id: 'read-large' });
    await handling;

    expect(getAssetReadView).toHaveBeenCalledTimes(1);
    expect(matchObject).toHaveBeenCalledTimes(1);
    expect(matchObject).toHaveBeenCalledWith(objects.at(-1), objects.at(-1));
    expect(port.messages.map((message) => message.type)).toEqual(['HEADERS', 'CHUNK', 'END']);
  });
});

describe('CanonicalResourceBrokerService', () => {
  it('probes only a complete and readable release', async () => {
    const port = new FakePort();
    await createService().handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PROBE',
        id: 'probe-a',
        releaseId: 'release-a',
        sessionId: 'session-a',
      },
      port,
    );
    expect(port.messages).toEqual([
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'RESULT',
        id: 'probe-a',
        value: {
          releaseId: 'release-a',
          ready: true,
          probePath: mapping.path,
          probeBytes: mapping.assetBytes,
          probeSha256: mapping.assetSha256,
        },
      },
    ]);
    expect(port.closed).toBe(true);
  });

  it('coalesces one readiness burst and performs a new real read after the bounded reuse window', async () => {
    let releaseProbeCount = 0;
    let now = 1_000;
    let releaseProbe!: () => void;
    const probeGate = new Promise<void>((resolve) => {
      releaseProbe = resolve;
    });
    const service = createService({
      onProbe: () => {
        releaseProbeCount += 1;
      },
      probeGate,
      now: () => now,
    });
    const first = new FakePort();
    const second = new FakePort();
    const request = {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'PROBE' as const,
      releaseId: 'release-a',
      sessionId: 'session-a',
    };
    const handling = [
      service.handle({ ...request, id: 'probe-first' }, first),
      service.handle({ ...request, id: 'probe-second' }, second),
    ];
    await vi.waitFor(() => expect(releaseProbeCount).toBe(1));
    releaseProbe();
    await Promise.all(handling);
    expect(first.messages[0]?.type).toBe('RESULT');
    expect(second.messages[0]?.type).toBe('RESULT');

    await service.handle({ ...request, id: 'probe-later' }, new FakePort());
    expect(releaseProbeCount).toBe(1);
    now += 30_001;
    await service.handle({ ...request, id: 'probe-after-expiry' }, new FakePort());
    expect(releaseProbeCount).toBe(2);
  });

  it('streams one bounded chunk per pull with exact MIME and range headers', async () => {
    const port = new FakePort();
    const service = createService();
    const handling = service.handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-a',
        releaseId: 'release-a',
        sessionId: 'session-a',
        path: 'sdkjs/app.js',
        range: { kind: 'closed', start: 1, end: 4 },
        windowBytes: 2,
      },
      port,
    );
    await vi.waitFor(() => expect(port.messages[0]?.type).toBe('HEADERS'));
    port.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'PULL', id: 'read-a' });
    await vi.waitFor(() => expect(port.messages.some((message) => message.type === 'CHUNK')).toBe(true));
    port.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'PULL', id: 'read-a' });
    await handling;

    expect(port.messages[0]).toMatchObject({
      type: 'HEADERS',
      status: 206,
      headers: {
        contentLength: 4,
        contentRange: 'bytes 1-4/6',
        contentType: 'text/javascript',
      },
    });
    const chunks = port.messages.filter(
      (message): message is Extract<ResourceBrokerServerMessage, { type: 'CHUNK' }> => message.type === 'CHUNK',
    );
    expect(chunks.map((message) => new TextDecoder().decode(message.bytes))).toEqual(['bc', 'de']);
    expect(port.messages.at(-1)).toMatchObject({ type: 'END', bytesSent: 4 });
    expect(port.closed).toBe(true);
    expect(service.metrics).toEqual({
      schemaVersion: 1,
      role: 'canonical-service',
      activeReads: 0,
      peakActiveReads: 1,
      reservedBytes: 0,
      peakReservedBytes: 4,
      maxActiveReads: 64,
      maxReservedBytes: CANONICAL_RESOURCE_BROKER_MAX_RESERVED_BYTES,
    });
  });

  it('retains one early pull while the release ledger is opening', async () => {
    let releaseLedger!: () => void;
    const ledgerReady = new Promise<void>((resolve) => {
      releaseLedger = resolve;
    });
    const service = new CanonicalResourceBrokerService({
      journal: {
        async getAssetReadView() {
          await ledgerReady;
          return {
            release: {
              releaseId: transaction.releaseId,
              manifestSha256: transaction.manifestSha256,
              storageSetSha256: transaction.storageSetSha256,
              state: transaction.state,
              updatedAt: transaction.updatedAt,
            },
            mapping,
            objects: [{ sha256: digest('a'), bytes: 6, verifiedAt: 1 }],
          };
        },
      },
      store: {
        async probeRelease() {
          throw new Error('not used');
        },
        async matchObject() {
          return new Response('abcdef');
        },
      },
    });
    const port = new FakePort();
    const handling = service.handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-early',
        releaseId: 'release-a',
        sessionId: 'session-a',
        path: 'sdkjs/app.js',
      },
      port,
    );
    port.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'PULL', id: 'read-early' });
    releaseLedger();
    await handling;
    expect(port.messages.map((message) => message.type)).toEqual(['HEADERS', 'CHUNK', 'END']);
    expect(port.messages.at(-1)).toMatchObject({ bytesSent: 6 });
  });

  it('returns 416 without reading storage', async () => {
    const port = new FakePort();
    await createService().handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-416',
        releaseId: 'release-a',
        sessionId: 'session-a',
        path: 'sdkjs/app.js',
        range: { kind: 'open', start: 99 },
      },
      port,
    );
    expect(port.messages.map((message) => message.type)).toEqual(['HEADERS', 'END']);
    expect(port.messages[0]).toMatchObject({
      status: 416,
      headers: { contentLength: 0, contentRange: 'bytes */6' },
    });
  });

  it('cancels an active reader and releases its budget', async () => {
    const port = new FakePort();
    const service = createService();
    const handling = service.handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-cancel',
        releaseId: 'release-a',
        sessionId: 'session-a',
        path: 'sdkjs/app.js',
      },
      port,
    );
    await vi.waitFor(() => expect(port.messages[0]?.type).toBe('HEADERS'));
    expect(service.activeReadCount).toBe(1);
    port.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'CANCEL', id: 'read-cancel' });
    await handling;
    expect(port.messages.at(-1)).toMatchObject({ type: 'CANCELLED', bytesSent: 0 });
    expect(service.activeReadCount).toBe(0);
    expect(service.reservedReadBytes).toBe(0);
  });

  it('queues bursts behind the global read budget instead of returning busy', async () => {
    const service = createService({ maxActiveReads: 1 });
    const firstPort = new FakePort();
    const secondPort = new FakePort();
    const request = (id: string) => ({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'READ' as const,
      id,
      releaseId: 'release-a',
      sessionId: 'session-a',
      path: 'sdkjs/app.js',
    });

    const first = service.handle(request('read-first'), firstPort);
    await vi.waitFor(() => expect(firstPort.messages[0]?.type).toBe('HEADERS'));
    const second = service.handle(request('read-second'), secondPort);
    await Promise.resolve();
    expect(secondPort.messages).toEqual([]);
    expect(service.activeReadCount).toBe(1);

    firstPort.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'CANCEL', id: 'read-first' });
    await first;
    await vi.waitFor(() => expect(secondPort.messages[0]?.type).toBe('HEADERS'));
    expect(secondPort.messages.some((message) => message.type === 'ERROR' && message.code === 'busy')).toBe(false);
    expect(service.activeReadCount).toBe(1);

    secondPort.send({ protocol: RESOURCE_BROKER_PROTOCOL, type: 'CANCEL', id: 'read-second' });
    await second;
    expect(service.activeReadCount).toBe(0);
    expect(service.metrics.peakActiveReads).toBe(1);
  });

  it('fails closed for a non-complete release and unknown path', async () => {
    const nonCompletePort = new FakePort();
    await createService({ state: 'installing' }).handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-installing',
        releaseId: 'release-a',
        sessionId: 'session-a',
        path: 'sdkjs/app.js',
      },
      nonCompletePort,
    );
    expect(nonCompletePort.messages).toEqual([
      { protocol: RESOURCE_BROKER_PROTOCOL, type: 'ERROR', id: 'read-installing', code: 'release' },
    ]);

    const missingPort = new FakePort();
    await createService().handle(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'READ',
        id: 'read-missing',
        releaseId: 'release-a',
        sessionId: 'session-a',
        path: 'sdkjs/missing.js',
      },
      missingPort,
    );
    expect(missingPort.messages).toEqual([
      { protocol: RESOURCE_BROKER_PROTOCOL, type: 'ERROR', id: 'read-missing', code: 'missing' },
    ]);
  });
});
