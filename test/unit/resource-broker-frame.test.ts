import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  OneTimeResourceBrokerCapabilityRegistry,
  RESOURCE_BROKER_PROTOCOL,
  type ResourceBrokerCapabilityClaim,
  type ResourceBrokerReadMessage,
  type ResourceBrokerServerMessage,
} from '../../src/lib/resource-broker-protocol';
import {
  RESOURCE_BROKER_CANONICAL_ORIGIN,
  ResourceBrokerFrameClient,
  createResourceBrokerFrameUrl,
  parseResourceBrokerFrameConnectedMessage,
  resolveResourceBrokerCanonicalOrigin,
  resolveResourceBrokerPhysicalEditorIdentity,
  resolveResourceBrokerTrustedParentIdentity,
  validateResourceBrokerFrameChallenge,
  validateResourceBrokerFrameChallengeEvent,
  type ResourceBrokerFrameExpectedIdentity,
} from '../../src/lib/resource-broker-frame-client';
import { MemoryReleaseLeaseLedger } from '../../src/lib/release-lease-gc';
import {
  ResourceBrokerFrameCapabilityGate,
  ResourceBrokerRelaySession,
  createResourceBrokerReleaseLease,
  resolveResourceBrokerFrameIdentity,
  type ResourceBrokerFrameIdentity,
  type ResourceBrokerPortLike,
  type ResourceBrokerWorkerPortOpener,
} from '../../src/resource-broker';

const expectedIdentity: ResourceBrokerFrameExpectedIdentity = {
  canonicalOrigin: RESOURCE_BROKER_CANONICAL_ORIGIN,
  parentOrigin: 'https://piwork.getpi.work',
  editorOrigin: 'https://aries.getpi.work',
  releaseId: 'onlyoffice-browser-0.6.0+release.1',
  sessionId: 'office-session_1',
};

const frameIdentity: ResourceBrokerFrameIdentity = {
  ...expectedIdentity,
  physicalEditorOrigin: 'https://aries.getpi.work',
  localMatrix: false,
};

class FakePort implements ResourceBrokerPortLike {
  onmessage: ((event: MessageEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  readonly posted: Array<{ message: unknown; transfer: Transferable[] }> = [];
  closed = false;
  started = false;

  close(): void {
    this.closed = true;
  }

  postMessage(message: unknown, transfer: Transferable[] = []): void {
    if (this.closed) throw new Error('closed');
    this.posted.push({ message, transfer });
  }

  start(): void {
    this.started = true;
  }

  emit(message: unknown): void {
    this.onmessage?.({ data: message } as MessageEvent);
  }
}

function readMessage(overrides: Partial<ResourceBrokerReadMessage> = {}): ResourceBrokerReadMessage {
  return {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'READ',
    id: 'read-1',
    releaseId: expectedIdentity.releaseId,
    sessionId: expectedIdentity.sessionId,
    path: 'blobs/sha256/aa',
    ...overrides,
  };
}

afterEach(() => {
  vi.useRealTimers();
});

describe('Resource Broker frame identity', () => {
  it('accepts only the canonical production origin, a constellation parent, and an allowed nested parent', () => {
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'https://onlyoffice.getpi.work',
        documentReferrer: 'https://aries.getpi.work/r/release/office-host.html',
        search:
          '?releaseId=onlyoffice-browser-0.6.0%2Brelease.1&sessionId=office-session_1&parentOrigin=https%3A%2F%2Fpiwork.getpi.work',
      }),
    ).toEqual(frameIdentity);

    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'https://evil.example',
        documentReferrer: 'https://aries.getpi.work/office-host.html',
        search: '?releaseId=release&sessionId=session&parentOrigin=https%3A%2F%2Fpiwork.getpi.work',
      }),
    ).toBeNull();
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'https://onlyoffice.getpi.work',
        documentReferrer: 'https://office-editor-evil.getpi.work/office-host.html',
        search: '?releaseId=release&sessionId=session&parentOrigin=https%3A%2F%2Fpiwork.getpi.work',
      }),
    ).toBeNull();
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'https://onlyoffice.getpi.work',
        documentReferrer: 'https://aries.getpi.work/office-host.html',
        search: '?releaseId=release&sessionId=session&parentOrigin=https%3A%2F%2Fevil.example',
      }),
    ).toBeNull();
  });

  it('binds each relay-frame lifetime to a unique durable release lease', async () => {
    const ledger = new MemoryReleaseLeaseLedger({ now: () => 1_000 });
    const first = createResourceBrokerReleaseLease(frameIdentity, {
      ledger,
      randomUUID: () => '00000000-0000-4000-8000-000000000001',
    });
    const second = createResourceBrokerReleaseLease(frameIdentity, {
      ledger,
      randomUUID: () => '00000000-0000-4000-8000-000000000002',
    });

    await Promise.all([first.start(), second.start()]);
    expect(ledger.listActiveLeases()).toEqual(
      expect.arrayContaining([
        expect.objectContaining({
          leaseId: first.binding.leaseId,
          releaseId: frameIdentity.releaseId,
          sessionId: frameIdentity.sessionId,
          editorOrigin: frameIdentity.editorOrigin,
        }),
        expect.objectContaining({
          leaseId: second.binding.leaseId,
          releaseId: frameIdentity.releaseId,
          sessionId: frameIdentity.sessionId,
          editorOrigin: frameIdentity.editorOrigin,
        }),
      ]),
    );

    await first.stop();
    expect(ledger.listActiveLeases().map((lease) => lease.leaseId)).toEqual([second.binding.leaseId]);
    await second.stop();
  });

  it('requires an explicit, same-port localhost matrix branch and maps it to production identities', () => {
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://taurus.localhost:8787/office-host.html',
        search:
          '?releaseId=release-a&sessionId=session-a&parentOrigin=https%3A%2F%2Fonlyoffice.getpi.work&localMatrix=1',
      }),
    ).toEqual({
      canonicalOrigin: 'http://onlyoffice.localhost:8787',
      physicalEditorOrigin: 'http://taurus.localhost:8787',
      editorOrigin: 'https://taurus.getpi.work',
      parentOrigin: 'https://onlyoffice.getpi.work',
      releaseId: 'release-a',
      sessionId: 'session-a',
      localMatrix: true,
    });
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://taurus.localhost:8787/office-host.html',
        search: '?releaseId=release-a&sessionId=session-a&parentOrigin=https%3A%2F%2Fonlyoffice.getpi.work',
      }),
    ).toBeNull();
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'http://onlyoffice.localhost:8787',
        documentReferrer: 'http://taurus.localhost:9999/office-host.html',
        search:
          '?releaseId=release-a&sessionId=session-a&parentOrigin=https%3A%2F%2Fonlyoffice.getpi.work&localMatrix=1',
      }),
    ).toBeNull();
  });

  it('rejects ambiguous query keys and non-canonical origin spellings', () => {
    expect(
      resolveResourceBrokerFrameIdentity({
        locationOrigin: 'https://onlyoffice.getpi.work',
        documentReferrer: 'https://aries.getpi.work/office-host.html',
        search: '?releaseId=a&releaseId=b&sessionId=session&parentOrigin=https%3A%2F%2Fpiwork.getpi.work',
      }),
    ).toBeNull();
    expect(resolveResourceBrokerCanonicalOrigin('https://ONLYOFFICE.getpi.work')).toBeNull();
    expect(
      resolveResourceBrokerPhysicalEditorIdentity('https://aries.getpi.work:444', RESOURCE_BROKER_CANONICAL_ORIGIN),
    ).toBeNull();
    expect(
      resolveResourceBrokerTrustedParentIdentity(
        'https://piwork.getpi.work.evil.example',
        RESOURCE_BROKER_CANONICAL_ORIGIN,
      ),
    ).toBeNull();
  });

  it('bounds active requests, reserved transfer windows, and queued flow control before worker startup', () => {
    const forwarded: ResourceBrokerServerMessage[] = [];
    const relay = new ResourceBrokerRelaySession({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      maxActiveRequests: 2,
      maxReservedBytes: 1024 * 1024,
      openWorkerPort: () => new Promise(() => undefined),
      postToClient(message) {
        forwarded.push(message);
      },
    });

    expect(relay.handleClientMessage(readMessage({ id: 'read-1' }))).toBe(true);
    expect(relay.handleClientMessage(readMessage({ id: 'read-2' }))).toBe(true);
    expect(relay.activeRequestCount).toBe(2);
    expect(relay.reservedBytes).toBe(512 * 1024);

    expect(relay.handleClientMessage(readMessage({ id: 'read-3' }))).toBe(true);
    expect(forwarded.at(-1)).toMatchObject({ type: 'ERROR', id: 'read-3', code: 'busy' });
    expect(relay.activeRequestCount).toBe(2);

    expect(
      relay.handleClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PULL',
        id: 'read-1',
      }),
    ).toBe(true);
    expect(
      relay.handleClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PULL',
        id: 'read-1',
      }),
    ).toBe(true);
    expect(forwarded.at(-1)).toMatchObject({ type: 'ERROR', id: 'read-1', code: 'busy' });
    expect(relay.activeRequestCount).toBe(1);
    expect(relay.reservedBytes).toBe(256 * 1024);

    relay.destroy();
    expect(relay.reservedBytes).toBe(0);
  });

  it('keeps the canonical relay frame alive after transferring the broker port', async () => {
    const messageListeners = new Set<(event: MessageEvent) => void>();
    const remove = vi.fn();
    const now = Date.now();
    const contentWindow = {
      postMessage(message: unknown, targetOrigin: string, transfer: Transferable[] = []) {
        expect(targetOrigin).toBe('http://onlyoffice.localhost:8787');
        expect(message).toMatchObject({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CONNECT',
        });
        const port = transfer[0] as MessagePort;
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CONNECTED',
          releaseId: expectedIdentity.releaseId,
          sessionId: expectedIdentity.sessionId,
        });
      },
    } as unknown as Window;
    const frame = {
      hidden: false,
      tabIndex: 0,
      referrerPolicy: '',
      src: '',
      contentWindow,
      remove,
      setAttribute: vi.fn(),
    } as unknown as HTMLIFrameElement;
    const windowLike = {
      location: { origin: 'http://host-aries.office.localhost:8787' },
      setTimeout,
      clearTimeout,
      addEventListener(type: string, listener: EventListener) {
        if (type === 'message') messageListeners.add(listener as (event: MessageEvent) => void);
      },
      removeEventListener(type: string, listener: EventListener) {
        if (type === 'message') messageListeners.delete(listener as (event: MessageEvent) => void);
      },
    } as unknown as Window;
    const documentLike = {
      createElement: vi.fn(() => frame),
      body: {
        append: vi.fn(() => {
          queueMicrotask(() => {
            for (const listener of messageListeners) {
              listener({
                source: contentWindow,
                origin: 'http://onlyoffice.localhost:8787',
                data: {
                  protocol: RESOURCE_BROKER_PROTOCOL,
                  type: 'CHALLENGE',
                  capability: {
                    token: 'ab'.repeat(32),
                    parentOrigin: 'https://onlyoffice.getpi.work',
                    editorOrigin: 'https://aries.getpi.work',
                    releaseId: expectedIdentity.releaseId,
                    sessionId: expectedIdentity.sessionId,
                    issuedAtMs: now,
                    expiresAtMs: now + 30_000,
                  },
                },
              } as MessageEvent);
            }
          });
        }),
      },
    } as unknown as Document;
    const client = new ResourceBrokerFrameClient({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      parentOrigin: 'http://onlyoffice.localhost:8787',
      canonicalOrigin: 'http://onlyoffice.localhost:8787',
      allowLocalMatrix: true,
      window: windowLike,
      document: documentLike,
    });

    await client.connect();
    const port = client.takeConnectionPort();
    expect(remove).not.toHaveBeenCalled();
    client.destroy();
    expect(remove).toHaveBeenCalledOnce();
    port.close();
  });
});

describe('Resource Broker capability handshake', () => {
  it('uses a short-lived one-time capability and rejects replay or cross-release claims', () => {
    const registry = new OneTimeResourceBrokerCapabilityRegistry({
      now: () => 1_000,
      randomFill(bytes) {
        bytes.fill(0xab);
        return bytes;
      },
    });
    const gate = new ResourceBrokerFrameCapabilityGate(frameIdentity, registry);
    const capability = gate.issue();
    const claim: ResourceBrokerCapabilityClaim = {
      token: capability.token,
      parentOrigin: capability.parentOrigin,
      editorOrigin: capability.editorOrigin,
      releaseId: capability.releaseId,
      sessionId: capability.sessionId,
    };
    const connect = { protocol: RESOURCE_BROKER_PROTOCOL, type: 'CONNECT', capability: claim };

    expect(gate.consume(connect, frameIdentity.physicalEditorOrigin, 1_001).ok).toBe(true);
    expect(gate.consume(connect, frameIdentity.physicalEditorOrigin, 1_002)).toEqual({
      ok: false,
      code: 'mismatch',
    });

    const mismatchGate = new ResourceBrokerFrameCapabilityGate(frameIdentity, registry);
    const mismatchCapability = mismatchGate.issue();
    expect(
      mismatchGate.consume(
        {
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CONNECT',
          capability: {
            token: mismatchCapability.token,
            parentOrigin: mismatchCapability.parentOrigin,
            editorOrigin: mismatchCapability.editorOrigin,
            releaseId: 'other-release',
            sessionId: mismatchCapability.sessionId,
          },
        },
        frameIdentity.physicalEditorOrigin,
        1_001,
      ),
    ).toEqual({ ok: false, code: 'mismatch' });
  });

  it('validates challenge source identity fields and expiry before connecting', () => {
    const message = {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHALLENGE',
      capability: {
        token: 'ab'.repeat(32),
        parentOrigin: expectedIdentity.parentOrigin,
        editorOrigin: expectedIdentity.editorOrigin,
        releaseId: expectedIdentity.releaseId,
        sessionId: expectedIdentity.sessionId,
        issuedAtMs: 1_000,
        expiresAtMs: 31_000,
      },
    };
    expect(validateResourceBrokerFrameChallenge(message, expectedIdentity, 2_000)).toEqual(message);
    const source = {} as Window;
    expect(
      validateResourceBrokerFrameChallengeEvent(
        { data: message, origin: expectedIdentity.canonicalOrigin, source },
        source,
        expectedIdentity,
        2_000,
      ),
    ).toEqual(message);
    expect(
      validateResourceBrokerFrameChallengeEvent(
        { data: message, origin: 'https://evil.example', source },
        source,
        expectedIdentity,
        2_000,
      ),
    ).toBeNull();
    expect(
      validateResourceBrokerFrameChallengeEvent(
        { data: message, origin: expectedIdentity.canonicalOrigin, source: {} as Window },
        source,
        expectedIdentity,
        2_000,
      ),
    ).toBeNull();
    expect(
      validateResourceBrokerFrameChallenge(
        {
          ...message,
          capability: { ...message.capability, releaseId: 'other-release' },
        },
        expectedIdentity,
        2_000,
      ),
    ).toBeNull();
    expect(validateResourceBrokerFrameChallenge(message, expectedIdentity, 31_000)).toBeNull();
  });

  it('uses exact target identity in the frame URL and strictly parses connection responses', () => {
    expect(createResourceBrokerFrameUrl(expectedIdentity).href).toBe(
      'https://onlyoffice.getpi.work/resource-broker.html?releaseId=onlyoffice-browser-0.6.0%2Brelease.1&sessionId=office-session_1&parentOrigin=https%3A%2F%2Fpiwork.getpi.work',
    );
    expect(
      parseResourceBrokerFrameConnectedMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CONNECTED',
        releaseId: expectedIdentity.releaseId,
        sessionId: expectedIdentity.sessionId,
      }),
    ).not.toBeNull();
    expect(
      parseResourceBrokerFrameConnectedMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CONNECTED',
        releaseId: expectedIdentity.releaseId,
        sessionId: expectedIdentity.sessionId,
        injected: true,
      }),
    ).toBeNull();
  });
});

describe('Resource Broker request relay', () => {
  it('queues an early PULL until the dedicated worker port is ready and transfers strict chunks', async () => {
    let resolvePort!: (port: ResourceBrokerPortLike) => void;
    const workerPortPromise = new Promise<ResourceBrokerPortLike>((resolve) => {
      resolvePort = resolve;
    });
    const workerPort = new FakePort();
    const forwarded: Array<{ message: ResourceBrokerServerMessage; transfer: Transferable[] }> = [];
    const relay = new ResourceBrokerRelaySession({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      openWorkerPort: () => workerPortPromise,
      postToClient(message, transfer = []) {
        forwarded.push({ message, transfer });
      },
    });

    expect(relay.handleClientMessage(readMessage())).toBe(true);
    expect(
      relay.handleClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PULL',
        id: 'read-1',
      }),
    ).toBe(true);
    await Promise.resolve();
    resolvePort(workerPort);
    await vi.waitFor(() => expect(workerPort.started).toBe(true));
    expect(workerPort.posted[0]?.message).toEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'PULL',
      id: 'read-1',
    });

    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'HEADERS',
      id: 'read-1',
      status: 200,
      headers: {
        acceptRanges: 'bytes',
        contentLength: 3,
        contentRange: null,
        contentType: 'application/octet-stream',
      },
    });
    const bytes = new Uint8Array([1, 2, 3]).buffer;
    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHUNK',
      id: 'read-1',
      bytes,
    });
    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: 'read-1',
      bytesSent: 3,
    });

    expect(forwarded.map(({ message }) => message.type)).toEqual(['HEADERS', 'CHUNK', 'END']);
    expect(forwarded[1]?.transfer).toEqual([bytes]);
    expect(relay.activeRequestCount).toBe(0);
    expect(workerPort.closed).toBe(true);
  });

  it('suppresses worker messages already in flight after an editor cancellation', async () => {
    const workerPort = new FakePort();
    const forwarded: ResourceBrokerServerMessage[] = [];
    const relay = new ResourceBrokerRelaySession({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      openWorkerPort: async () => workerPort,
      postToClient(message) {
        forwarded.push(message);
      },
    });
    relay.handleClientMessage(readMessage());
    await vi.waitFor(() => expect(workerPort.started).toBe(true));
    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'HEADERS',
      id: 'read-1',
      status: 200,
      headers: {
        acceptRanges: 'bytes',
        contentLength: 6,
        contentRange: null,
        contentType: 'application/octet-stream',
      },
    });
    const first = new Uint8Array([1, 2]).buffer;
    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHUNK',
      id: 'read-1',
      bytes: first,
    });
    expect(
      relay.handleClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CANCEL',
        id: 'read-1',
      }),
    ).toBe(true);
    const late = new Uint8Array([3, 4]).buffer;
    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHUNK',
      id: 'read-1',
      bytes: late,
    });
    workerPort.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCELLED',
      id: 'read-1',
      bytesSent: 4,
    });

    expect(forwarded.map((message) => message.type)).toEqual(['HEADERS', 'CHUNK', 'CANCELLED']);
    expect(forwarded.at(-1)).toMatchObject({ type: 'CANCELLED', bytesSent: 2 });
    expect(relay.activeRequestCount).toBe(0);
    expect(relay.reservedBytes).toBe(0);
  });

  it('rejects cross-release requests without opening a worker channel', () => {
    const openWorkerPort = vi.fn();
    const forwarded: ResourceBrokerServerMessage[] = [];
    const relay = new ResourceBrokerRelaySession({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      openWorkerPort,
      postToClient(message) {
        forwarded.push(message);
      },
    });
    relay.handleClientMessage(readMessage({ releaseId: 'other-release' }));
    expect(openWorkerPort).not.toHaveBeenCalled();
    expect(forwarded).toEqual([
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ERROR',
        id: 'read-1',
        code: 'release',
      },
    ]);
  });

  it('fails active requests on a worker controller change and reconnects the next request', async () => {
    const ports = [new FakePort(), new FakePort()];
    const openWorkerPort = vi
      .fn<ResourceBrokerWorkerPortOpener>()
      .mockResolvedValueOnce(ports[0])
      .mockResolvedValueOnce(ports[1]);
    const forwarded: ResourceBrokerServerMessage[] = [];
    const relay = new ResourceBrokerRelaySession({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      openWorkerPort,
      postToClient(message) {
        forwarded.push(message);
      },
    });

    relay.handleClientMessage(readMessage());
    await vi.waitFor(() => expect(ports[0].started).toBe(true));
    relay.notifyControllerChange();
    expect(forwarded.at(-1)).toMatchObject({ type: 'ERROR', id: 'read-1', code: 'protocol' });

    relay.handleClientMessage(readMessage({ id: 'read-2' }));
    await vi.waitFor(() => expect(ports[1].started).toBe(true));
    expect(openWorkerPort).toHaveBeenCalledTimes(2);
    expect(relay.activeRequestCount).toBe(1);
    relay.destroy();
  });

  it('times out a request with no worker progress after the fixed inactivity window', async () => {
    vi.useFakeTimers();
    const forwarded: ResourceBrokerServerMessage[] = [];
    const relay = new ResourceBrokerRelaySession({
      releaseId: expectedIdentity.releaseId,
      sessionId: expectedIdentity.sessionId,
      requestTimeoutMs: 1_000,
      openWorkerPort: () => new Promise(() => undefined),
      postToClient(message) {
        forwarded.push(message);
      },
    });
    relay.handleClientMessage(readMessage());
    await vi.advanceTimersByTimeAsync(1_000);
    expect(forwarded).toEqual([
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ERROR',
        id: 'read-1',
        code: 'timeout',
      },
    ]);
    expect(relay.activeRequestCount).toBe(0);
  });
});
