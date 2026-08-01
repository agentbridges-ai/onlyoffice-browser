import { describe, expect, it, vi } from 'vitest';
import {
  bindCanonicalBrokerToEditorServiceWorker,
  probeCanonicalResourceBroker,
  verifyEditorServiceWorkerRead,
} from '../../src/lib/editor-sw-readiness-probe';
import { RESOURCE_BROKER_PROTOCOL } from '../../src/lib/resource-broker-protocol';

class LinkedPort {
  onmessage: ((event: MessageEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  peer: LinkedPort | null = null;
  closed = false;
  sent: unknown[] = [];

  close(): void {
    this.closed = true;
  }

  postMessage(message: unknown): void {
    this.sent.push(message);
    this.peer?.onmessage?.({ data: message } as MessageEvent);
  }

  start(): void {}
}

function linkedChannel(): MessageChannel {
  const port1 = new LinkedPort();
  const port2 = new LinkedPort();
  port1.peer = port2;
  port2.peer = port1;
  return { port1, port2 } as unknown as MessageChannel;
}

const identity = {
  releaseId: 'release-a',
  sessionId: 'session-a',
};

const probeValue = {
  releaseId: identity.releaseId,
  ready: true as const,
  probePath: 'sdkjs/word/sdk-all.js',
  probeBytes: 4096,
  probeSha256: 'a'.repeat(64),
};

describe('Editor Service Worker readiness probe', () => {
  it('gets one deterministic manifest-bound object from the canonical broker', async () => {
    const port = new LinkedPort();
    port.postMessage = vi.fn((message: unknown) => {
      const request = message as { id: string };
      port.onmessage?.({
        data: {
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'RESULT',
          id: request.id,
          value: probeValue,
        },
      } as MessageEvent);
    });

    await expect(
      probeCanonicalResourceBroker(port, identity, {
        requestId: 'probe-a',
      }),
    ).resolves.toEqual(probeValue);
    expect(port.postMessage).toHaveBeenCalledWith({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'PROBE',
      id: 'probe-a',
      ...identity,
    });
  });

  it('transfers the exact broker port and requires a matching bind acknowledgement', async () => {
    const brokerPort = new LinkedPort();
    const worker = {
      postMessage: vi.fn((message: unknown, transfer?: Transferable[]) => {
        const request = message as { releaseId: string; sessionId: string };
        const replyPort = transfer?.[1] as unknown as LinkedPort;
        replyPort.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'ONLYOFFICE_BROKER_BOUND',
          ok: true,
          releaseId: request.releaseId,
          sessionId: request.sessionId,
        });
      }),
    };

    await bindCanonicalBrokerToEditorServiceWorker(worker, brokerPort, identity, {
      createMessageChannel: linkedChannel,
    });
    expect(worker.postMessage).toHaveBeenCalledWith(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ONLYOFFICE_BIND_BROKER',
        ...identity,
      },
      expect.arrayContaining([brokerPort]),
    );
  });

  it('proves the routed Editor Service Worker path with an exact one-byte Range response', async () => {
    const fetchMock = vi.fn(async (_input: RequestInfo | URL, _init?: RequestInit) => {
      return new Response(new Uint8Array([0x61]), {
        status: 206,
        headers: {
          'Accept-Ranges': 'bytes',
          'Content-Length': '1',
          'Content-Range': `bytes 0-0/${probeValue.probeBytes}`,
          'Content-Type': 'text/javascript',
        },
      });
    });

    await expect(
      verifyEditorServiceWorkerRead(fetchMock as typeof fetch, 'https://aries.getpi.work', probeValue),
    ).resolves.toBeUndefined();
    const [requestUrl, requestInit] = fetchMock.mock.calls[0]!;
    expect(String(requestUrl)).toBe(`https://aries.getpi.work/r/${identity.releaseId}/${probeValue.probePath}`);
    expect(requestInit).toMatchObject({
      method: 'GET',
      headers: { Range: 'bytes=0-0' },
      cache: 'no-store',
      credentials: 'omit',
      redirect: 'error',
    });
  });

  it('fails closed when a direct network-style 200 response bypasses Range semantics', async () => {
    const fetchMock = vi.fn(async (_input: RequestInfo | URL, _init?: RequestInit) => {
      return new Response(new Uint8Array([0x61]), { status: 200 });
    });
    await expect(
      verifyEditorServiceWorkerRead(fetchMock as typeof fetch, 'https://aries.getpi.work', probeValue),
    ).rejects.toThrow('Range response is invalid');
  });
});
