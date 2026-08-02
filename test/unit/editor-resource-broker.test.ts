import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE,
  EDITOR_RESOURCE_BROKER_BIND_TYPE,
  EDITOR_RESOURCE_BROKER_BOUND_TYPE,
  EDITOR_RESOURCE_BROKER_UNBIND_TYPE,
  EditorResourceBrokerClient,
  parseEditorResourceBrokerBindMessage,
  parseEditorResourceBrokerUnbindMessage,
  type EditorResourceBrokerIdentity,
  type EditorResourceBrokerPort,
} from '../../src/lib/editor-resource-broker';
import {
  RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES,
  RESOURCE_BROKER_PROTOCOL,
  type ResourceBrokerClientMessage,
  type ResourceBrokerServerMessage,
} from '../../src/lib/resource-broker-protocol';
import { EditorClientIdentityRegistry } from '../../src/lib/editor-client-identity-registry';

class FakePort implements EditorResourceBrokerPort {
  onmessage: ((event: MessageEvent) => void) | null = null;
  onmessageerror: ((event: MessageEvent) => void) | null = null;
  readonly posted: Array<{ message: unknown; transfer: Transferable[] }> = [];
  readonly closeListeners = new Set<EventListener>();
  closed = false;
  started = false;

  addEventListener(type: string, listener: EventListener): void {
    if (type === 'close') this.closeListeners.add(listener);
  }

  removeEventListener(type: string, listener: EventListener): void {
    if (type === 'close') this.closeListeners.delete(listener);
  }

  close(): void {
    this.closed = true;
  }

  postMessage(message: unknown, transfer: Transferable[] = []): void {
    if (this.closed) throw new Error('port is closed');
    this.posted.push({ message, transfer });
  }

  start(): void {
    this.started = true;
  }

  emit(message: unknown): void {
    this.onmessage?.({ data: message } as MessageEvent);
  }

  emitError(): void {
    this.onmessageerror?.({ data: null } as MessageEvent);
  }

  emitClose(): void {
    for (const listener of this.closeListeners) listener(new Event('close'));
  }
}

const identityA: EditorResourceBrokerIdentity = {
  releaseId: 'release-a',
  sessionId: 'session-a',
};
const identityB: EditorResourceBrokerIdentity = {
  releaseId: 'release-b',
  sessionId: 'session-b',
};

const bindMessage = (identity: EditorResourceBrokerIdentity = identityA) => ({
  protocol: RESOURCE_BROKER_PROTOCOL,
  type: EDITOR_RESOURCE_BROKER_BIND_TYPE,
  ...identity,
});

const unbindMessage = (identity: EditorResourceBrokerIdentity = identityA) => ({
  protocol: RESOURCE_BROKER_PROTOCOL,
  type: EDITOR_RESOURCE_BROKER_UNBIND_TYPE,
  ...identity,
});

function bind(
  client: EditorResourceBrokerClient,
  identity: EditorResourceBrokerIdentity = identityA,
): { broker: FakePort; reply: FakePort } {
  const broker = new FakePort();
  const reply = new FakePort();
  expect(client.handleBindMessage(bindMessage(identity), [broker, reply])).toBe(true);
  expect(reply.posted).toEqual([
    {
      message: {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: EDITOR_RESOURCE_BROKER_BOUND_TYPE,
        ok: true,
        ...identity,
      },
      transfer: [],
    },
  ]);
  return { broker, reply };
}

function postedMessages(port: FakePort): ResourceBrokerClientMessage[] {
  return port.posted.map((entry) => entry.message as ResourceBrokerClientMessage);
}

function currentRead(port: FakePort): Extract<ResourceBrokerClientMessage, { type: 'READ' }> {
  const message = postedMessages(port).find(
    (candidate): candidate is Extract<ResourceBrokerClientMessage, { type: 'READ' }> => candidate.type === 'READ',
  );
  if (!message) throw new Error('READ was not posted');
  return message;
}

function headers(
  id: string,
  options: {
    status?: 200 | 206 | 416;
    contentLength?: number;
    contentRange?: string | null;
    contentType?: string;
  } = {},
): ResourceBrokerServerMessage {
  return {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'HEADERS',
    id,
    status: options.status ?? 200,
    headers: {
      acceptRanges: 'bytes',
      contentLength: options.contentLength ?? 6,
      contentRange: options.contentRange ?? null,
      contentType: options.contentType ?? 'text/javascript',
    },
  };
}

function chunk(id: string, value: string): ResourceBrokerServerMessage {
  const source = new TextEncoder().encode(value);
  const bytes = new ArrayBuffer(source.byteLength);
  new Uint8Array(bytes).set(source);
  return {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'CHUNK',
    id,
    bytes,
  };
}

async function flush(): Promise<void> {
  await Promise.resolve();
  await Promise.resolve();
}

afterEach(() => {
  vi.useRealTimers();
});

describe('EditorResourceBrokerClient binding', () => {
  it('binds browser timer methods before storing them on the client', async () => {
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
      const client = new EditorResourceBrokerClient();
      const { broker } = bind(client);
      const responsePromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
      await flush();
      const read = currentRead(broker);
      broker.emit(headers(read.id, { contentLength: 0 }));
      await expect(responsePromise).resolves.toMatchObject({ status: 200 });
    } finally {
      vi.unstubAllGlobals();
    }
  });

  it('commits the host identity before acknowledging a usable Broker connection', () => {
    const client = new EditorResourceBrokerClient();
    const broker = new FakePort();
    const reply = new FakePort();
    const commit = vi.fn(() => {
      expect(reply.posted).toEqual([]);
      expect(client.connectionState).toMatchObject({ status: 'connected', identity: identityA });
    });

    expect(client.handleBindMessage(bindMessage(identityA), [broker, reply], commit)).toBe(true);
    expect(commit).toHaveBeenCalledWith(identityA);
    expect(reply.posted).toHaveLength(1);
  });

  it('strictly parses binding identity and requires exactly one broker and one reply port', () => {
    expect(parseEditorResourceBrokerBindMessage(bindMessage())).toEqual(bindMessage());
    expect(parseEditorResourceBrokerBindMessage({ ...bindMessage(), extra: true })).toBeNull();
    expect(parseEditorResourceBrokerBindMessage({ ...bindMessage(), protocol: 'old' })).toBeNull();
    expect(parseEditorResourceBrokerUnbindMessage(unbindMessage())).toEqual(unbindMessage());
    expect(parseEditorResourceBrokerUnbindMessage({ ...unbindMessage(), extra: true })).toBeNull();
    expect(parseEditorResourceBrokerUnbindMessage({ ...unbindMessage(), sessionId: '../session' })).toBeNull();

    const client = new EditorResourceBrokerClient();
    const broker = new FakePort();
    const reply = new FakePort();
    const extra = new FakePort();
    expect(client.handleBindMessage(bindMessage(), [broker, reply, extra])).toBe(false);
    expect(broker.closed).toBe(true);
    expect(reply.closed).toBe(true);
    expect(extra.closed).toBe(true);

    const rejectedBroker = new FakePort();
    const rejectedReply = new FakePort();
    expect(client.bind({ ...bindMessage(), releaseId: '../release' }, rejectedBroker, rejectedReply)).toBe(false);
    expect(rejectedReply.posted[0]?.message).toEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE,
      ok: false,
      code: 'identity',
    });
  });

  it('pins a connection identity and cancels every active request when it is replaced', async () => {
    const client = new EditorResourceBrokerClient();
    const first = bind(client);
    const pending = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    await flush();
    const read = currentRead(first.broker);

    const second = bind(client, identityB);
    await expect(pending).rejects.toMatchObject({ code: 'replaced', stage: 'connection' });
    expect(postedMessages(first.broker)).toContainEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCEL',
      id: read.id,
    });
    expect(first.broker.closed).toBe(true);
    expect(client.connectionState).toEqual({
      status: 'connected',
      identity: identityB,
      generation: 2,
    });
    expect(second.broker.closed).toBe(false);
    expect(client.activeRequestCount).toBe(0);
    expect(client.reservedBytes).toBe(0);
  });

  it('keeps the A channel and stream alive when a live A Host makes a B bind ineligible', async () => {
    const registry = new EditorClientIdentityRegistry();
    const client = new EditorResourceBrokerClient();
    const first = bind(client, identityA);
    registry.bindHost('host-a', identityA);

    const responsePromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js', {
      identity: identityA,
    });
    await flush();
    const read = currentRead(first.broker);
    first.broker.emit(headers(read.id));
    const response = await responsePromise;

    const bEligible = registry.canBindHost('host-b', identityB, [
      { clientId: 'host-a', identity: identityA },
      { clientId: 'host-b', identity: identityB },
    ]);
    expect(bEligible).toBe(false);
    expect(client.connectionState).toMatchObject({ status: 'connected', identity: identityA });
    expect(first.broker.closed).toBe(false);

    const bodyPromise = response.text();
    await flush();
    first.broker.emit(chunk(read.id, 'abc'));
    await flush();
    first.broker.emit(chunk(read.id, 'def'));
    first.broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: read.id,
      bytesSent: 6,
    });
    await expect(bodyPromise).resolves.toBe('abcdef');

    const subsequent = client.fetchAsset(
      new Request('https://aries.getpi.work/sdkjs/next.js', { method: 'HEAD' }),
      'sdkjs/next.js',
      { identity: identityA },
    );
    await flush();
    const subsequentRead = postedMessages(first.broker)
      .filter((message): message is Extract<ResourceBrokerClientMessage, { type: 'READ' }> => message.type === 'READ')
      .at(-1)!;
    first.broker.emit(headers(subsequentRead.id));
    await expect(subsequent).resolves.toMatchObject({ status: 200 });
    first.broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCELLED',
      id: subsequentRead.id,
      bytesSent: 0,
    });

    expect(registry.canBindHost('host-b', identityB, [{ clientId: 'host-b', identity: identityB }])).toBe(true);
    const second = bind(client, identityB);
    registry.bindHost('host-b', identityB);
    expect(first.broker.closed).toBe(true);
    expect(second.broker.closed).toBe(false);
    expect(client.connectionState).toMatchObject({ status: 'connected', identity: identityB });
  });
});

describe('EditorResourceBrokerClient streaming fetch', () => {
  it('streams one bounded chunk per consumer pull without aggregating the asset', async () => {
    const client = new EditorResourceBrokerClient();
    const { broker } = bind(client);
    const responsePromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js', {
      windowBytes: 3,
      identity: identityA,
    });
    await flush();
    const read = currentRead(broker);
    expect(read).toMatchObject({
      releaseId: identityA.releaseId,
      sessionId: identityA.sessionId,
      path: 'sdkjs/app.js',
      windowBytes: 3,
    });

    broker.emit(headers(read.id));
    const response = await responsePromise;
    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toBe('text/javascript');
    expect(response.headers.get('content-length')).toBe('6');
    expect(response.headers.get('accept-ranges')).toBe('bytes');

    const reader = response.body!.getReader();
    const firstRead = reader.read();
    await flush();
    expect(postedMessages(broker).filter((message) => message.type === 'PULL')).toHaveLength(1);
    broker.emit(chunk(read.id, 'abc'));
    expect(new TextDecoder().decode((await firstRead).value)).toBe('abc');

    const secondRead = reader.read();
    await flush();
    expect(postedMessages(broker).filter((message) => message.type === 'PULL')).toHaveLength(2);
    broker.emit(chunk(read.id, 'def'));
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: read.id,
      bytesSent: 6,
    });
    expect(new TextDecoder().decode((await secondRead).value)).toBe('def');
    await expect(reader.read()).resolves.toEqual({ done: true, value: undefined });
    expect(client.activeRequestCount).toBe(0);
    expect(client.reservedBytes).toBe(0);
  });

  it('preserves exact Range response semantics and maps an invalid Range to 416', async () => {
    const client = new EditorResourceBrokerClient();
    const { broker } = bind(client);
    const rangePromise = client.fetchAsset(
      new Request('https://aries.getpi.work/sdkjs/app.js', {
        headers: { Range: 'bytes=1-3' },
      }),
      'sdkjs/app.js',
    );
    await flush();
    const rangeRead = currentRead(broker);
    expect(rangeRead.range).toEqual({ kind: 'closed', start: 1, end: 3 });
    broker.emit(
      headers(rangeRead.id, {
        status: 206,
        contentLength: 3,
        contentRange: 'bytes 1-3/6',
      }),
    );
    const rangeResponse = await rangePromise;
    expect(rangeResponse.status).toBe(206);
    expect(rangeResponse.headers.get('content-range')).toBe('bytes 1-3/6');
    const rangeReader = rangeResponse.body!.getReader();
    const rangeBody = rangeReader.read();
    await flush();
    broker.emit(chunk(rangeRead.id, 'bcd'));
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: rangeRead.id,
      bytesSent: 3,
    });
    expect(new TextDecoder().decode((await rangeBody).value)).toBe('bcd');

    const invalidPromise = client.fetchAsset(
      new Request('https://aries.getpi.work/sdkjs/app.js', {
        headers: { Range: 'bytes=1-2,4-5' },
      }),
      'sdkjs/app.js',
    );
    await flush();
    const invalidRead = postedMessages(broker)
      .filter((message): message is Extract<ResourceBrokerClientMessage, { type: 'READ' }> => message.type === 'READ')
      .at(-1)!;
    expect(invalidRead.range).toEqual({ kind: 'suffix', bytes: 0 });
    broker.emit(
      headers(invalidRead.id, {
        status: 416,
        contentLength: 0,
        contentRange: 'bytes */6',
      }),
    );
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: invalidRead.id,
      bytesSent: 0,
    });
    const invalidResponse = await invalidPromise;
    expect(invalidResponse.status).toBe(416);
    expect(invalidResponse.body).toBeNull();
    expect(invalidResponse.headers.get('content-range')).toBe('bytes */6');
  });

  it('handles HEAD without reading a body and propagates consumer cancellation', async () => {
    const client = new EditorResourceBrokerClient();
    const { broker } = bind(client);
    const headPromise = client.fetchAsset(
      new Request('https://aries.getpi.work/sdkjs/app.js', { method: 'HEAD' }),
      'sdkjs/app.js',
    );
    await flush();
    const headRead = currentRead(broker);
    broker.emit(headers(headRead.id));
    const headResponse = await headPromise;
    expect(headResponse.body).toBeNull();
    expect(postedMessages(broker)).toContainEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCEL',
      id: headRead.id,
    });
    expect(postedMessages(broker).some((message) => message.type === 'PULL')).toBe(false);
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCELLED',
      id: headRead.id,
      bytesSent: 0,
    });

    const bodyPromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    await flush();
    const bodyRead = postedMessages(broker)
      .filter((message): message is Extract<ResourceBrokerClientMessage, { type: 'READ' }> => message.type === 'READ')
      .at(-1)!;
    broker.emit(headers(bodyRead.id));
    const bodyResponse = await bodyPromise;
    await bodyResponse.body!.cancel();
    expect(postedMessages(broker)).toContainEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCEL',
      id: bodyRead.id,
    });
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCELLED',
      id: bodyRead.id,
      bytesSent: 0,
    });
    expect(client.activeRequestCount).toBe(0);
  });

  it('surfaces structured broker failures and rejects malformed or cross-request responses', async () => {
    const client = new EditorResourceBrokerClient();
    const { broker } = bind(client);
    const missingPromise = client.fetchAsset(new Request('https://aries.getpi.work/missing.js'), 'missing.js');
    await flush();
    const missingRead = currentRead(broker);
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'ERROR',
      id: missingRead.id,
      code: 'missing',
    });
    await expect(missingPromise).rejects.toMatchObject({
      name: 'EditorResourceBrokerError',
      code: 'missing',
      stage: 'response',
    });

    const malformedPromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    await flush();
    broker.emit(headers('another-request'));
    await expect(malformedPromise).rejects.toMatchObject({ code: 'protocol', stage: 'response' });
    expect(client.connectionState.status).toBe('disconnected');
    expect(broker.closed).toBe(true);
  });

  it('enforces the active request and reserved-byte budgets before posting READ', async () => {
    const client = new EditorResourceBrokerClient({
      maxActiveRequests: 1,
      maxReservedBytes: 1024 * 1024,
    });
    const { broker } = bind(client);
    const first = client.fetchAsset(new Request('https://aries.getpi.work/a.js'), 'a.js');
    await flush();
    expect(client.reservedBytes).toBe(RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES);
    await expect(client.fetchAsset(new Request('https://aries.getpi.work/b.js'), 'b.js')).rejects.toMatchObject({
      code: 'busy',
      stage: 'request',
    });
    expect(postedMessages(broker).filter((message) => message.type === 'READ')).toHaveLength(1);

    const read = currentRead(broker);
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'ERROR',
      id: read.id,
      code: 'cancelled',
    });
    await expect(first).rejects.toMatchObject({ code: 'cancelled' });
  });
});

describe('EditorResourceBrokerClient recovery and timeouts', () => {
  it('requests a connection, rebinds after a port closes, and never falls back to network', async () => {
    let requestCount = 0;
    const client = new EditorResourceBrokerClient({
      requestConnection() {
        requestCount += 1;
      },
    });
    const firstFetch = client.fetchAsset(
      new Request('https://aries.getpi.work/sdkjs/app.js', { method: 'HEAD' }),
      'sdkjs/app.js',
      { identity: identityA },
    );
    await flush();
    expect(requestCount).toBe(1);
    expect(client.connectionState.status).toBe('connecting');

    const first = bind(client);
    await flush();
    const firstRead = currentRead(first.broker);
    first.broker.emit(headers(firstRead.id));
    await expect(firstFetch).resolves.toMatchObject({ status: 200 });
    first.broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCELLED',
      id: firstRead.id,
      bytesSent: 0,
    });
    first.broker.emitClose();
    expect(client.connectionState.status).toBe('disconnected');

    const secondFetch = client.fetchAsset(
      new Request('https://aries.getpi.work/sdkjs/app.js', { method: 'HEAD' }),
      'sdkjs/app.js',
      { identity: identityA },
    );
    await flush();
    expect(requestCount).toBe(2);
    const second = bind(client);
    await flush();
    const secondRead = currentRead(second.broker);
    second.broker.emit(headers(secondRead.id));
    await expect(secondFetch).resolves.toMatchObject({ status: 200 });
    second.broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CANCELLED',
      id: secondRead.id,
      bytesSent: 0,
    });
  });

  it('rejects a cross-release request before posting READ', async () => {
    const client = new EditorResourceBrokerClient();
    const { broker } = bind(client, identityA);
    await expect(
      client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js', { identity: identityB }),
    ).rejects.toMatchObject({ code: 'identity', stage: 'connection' });
    expect(postedMessages(broker)).toEqual([]);
  });

  it('fails with a structured timeout, closes the stale channel, and can request a fresh binding', async () => {
    vi.useFakeTimers();
    let requestCount = 0;
    const client = new EditorResourceBrokerClient({
      requestTimeoutMs: 20,
      requestConnection() {
        requestCount += 1;
      },
    });
    const first = bind(client);
    const response = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    const timeoutExpectation = expect(response).rejects.toMatchObject({ code: 'timeout', stage: 'request' });
    await flush();
    await vi.advanceTimersByTimeAsync(21);
    await timeoutExpectation;
    expect(first.broker.closed).toBe(true);
    expect(client.connectionState.status).toBe('disconnected');

    const reconnect = client.awaitConnection(100, identityA);
    await flush();
    expect(requestCount).toBe(1);
    bind(client);
    await expect(reconnect).resolves.toEqual(identityA);
  });

  it('treats the request timeout as sliding inactivity while a stream keeps making progress', async () => {
    vi.useFakeTimers();
    const client = new EditorResourceBrokerClient({ requestTimeoutMs: 20 });
    const { broker } = bind(client);
    const responsePromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    await flush();
    const read = currentRead(broker);
    broker.emit(headers(read.id, { contentLength: 3 }));
    const response = await responsePromise;
    const body = response.text();

    for (const value of ['a', 'b', 'c']) {
      await vi.advanceTimersByTimeAsync(15);
      broker.emit(chunk(read.id, value));
      await flush();
    }
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: read.id,
      bytesSent: 3,
    });

    await expect(body).resolves.toBe('abc');
    expect(client.connectionState.status).toBe('connected');
    expect(client.activeRequestCount).toBe(0);
  });

  it('treats the request timeout as sliding inactivity while a stream keeps making progress', async () => {
    vi.useFakeTimers();
    const client = new EditorResourceBrokerClient({ requestTimeoutMs: 20 });
    const { broker } = bind(client);
    const responsePromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    await flush();
    const read = currentRead(broker);
    broker.emit(headers(read.id, { contentLength: 3 }));
    const response = await responsePromise;
    const body = response.text();

    for (const value of ['a', 'b', 'c']) {
      await vi.advanceTimersByTimeAsync(15);
      broker.emit(chunk(read.id, value));
      await flush();
    }
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: read.id,
      bytesSent: 3,
    });

    await expect(body).resolves.toBe('abc');
    expect(client.connectionState.status).toBe('connected');
    expect(client.activeRequestCount).toBe(0);
  });

  it('treats the request timeout as sliding inactivity while a stream keeps making progress', async () => {
    vi.useFakeTimers();
    const client = new EditorResourceBrokerClient({ requestTimeoutMs: 20 });
    const { broker } = bind(client);
    const responsePromise = client.fetchAsset(new Request('https://aries.getpi.work/sdkjs/app.js'), 'sdkjs/app.js');
    await flush();
    const read = currentRead(broker);
    broker.emit(headers(read.id, { contentLength: 3 }));
    const response = await responsePromise;
    const body = response.text();

    for (const value of ['a', 'b', 'c']) {
      await vi.advanceTimersByTimeAsync(15);
      broker.emit(chunk(read.id, value));
      await flush();
    }
    broker.emit({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'END',
      id: read.id,
      bytesSent: 3,
    });

    await expect(body).resolves.toBe('abc');
    expect(client.connectionState.status).toBe('connected');
    expect(client.activeRequestCount).toBe(0);
  });

  it('bounds connection recovery to the requested timeout', async () => {
    vi.useFakeTimers();
    const requestConnection = vi.fn();
    const client = new EditorResourceBrokerClient({ requestConnection });
    const waiting = client.awaitConnection(20, identityA);
    const timeoutExpectation = expect(waiting).rejects.toMatchObject({
      name: 'EditorResourceBrokerError',
      code: 'timeout',
      stage: 'connection',
    });
    await flush();
    expect(requestConnection).toHaveBeenCalledOnce();
    await vi.advanceTimersByTimeAsync(21);
    await timeoutExpectation;
    expect(client.connectionState.status).toBe('disconnected');
  });
});
