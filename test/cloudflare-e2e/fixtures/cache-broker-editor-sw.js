const BROKER_PROXY_PATTERN = /^\/__matrix__\/broker-segments\/([a-f0-9]{64})$/;
const REQUEST_TIMEOUT_MS = 30_000;
const EDITOR_SHELL_CACHE = 'onlyoffice-editor-broker-shell-experiment-v1';
const EDITOR_SHELL_PATH = '/__matrix__/cache-broker-editor.html';
let brokerPort = null;
let brokerIdentity = null;
let nextRequestId = 0;
const pending = new Map();
const bindingWaiters = new Set();

self.addEventListener('install', (event) => {
  event.waitUntil(
    Promise.all([self.skipWaiting(), caches.open(EDITOR_SHELL_CACHE).then((cache) => cache.add(EDITOR_SHELL_PATH))]),
  );
});
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));

const failPending = (error) => {
  for (const request of pending.values()) {
    clearTimeout(request.timeout);
    request.rejectHeaders(error);
    request.controller?.error(error);
  }
  pending.clear();
};

const waitForBroker = async (clientId) => {
  if (brokerPort && brokerIdentity) return;
  const client = clientId ? await self.clients.get(clientId) : null;
  if (!client) throw new Error('broker-client-missing');
  const connection = new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      bindingWaiters.delete(waiter);
      reject(new Error('broker-connect-timeout'));
    }, 5_000);
    const waiter = {
      reject,
      resolve: () => {
        clearTimeout(timeout);
        resolve();
      },
    };
    bindingWaiters.add(waiter);
  });
  client.postMessage({ type: 'ONLYOFFICE_BROKER_NEEDED' });
  await connection;
};

const finish = (request) => {
  clearTimeout(request.timeout);
  pending.delete(request.id);
};

const handleBrokerMessage = (event) => {
  const message = event.data;
  const request = pending.get(message?.id);
  if (!request) return;
  if (message.type === 'HEADERS') {
    request.headersReceived = true;
    request.resolveHeaders(message);
    if (request.headOnly || message.status === 416) {
      brokerPort?.postMessage({ type: 'CANCEL', id: request.id });
    }
    return;
  }
  if (message.type === 'CHUNK') {
    if (request.headOnly || request.cancelled) {
      brokerPort?.postMessage({ type: 'CANCEL', id: request.id });
      return;
    }
    request.controller?.enqueue(new Uint8Array(message.bytes));
    request.resolvePull?.();
    request.pullPromise = null;
    request.resolvePull = null;
    return;
  }
  if (message.type === 'END' || message.type === 'CANCELLED') {
    request.ended = true;
    if (request.headOnly || request.status === 416) {
      finish(request);
      return;
    }
    request.resolvePull?.();
    request.controller?.close();
    finish(request);
    return;
  }
  if (message.type === 'ERROR') {
    const error = new Error(message.error);
    request.rejectHeaders(error);
    request.controller?.error(error);
    finish(request);
  }
};

self.addEventListener('message', (event) => {
  if (event.data?.type !== 'ONLYOFFICE_BIND_BROKER' || event.ports.length !== 2) return;
  const [nextBrokerPort, replyPort] = event.ports;
  if (
    event.data.protocol !== 1 ||
    typeof event.data.releaseId !== 'string' ||
    typeof event.data.sessionId !== 'string'
  ) {
    replyPort.postMessage({ ok: false, error: 'identity' });
    replyPort.close();
    nextBrokerPort.close();
    return;
  }
  if (brokerPort) {
    brokerPort.close();
    failPending(new Error('broker-replaced'));
  }
  brokerPort = nextBrokerPort;
  brokerIdentity = {
    protocol: event.data.protocol,
    releaseId: event.data.releaseId,
    sessionId: event.data.sessionId,
  };
  brokerPort.onmessage = handleBrokerMessage;
  brokerPort.start();
  for (const waiter of bindingWaiters) waiter.resolve();
  bindingWaiters.clear();
  replyPort.postMessage({ ok: true, ...brokerIdentity });
  replyPort.close();
});

const parseRange = (value) => {
  if (!value) return undefined;
  const match = /^bytes=(\d+)-(\d+)$/.exec(value);
  if (!match) return null;
  const start = Number(match[1]);
  const end = Number(match[2]);
  return Number.isSafeInteger(start) && Number.isSafeInteger(end) ? { start, end } : null;
};

const brokerResponse = async (request, sha256, clientId) => {
  try {
    await waitForBroker(clientId);
  } catch {
    return new Response('Broker unavailable', {
      status: 503,
      headers: { 'content-type': 'text/plain; charset=utf-8', 'retry-after': '1' },
    });
  }
  const range = parseRange(request.headers.get('range'));
  if (range === null) {
    return new Response(null, {
      status: 416,
      headers: { 'accept-ranges': 'bytes', 'content-range': 'bytes */*' },
    });
  }
  const id = `sw-${++nextRequestId}`;
  let resolveHeaders;
  let rejectHeaders;
  const headersPromise = new Promise((resolve, reject) => {
    resolveHeaders = resolve;
    rejectHeaders = reject;
  });
  const state = {
    cancelled: false,
    controller: null,
    ended: false,
    headOnly: request.method === 'HEAD',
    headersReceived: false,
    id,
    pullPromise: null,
    rejectHeaders,
    resolvePull: null,
    resolveHeaders,
    status: 0,
    timeout: setTimeout(() => {
      const current = pending.get(id);
      if (!current) return;
      current.cancelled = true;
      brokerPort?.postMessage({ type: 'CANCEL', id });
      const error = new Error('broker-timeout');
      brokerPort?.close();
      brokerPort = null;
      brokerIdentity = null;
      current.rejectHeaders(error);
      current.controller?.error(error);
      finish(current);
    }, REQUEST_TIMEOUT_MS),
  };
  const body = state.headOnly
    ? null
    : new ReadableStream({
        start(controller) {
          state.controller = controller;
        },
        pull() {
          if (state.ended || state.cancelled) return undefined;
          if (!state.pullPromise) {
            state.pullPromise = new Promise((resolve) => {
              state.resolvePull = resolve;
            });
            brokerPort?.postMessage({ type: 'PULL', id });
          }
          return state.pullPromise;
        },
        cancel() {
          state.cancelled = true;
          brokerPort?.postMessage({ type: 'CANCEL', id });
          finish(state);
        },
      });
  pending.set(id, state);
  brokerPort.postMessage({
    type: 'READ',
    id,
    sha256,
    releaseId: brokerIdentity.releaseId,
    ...(range ? { range } : {}),
  });
  try {
    const metadata = await headersPromise;
    state.status = metadata.status;
    const headers = new Headers({
      'accept-ranges': metadata.headers.acceptRanges,
      'content-length': String(metadata.headers.contentLength),
      'content-type': metadata.headers.contentType,
      'x-onlyoffice-broker': '1',
      'x-onlyoffice-release-id': brokerIdentity.releaseId,
    });
    if (metadata.headers.contentRange) headers.set('content-range', metadata.headers.contentRange);
    return new Response(metadata.status === 416 ? null : body, {
      status: metadata.status,
      headers,
    });
  } catch (error) {
    finish(state);
    return new Response(error instanceof Error ? error.message : String(error), {
      status: 503,
      headers: { 'content-type': 'text/plain; charset=utf-8', 'retry-after': '1' },
    });
  }
};

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (event.request.method === 'GET' && url.origin === self.location.origin && url.pathname === EDITOR_SHELL_PATH) {
    event.respondWith(
      caches
        .open(EDITOR_SHELL_CACHE)
        .then((cache) => cache.match(EDITOR_SHELL_PATH))
        .then(async (response) => {
          if (!response) return fetch(event.request);
          return new Response(await response.arrayBuffer(), {
            status: 200,
            headers: { 'content-type': 'text/html; charset=utf-8' },
          });
        })
        .catch(
          (error) =>
            new Response(`Persisted editor shell failed: ${error instanceof Error ? error.message : String(error)}`, {
              status: 503,
              headers: { 'content-type': 'text/plain; charset=utf-8' },
            }),
        ),
    );
    return;
  }
  const match = BROKER_PROXY_PATTERN.exec(url.pathname);
  if (!match || (event.request.method !== 'GET' && event.request.method !== 'HEAD')) return;
  event.respondWith(brokerResponse(event.request, match[1], event.clientId));
});
