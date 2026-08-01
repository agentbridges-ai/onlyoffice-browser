const SHELL_CACHE = 'onlyoffice-cache-broker-shell-experiment-v2';
const CONTENT_CACHE = 'onlyoffice-cloudflare-cache-broker-experiment-v3';
const MANIFEST_CACHE = 'onlyoffice-cache-broker-manifests-experiment-v2';
const SHELL_PATHS = ['/__matrix__/cache-broker.html', '/__matrix__/cache-broker.js'];
const CHUNK_BYTES = 256 * 1024;
const DIGEST_PATTERN = /^[a-f0-9]{64}$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const activeReads = new Map();

self.addEventListener('install', (event) => {
  event.waitUntil(
    Promise.all([self.skipWaiting(), caches.open(SHELL_CACHE).then((cache) => cache.addAll(SHELL_PATHS))]),
  );
});

self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin || !SHELL_PATHS.includes(url.pathname)) return;
  event.respondWith(
    caches.match(url.pathname, { cacheName: SHELL_CACHE }).then((response) => response || fetch(event.request)),
  );
});

const hex = (bytes) => Array.from(bytes, (value) => value.toString(16).padStart(2, '0')).join('');
const segmentUrl = (sha256) => new URL(`/segments/sha256/${sha256}`, self.location.origin).href;
const replyError = (port, id, error) =>
  port.postMessage({
    type: 'ERROR',
    id,
    error: error instanceof Error ? error.message : String(error),
  });

const loadSegments = async (releaseId) => {
  if (!RELEASE_ID_PATTERN.test(releaseId || '')) throw new Error('release');
  const manifestUrl = `/releases/${encodeURIComponent(releaseId)}/manifest.json`;
  const cache = await caches.open(MANIFEST_CACHE);
  let response;
  if (self.navigator.onLine) {
    try {
      response = await fetch(manifestUrl, { cache: 'no-store', credentials: 'omit' });
      if (response.ok) await cache.put(manifestUrl, response.clone());
    } catch {
      response = await cache.match(manifestUrl);
    }
  } else {
    response = await cache.match(manifestUrl);
  }
  if (!response?.ok) throw new Error('manifest');
  const manifest = await response.json();
  if (manifest?.version !== 4 || manifest.releaseId !== releaseId || !Array.isArray(manifest.package?.segments)) {
    throw new Error('manifest');
  }
  const segments = new Map();
  for (const segment of manifest.package.segments) {
    if (!DIGEST_PATTERN.test(segment?.sha256) || !Number.isSafeInteger(segment?.bytes) || segment.bytes <= 0) {
      throw new Error('manifest');
    }
    segments.set(segment.sha256, { bytes: segment.bytes, sha256: segment.sha256 });
  }
  return segments;
};

const resolveSegment = async (releaseId, sha256) => {
  if (!DIGEST_PATTERN.test(sha256 || '')) throw new Error('digest');
  const segment = (await loadSegments(releaseId)).get(sha256);
  if (!segment) throw new Error('manifest');
  return segment;
};

const install = async (releaseId, sha256) => {
  const segment = await resolveSegment(releaseId, sha256);
  const url = segmentUrl(segment.sha256);
  return self.navigator.locks.request(`onlyoffice-content:${segment.sha256}`, { mode: 'exclusive' }, async () => {
    const cache = await caches.open(CONTENT_CACHE);
    const cached = await cache.match(url);
    if (
      cached?.ok &&
      cached.headers.get('x-content-sha256') === segment.sha256 &&
      Number(cached.headers.get('content-length')) === segment.bytes
    ) {
      return { bytes: segment.bytes, sha256: segment.sha256, reused: true };
    }
    const response = await fetch(url, { cache: 'no-store', credentials: 'omit' });
    if (!response.ok) throw new Error(`network ${response.status}`);
    const bytes = await response.arrayBuffer();
    const digest = hex(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)));
    if (digest !== segment.sha256 || bytes.byteLength !== segment.bytes) throw new Error('integrity');
    await cache.put(
      url,
      new Response(bytes, {
        headers: {
          'content-length': String(bytes.byteLength),
          'content-type': 'application/vnd.onlyoffice.browser-pack-segment',
          'x-content-sha256': digest,
        },
      }),
    );
    return { bytes: bytes.byteLength, sha256: digest, reused: false };
  });
};

const read = async (port, message, requestKey, control) => {
  const segment = await resolveSegment(message.releaseId, message.sha256);
  const cache = await caches.open(CONTENT_CACHE);
  const response = await cache.match(segmentUrl(segment.sha256));
  if (
    !response?.body ||
    response.headers.get('x-content-sha256') !== segment.sha256 ||
    Number(response.headers.get('content-length')) !== segment.bytes
  ) {
    throw new Error('missing');
  }
  const hasRange = message.range !== undefined;
  const start = hasRange ? message.range?.start : 0;
  const requestedEnd = hasRange ? message.range?.end : segment.bytes - 1;
  if (
    (hasRange && (!Number.isSafeInteger(start) || !Number.isSafeInteger(requestedEnd))) ||
    start < 0 ||
    requestedEnd < start ||
    start >= segment.bytes
  ) {
    await response.body.cancel();
    port.postMessage({
      type: 'HEADERS',
      id: message.id,
      status: 416,
      headers: {
        acceptRanges: 'bytes',
        contentLength: 0,
        contentRange: `bytes */${segment.bytes}`,
        contentType: 'application/vnd.onlyoffice.browser-pack-segment',
      },
    });
    port.postMessage({ type: 'END', id: message.id, bytesSent: 0 });
    return;
  }
  const end = Math.min(requestedEnd, segment.bytes - 1);
  const state = {
    bytesSent: 0,
    cancelled: false,
    end,
    id: message.id,
    port,
    requestKey,
    reader: response.body.getReader(),
    sourceOffset: 0,
    start,
    buffered: new Uint8Array(0),
  };
  const completion = new Promise((resolve, reject) => {
    state.resolve = resolve;
    state.reject = reject;
  });
  control.state = state;
  activeReads.set(requestKey, state);
  port.postMessage({
    type: 'HEADERS',
    id: message.id,
    status: hasRange ? 206 : 200,
    headers: {
      acceptRanges: 'bytes',
      contentLength: end - start + 1,
      contentRange: hasRange ? `bytes ${start}-${end}/${segment.bytes}` : null,
      contentType: 'application/vnd.onlyoffice.browser-pack-segment',
    },
  });
  if (control.cancelled) {
    await cancelRead(state);
    return;
  }
  while (control.pendingPulls > 0) {
    control.pendingPulls -= 1;
    await nextChunk(state);
  }
  await completion;
};

const nextChunk = async (state) => {
  if (state.cancelled) return;
  while (state.buffered.byteLength === 0) {
    if (state.sourceOffset > state.end) {
      await state.reader.cancel();
      state.port.postMessage({ type: 'END', id: state.id, bytesSent: state.bytesSent });
      activeReads.delete(state.requestKey);
      state.resolve();
      return;
    }
    const result = await state.reader.read();
    if (result.done) {
      await state.reader.cancel();
      state.port.postMessage({ type: 'END', id: state.id, bytesSent: state.bytesSent });
      activeReads.delete(state.requestKey);
      state.resolve();
      return;
    }
    const sourceStart = state.sourceOffset;
    const sourceEnd = sourceStart + result.value.byteLength - 1;
    state.sourceOffset += result.value.byteLength;
    if (sourceEnd < state.start) continue;
    const overlapStart = Math.max(state.start, sourceStart);
    const overlapEnd = Math.min(state.end, sourceEnd);
    state.buffered = result.value.slice(overlapStart - sourceStart, overlapEnd - sourceStart + 1);
  }
  const piece = state.buffered.subarray(0, CHUNK_BYTES);
  state.buffered = state.buffered.subarray(piece.byteLength);
  const transferred = piece.buffer.slice(piece.byteOffset, piece.byteOffset + piece.byteLength);
  state.bytesSent += piece.byteLength;
  state.port.postMessage({ type: 'CHUNK', id: state.id, bytes: transferred }, [transferred]);
};

const cancelRead = async (state) => {
  if (!state || state.cancelled) return;
  state.cancelled = true;
  await state.reader.cancel();
  state.port.postMessage({ type: 'CANCELLED', id: state.id, bytesSent: state.bytesSent });
  activeReads.delete(state.requestKey);
  state.resolve();
};

const handleRequest = async (event) => {
  const port = event.ports[0];
  const message = event.data;
  if (!port || !message || typeof message.id !== 'string') return;
  const clientUrl = event.source?.url ? new URL(event.source.url) : null;
  if (
    clientUrl?.origin !== self.location.origin ||
    clientUrl.pathname !== '/__matrix__/cache-broker.html' ||
    clientUrl.searchParams.get('releaseId') !== message.releaseId ||
    clientUrl.searchParams.get('sessionId') !== message.sessionId
  ) {
    replyError(port, message.id, new Error('client'));
    port.close();
    return;
  }
  const requestKey = `${event.source.id}:${message.id}`;
  const control = { cancelled: false, pendingPulls: 0, state: null };
  port.onmessage = (portEvent) => {
    const state = control.state;
    if (portEvent.data?.type === 'PULL') {
      if (state) nextChunk(state).catch((error) => state.reject(error));
      else control.pendingPulls += 1;
    } else if (portEvent.data?.type === 'CANCEL') {
      if (state) cancelRead(state).catch((error) => state.reject(error));
      else control.cancelled = true;
    }
  };
  port.start();
  try {
    if (message.type === 'PROBE') {
      await loadSegments(message.releaseId);
      port.postMessage({ type: 'RESULT', id: message.id, value: { protocol: 1, releaseId: message.releaseId } });
      return;
    }
    if (message.type === 'INSTALL') {
      port.postMessage({ type: 'RESULT', id: message.id, value: await install(message.releaseId, message.sha256) });
      return;
    }
    if (message.type === 'STATS') {
      const cache = await caches.open(CONTENT_CACHE);
      port.postMessage({
        type: 'RESULT',
        id: message.id,
        value: {
          activeReads: activeReads.size,
          cacheNames: await caches.keys(),
          keys: (await cache.keys()).map((request) => request.url),
        },
      });
      return;
    }
    if (message.type === 'READ') {
      await read(port, message, requestKey, control);
      return;
    }
    throw new Error('unsupported');
  } catch (error) {
    replyError(port, message.id, error);
  } finally {
    port.close();
  }
};

self.addEventListener('message', (event) => {
  event.waitUntil(handleRequest(event));
});
