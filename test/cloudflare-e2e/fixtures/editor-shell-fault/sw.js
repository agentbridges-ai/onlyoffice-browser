const CACHE_NAME = 'onlyoffice-editor-shell-fault-v1';
const RELEASE_ID = 'matrix-shell-release-v1';
const SHELL_PATH = '/__matrix__/editor-shell-fault/office-host.html';
const protocol = 'onlyoffice-editor-shell-fault-v1';
const fault = new URL(self.location.href).searchParams.get('fault') || 'none';

const sleep = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));
const storageError = (name) => new DOMException(`simulated Cache Storage ${name}`, name);

async function prime() {
  if (fault === 'primeDelayMs') await sleep(1_500);
  let storageMode = 'cache';
  try {
    if (fault === 'cacheStorageOpenError') throw storageError('QuotaExceededError');
    const cache = await caches.open(CACHE_NAME);
    if (fault === 'cacheStoragePutError') throw storageError('UnknownError');
    await cache.put(
      new Request(SHELL_PATH),
      new Response('<!doctype html><title>cached shell</title>', {
        headers: {
          'content-type': 'text/html; charset=utf-8',
          'content-length': '42',
          'x-onlyoffice-asset-version': RELEASE_ID,
          'x-onlyoffice-editor-shell-cache': '1',
        },
      }),
    );
  } catch (error) {
    if (!['QuotaExceededError', 'UnknownError', 'InvalidStateError', 'NotAllowedError'].includes(error?.name))
      throw error;
    storageMode = 'network';
  }
  return { type: 'ONLYOFFICE_EDITOR_SHELL_PRIMED', releaseId: RELEASE_ID, storageMode };
}

self.addEventListener('install', (event) => {
  event.waitUntil(fault === 'serviceWorkerNoController' ? Promise.resolve() : self.skipWaiting());
});

self.addEventListener('activate', (event) => {
  event.waitUntil(fault === 'serviceWorkerNoController' ? Promise.resolve() : self.clients.claim());
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  if (url.origin !== self.location.origin || url.pathname !== SHELL_PATH) return;
  event.respondWith(
    (async () => {
      if (fault === 'shellNetwork404') return new Response('missing', { status: 404 });
      if (fault === 'shellReleaseMismatch') {
        return new Response('<!doctype html><title>wrong release</title>', {
          status: 200,
          headers: {
            'content-type': 'text/html; charset=utf-8',
            'content-length': '48',
            'x-onlyoffice-asset-version': 'matrix-shell-release-old',
            'x-onlyoffice-editor-shell-storage': 'network',
          },
        });
      }
      const cached = await caches.match(event.request, { cacheName: CACHE_NAME });
      if (cached) return cached;
      const response = await fetch(event.request, { cache: 'no-store', credentials: 'omit' });
      const headers = new Headers(response.headers);
      headers.set('x-onlyoffice-asset-version', RELEASE_ID);
      headers.set('x-onlyoffice-editor-shell-storage', 'network');
      return new Response(response.body, { status: response.status, headers });
    })(),
  );
});

self.addEventListener('message', (event) => {
  if (event.data?.protocol !== protocol || event.ports.length !== 1 || event.data.releaseId !== RELEASE_ID) return;
  const port = event.ports[0];
  event.waitUntil(
    (async () => {
      if (event.data.type === 'ONLYOFFICE_PRIME_EDITOR_SHELL') {
        port.postMessage(await prime());
      } else if (event.data.type === 'ONLYOFFICE_BROKER_PROBE') {
        if (fault === 'brokerProbeTimeout') {
          await sleep(1_500);
          return;
        }
        if (fault === 'offlineAfterLastObject') {
          port.postMessage({ type: 'ONLYOFFICE_BROKER_PROBE_RESULT', ok: false, code: 'offline' });
          return;
        }
        if (fault === 'abortAtLastMegabyte') {
          port.postMessage({ type: 'ONLYOFFICE_BROKER_PROBE_RESULT', ok: false, code: 'aborted' });
          return;
        }
        port.postMessage({ type: 'ONLYOFFICE_BROKER_PROBE_RESULT', ok: true });
      }
    })().catch((error) =>
      port.postMessage({ type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_ERROR', code: error?.name || 'storage' }),
    ),
  );
});
