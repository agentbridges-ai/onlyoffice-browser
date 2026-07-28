const CACHE_VERSION = 'SW_VERSION_PLACEHOLDER'.includes('PLACEHOLDER') ? 'dev-' + Date.now() : 'SW_VERSION_PLACEHOLDER';
const CACHE_NAME = `document-editor-${CACHE_VERSION}`;
const PRINT_PDF_CACHE_NAME = 'onlyoffice-browser-print-pdfs';
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const ASSETS_TO_CACHE = ['./plugins.json', './themes.json'];
const ONLYOFFICE_RUNTIME_ASSET_REGEX = /(^|\/)(web-apps|sdkjs|wasm\/x2t)\//;
const ONLYOFFICE_NAVIGATION_PATHS = new Set(['/office-host.html', '/reset.html']);
const PWA_APP_NAVIGATION_PATHS = new Set(['/', '/index.html']);
const ONLYOFFICE_RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const CANONICAL_OFFICE_HOST = 'onlyoffice.getpi.work';
const EDITOR_HOST_PATTERN = /^office-[a-z0-9-]+\.getpi\.work$/;
const SHARED_ASSET_VERSION_QUERY = '__oobv';
const ORIGIN_BOUND_DESTINATIONS = new Set(['document', 'iframe', 'worker', 'sharedworker', 'serviceworker']);

// Cache limits and clean-up configuration
const MAX_CACHE_ITEMS = 100;

const isCanonicalPwaHost = self.location.hostname === CANONICAL_OFFICE_HOST;
const isIsolatedEditorHost = EDITOR_HOST_PATTERN.test(self.location.hostname);
let sharedAssetVersionPromise;

const shouldProxySharedAsset = (request, url) => {
  if (!isIsolatedEditorHost || ORIGIN_BOUND_DESTINATIONS.has(request.destination)) return false;
  if (url.pathname.startsWith('/assets/') || url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX)) return false;
  return !ONLYOFFICE_NAVIGATION_PATHS.has(url.pathname) &&
    url.pathname !== '/document_editor_service_worker.js' &&
    url.pathname !== '/sw.js';
};

const resolveSharedAssetVersion = async () => {
  if (!sharedAssetVersionPromise) {
    const manifestUrl = new URL(ONLYOFFICE_RUNTIME_MANIFEST_PATH, `https://${CANONICAL_OFFICE_HOST}`);
    sharedAssetVersionPromise = fetch(manifestUrl.href, {
      method: 'HEAD',
      cache: 'no-cache',
      credentials: 'omit',
    }).then((response) => {
      const version = response.headers.get('X-OnlyOffice-Asset-Version');
      if (!response.ok || !version) {
        throw new Error(`Shared Office asset version request failed (${response.status})`);
      }
      return version;
    }).catch((error) => {
      sharedAssetVersionPromise = undefined;
      throw error;
    });
  }
  return sharedAssetVersionPromise;
};

const fetchSharedAsset = async (request, url) => {
  try {
    const version = await resolveSharedAssetVersion();
    const canonicalUrl = new URL(url);
    canonicalUrl.hostname = CANONICAL_OFFICE_HOST;
    canonicalUrl.searchParams.set(SHARED_ASSET_VERSION_QUERY, version);
    const headers = new Headers();
    for (const name of ['accept', 'accept-language', 'range']) {
      const value = request.headers.get(name);
      if (value) headers.set(name, value);
    }
    return await fetch(canonicalUrl.href, {
      method: request.method,
      headers,
      mode: 'cors',
      credentials: 'omit',
      cache: 'force-cache',
      redirect: 'follow',
    });
  } catch {
    // Preserve the Worker redirect path as a safe fallback if version discovery
    // is temporarily unavailable.
    return fetch(request);
  }
};

const precachePwaShell = async (cache) => {
  const indexUrl = new URL('/index.html', self.location.origin);
  const response = await fetch(new Request(indexUrl.href, { cache: 'reload' }));
  if (!response.ok) throw new Error(`PWA shell request failed (${response.status}): ${indexUrl.pathname}`);
  const html = await response.clone().text();
  const assetUrls = [...html.matchAll(/(?:src|href)=["']([^"']+)["']/g)]
    .map((match) => new URL(match[1], indexUrl))
    .filter((url) => url.origin === self.location.origin)
    .map((url) => url.href);
  await cache.put(indexUrl.href, response.clone());
  await cache.put(new URL('/', self.location.origin).href, response);
  await cache.addAll([...new Set(assetUrls)]);
};

// Helper: Trim cache to a certain size
const limitCacheSize = (name, maxItems) => {
  caches.open(name).then((cache) => {
    cache.keys().then((keys) => {
      if (keys.length > maxItems) {
        cache.delete(keys[0]).then(() => limitCacheSize(name, maxItems));
      }
    });
  });
};

const isOnlyOfficeRuntimeAsset = (url) => ONLYOFFICE_RUNTIME_ASSET_REGEX.test(url.pathname);

const parseRangeHeader = (rangeHeader, byteLength) => {
  const match = /^bytes=(\d*)-(\d*)$/.exec(rangeHeader || '');
  if (!match || byteLength <= 0) return null;

  let start;
  let end;
  if (match[1] === '' && match[2] === '') return null;
  if (match[1] === '') {
    const suffixLength = Number(match[2]);
    if (!Number.isFinite(suffixLength) || suffixLength <= 0) return null;
    start = Math.max(byteLength - suffixLength, 0);
    end = byteLength - 1;
  } else {
    start = Number(match[1]);
    end = match[2] === '' ? byteLength - 1 : Number(match[2]);
  }

  if (!Number.isFinite(start) || !Number.isFinite(end) || start < 0 || end < start || start >= byteLength) {
    return null;
  }

  return {
    start,
    end: Math.min(end, byteLength - 1),
  };
};

const responseForCachedPrintPdf = async (request, cached) => {
  if (!cached) {
    return new Response('Print PDF expired', {
      status: 404,
      headers: { 'content-type': 'text/plain; charset=utf-8' },
    });
  }

  const rangeHeader = request.headers.get('range');
  if (!rangeHeader) {
    if (request.method === 'HEAD') {
      return new Response(null, {
        status: cached.status,
        statusText: cached.statusText,
        headers: cached.headers,
      });
    }
    return cached;
  }

  const bytes = await cached.arrayBuffer();
  const range = parseRangeHeader(rangeHeader, bytes.byteLength);
  if (!range) {
    return new Response(null, {
      status: 416,
      headers: {
        'content-range': `bytes */${bytes.byteLength}`,
      },
    });
  }

  const headers = new Headers(cached.headers);
  const body = bytes.slice(range.start, range.end + 1);
  headers.set('accept-ranges', 'bytes');
  headers.set('content-length', String(body.byteLength));
  headers.set('content-range', `bytes ${range.start}-${range.end}/${bytes.byteLength}`);
  headers.set('content-type', cached.headers.get('content-type') || 'application/pdf');

  return new Response(request.method === 'HEAD' ? null : body, {
    status: 206,
    statusText: 'Partial Content',
    headers,
  });
};

// Install event: Pre-cache core UI assets
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      if (isIsolatedEditorHost) return undefined;
      return cache
        .addAll(ASSETS_TO_CACHE)
        .then(() => (isCanonicalPwaHost ? precachePwaShell(cache) : undefined));
    }),
  );
  self.skipWaiting();
});

// Activate event: Clean up old caches
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (
            cacheName !== CACHE_NAME &&
            cacheName !== PRINT_PDF_CACHE_NAME
          ) {
            return caches.delete(cacheName);
          }
        }),
      );
    }),
  );
  self.clients.claim();
});

// Fetch event: Strategy-based resource handling
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // 1. Only handle same-origin requests to avoid caching external APIs/documents
  if (url.origin !== self.location.origin) return;

  // 2. Temporary print PDFs are written by the host page into Cache API and
  // served from a normal same-origin URL so OnlyOffice's built-in print iframe
  // does not have to load blob: URLs, which Chrome/extensions may block.
  if (url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX)) {
    if (event.request.method !== 'GET' && event.request.method !== 'HEAD') return;
    event.respondWith(
      caches.open(PRINT_PDF_CACHE_NAME).then((cache) =>
        cache.match(url.href).then((cached) => responseForCachedPrintPdf(event.request, cached)),
      ),
    );
    return;
  }

  // 3. Only handle GET requests for all other routes.
  if (event.request.method !== 'GET') return;

  const isCanonicalPwaNavigation =
    isCanonicalPwaHost &&
    event.request.mode === 'navigate' &&
    PWA_APP_NAVIGATION_PATHS.has(url.pathname);

  // 4. Only Office host documents and this deployment's canonical PWA shell
  // are handled as navigations. A broad integration scope must not intercept
  // another application's root or login page.
  if (
    event.request.mode === 'navigate' &&
    !ONLYOFFICE_NAVIGATION_PATHS.has(url.pathname) &&
    !isCanonicalPwaNavigation
  ) {
    return;
  }

  // 5. Never intercept application/API surfaces if this service worker is
  // installed at a broad scope by the host integration.
  if (
    url.pathname.startsWith('/api/') ||
    url.pathname.startsWith('/internal/') ||
    url.pathname.startsWith('/ws/') ||
    url.pathname.startsWith('/@vite/') ||
    url.pathname.startsWith('/@react-refresh') ||
    url.pathname.startsWith('/@id/') ||
    url.pathname.startsWith('/@fs/') ||
    url.pathname.startsWith('/node_modules/') ||
    url.pathname.startsWith('/src/')
  ) {
    return;
  }

  // 6. Skip caching for requests with dynamic parameters (like ?file= or ?src=)
  // These are typically documents being edited, which should always be fresh.
  if (url.searchParams.has('file') || url.searchParams.has('src')) return;

  // 7. Keep editor documents, workers, and the origin-bound host bundle on the
  // unique editor origin, but fetch every shareable static asset directly from
  // the immutable canonical URL. This removes the wildcard host's 307 cascade
  // while preserving one browser HTTP-cache entry shared by every editor.
  if (shouldProxySharedAsset(event.request, url)) {
    event.respondWith(fetchSharedAsset(event.request, url));
    return;
  }

  if (isCanonicalPwaNavigation) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          if (response.ok) {
            caches.open(CACHE_NAME).then((cache) =>
              cache.put(new URL('/index.html', self.location.origin).href, response.clone()),
            );
          }
          return response;
        })
        .catch(() =>
          caches.open(CACHE_NAME).then((cache) =>
            cache.match(new URL('/index.html', self.location.origin).href),
          ),
        ),
    );
    return;
  }

  // 8. Skip font files — let the browser cache them natively to avoid SW
  // interception latency triggering Chrome's font-loading intervention in
  // OnlyOffice's fallback font loading path.
  if (url.pathname.startsWith('/fonts/') || /\.(ttf|tte|ttc|otf|otc|woff2?|eot)(\?.*)?$/.test(url.pathname)) return;

  // 9. Determine Strategy
  const isHtml =
    event.request.mode === 'navigate' ||
    url.pathname.endsWith('.html') ||
    url.pathname === '/' ||
    url.pathname.endsWith('/');

  if (isHtml || isOnlyOfficeRuntimeAsset(url) || url.pathname === ONLYOFFICE_RUNTIME_MANIFEST_PATH) {
    // Strategy: Network-First for HTML/Navigation
    // Ensuring the user always gets the latest version if online,
    // but can still access the app when offline.
    event.respondWith(
      fetch(event.request)
        .then((networkResponse) => {
          // If network is ok, cache and return
          if (networkResponse && networkResponse.status === 200) {
            const responseToCache = networkResponse.clone();
            caches.open(CACHE_NAME).then((cache) => {
              cache.put(event.request, responseToCache);
              limitCacheSize(CACHE_NAME, MAX_CACHE_ITEMS);
            });
            return networkResponse;
          }
          // If status is not 200, try cache
          return caches.match(event.request).then((cached) => cached || networkResponse);
        })
        .catch(() => {
          // If fetch fails (offline), try cache
          return caches.match(event.request);
        }),
    );
  } else {
    // Strategy: Stale-While-Revalidate for other static assets (JS, CSS, Images)
    event.respondWith(
      caches.match(event.request).then((cachedResponse) => {
        const fetchPromise = fetch(event.request)
          .then((networkResponse) => {
            // Only cache valid 200 responses
            if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
              const responseToCache = networkResponse.clone();
              caches.open(CACHE_NAME).then((cache) => {
                cache.put(event.request, responseToCache);
                limitCacheSize(CACHE_NAME, MAX_CACHE_ITEMS);
              });
            }
            return networkResponse;
          })
          .catch(() => {
            return cachedResponse;
          });

        return cachedResponse || fetchPromise;
      }),
    );
  }
});
