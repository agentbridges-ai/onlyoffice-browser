import { clientsClaim, setCacheNameDetails } from 'workbox-core';
import { ExpirationPlugin } from 'workbox-expiration';
import { cleanupOutdatedCaches, matchPrecache, precacheAndRoute } from 'workbox-precaching';
import { registerRoute } from 'workbox-routing';
import { NetworkFirst, StaleWhileRevalidate } from 'workbox-strategies';

const SERVICE_WORKER_VERSION = 'SW_VERSION_PLACEHOLDER';
const PRINT_PDF_CACHE_NAME = 'onlyoffice-browser-print-pdfs';
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const ONLYOFFICE_RUNTIME_ASSET_REGEX = /(^|\/)(web-apps|sdkjs|wasm\/x2t)\//;
const ONLYOFFICE_NAVIGATION_PATHS = new Set(['/office-host.html', '/reset.html']);
const PWA_APP_NAVIGATION_PATHS = new Set(['/', '/index.html']);
const ONLYOFFICE_RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const FONT_MANIFEST_PATH = '/onlyoffice-browser-font-assets.json';
const CANONICAL_OFFICE_HOST = 'onlyoffice.getpi.work';
const EDITOR_HOST_PATTERN = /^office-[a-z0-9-]+\.getpi\.work$/;
const LOCAL_EDITOR_HOST_PATTERN = /^host-office-editor-[a-z0-9-]+\.office\.localhost$/;
const LOCAL_CANONICAL_OFFICE_HOST = 'assets.office.localhost';
const SHARED_ASSET_VERSION_QUERY = '__oobv';
const ORIGIN_BOUND_DESTINATIONS = new Set(['document', 'iframe', 'worker', 'sharedworker', 'serviceworker']);

const MAX_CACHE_ITEMS = 100;
const RUNTIME_CACHE_NAME = 'onlyoffice-browser-runtime-v1';
const STATIC_CACHE_NAME = 'onlyoffice-browser-static-v1';

const isCanonicalPwaHost = self.location.hostname === CANONICAL_OFFICE_HOST;
const isLocalEditorHost = LOCAL_EDITOR_HOST_PATTERN.test(self.location.hostname);
const isIsolatedEditorHost = EDITOR_HOST_PATTERN.test(self.location.hostname) || isLocalEditorHost;
const canonicalOfficeOrigin = isLocalEditorHost
  ? `${self.location.protocol}//${LOCAL_CANONICAL_OFFICE_HOST}${self.location.port ? `:${self.location.port}` : ''}`
  : `https://${CANONICAL_OFFICE_HOST}`;
let sharedAssetManifestPromise;
let downloadedFontPaths = new Set();

const fontFallbackRole = (families = []) => {
  const names = families.join(' ').toLowerCase();
  if (names.includes('emoji')) return 'emoji';
  if (/(?:\bmath\b|stix)/.test(names)) return 'math';
  if (/(?:arabic|amiri|kacst|kufi|naskh|scheherazade)/.test(names)) return 'arabic';
  if (/(?:batang|dotum|gulim|gungsuh|malgun|nanum|noto sans kr)/.test(names)) return 'korean';
  if (/(?:meiryo|ms (?:p?gothic|p?mincho)|noto sans jp|takao|yu (?:gothic|mincho))/.test(names)) {
    return 'japanese';
  }
  if (/(?:symbol|wingdings|webdings|dingbats|marlett|monotype sorts|mt extra)/.test(names)) return 'symbol';
  return 'default';
};

const selectFallbackFont = ({ requested, fallbackFonts, defaultFonts, fontAssets }) => {
  const role = fontFallbackRole(requested?.families);
  const candidates = fallbackFonts[role] || fallbackFonts.default || defaultFonts;
  const requestedStyles = new Set(requested?.styles || []);
  const wantsBold = requestedStyles.has('bold') || requestedStyles.has('boldItalic');
  const wantsItalic = requestedStyles.has('italic') || requestedStyles.has('boldItalic');
  const exactStyle = candidates.find((candidatePath) => {
    const styles = new Set(fontAssets.get(candidatePath)?.styles || []);
    return (
      (styles.has('bold') || styles.has('boldItalic')) === wantsBold &&
      (styles.has('italic') || styles.has('boldItalic')) === wantsItalic
    );
  });
  if (exactStyle) return exactStyle;
  return (
    candidates.find((candidatePath) => {
      const styles = new Set(fontAssets.get(candidatePath)?.styles || []);
      return (styles.has('bold') || styles.has('boldItalic')) === wantsBold;
    }) ||
    candidates[0] ||
    defaultFonts[0]
  );
};

setCacheNameDetails({
  prefix: 'onlyoffice-browser',
  precache: 'shell-v1',
  suffix: '',
});
if (isCanonicalPwaHost) {
  precacheAndRoute(self.__WB_MANIFEST || []);
  cleanupOutdatedCaches();
}
if (isIsolatedEditorHost) {
  clientsClaim();
}

const shouldProxySharedAsset = (request, url) => {
  if (!isIsolatedEditorHost || ORIGIN_BOUND_DESTINATIONS.has(request.destination)) return false;
  if (url.pathname.startsWith('/assets/') || url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX)) return false;
  return (
    !ONLYOFFICE_NAVIGATION_PATHS.has(url.pathname) &&
    url.pathname !== '/document_editor_service_worker.js' &&
    url.pathname !== '/sw.js'
  );
};

const loadVersionedJson = async (pathname, version) => {
  const url = new URL(pathname, canonicalOfficeOrigin);
  url.searchParams.set(SHARED_ASSET_VERSION_QUERY, version);
  const response = await fetch(url.href, {
    cache: 'force-cache',
    credentials: 'omit',
  });
  if (!response.ok) throw new Error(`Shared Office manifest request failed (${response.status}): ${pathname}`);
  return response.json();
};

const resolveSharedAssetManifest = async () => {
  if (!sharedAssetManifestPromise) {
    const manifestUrl = new URL(ONLYOFFICE_RUNTIME_MANIFEST_PATH, canonicalOfficeOrigin);
    sharedAssetManifestPromise = fetch(manifestUrl.href, {
      method: 'HEAD',
      cache: 'no-cache',
      credentials: 'omit',
    })
      .then(async (response) => {
        const versionHeader =
          response.headers.get('X-OnlyOffice-Asset-Version') ||
          response.headers.get('etag') ||
          response.headers.get('last-modified');
        const version = versionHeader?.replace(/^W\//, '').replaceAll('"', '');
        if (!response.ok || !version) {
          throw new Error(`Shared Office asset version request failed (${response.status})`);
        }
        const [runtime, fonts] = await Promise.all([
          loadVersionedJson(ONLYOFFICE_RUNTIME_MANIFEST_PATH, version),
          loadVersionedJson(FONT_MANIFEST_PATH, version),
        ]);
        const revisions = new Map();
        const fontAssets = new Map();
        for (const asset of [...(runtime.assets || []), ...(fonts.assets || [])]) {
          if (typeof asset?.path === 'string' && typeof asset?.revision === 'string') {
            revisions.set(asset.path, asset.revision);
            if (asset.path.startsWith('fonts/')) fontAssets.set(asset.path, asset);
          }
        }
        revisions.set(ONLYOFFICE_RUNTIME_MANIFEST_PATH.slice(1), version);
        revisions.set(FONT_MANIFEST_PATH.slice(1), version);
        return {
          version,
          revisions,
          fontAssets,
          defaultFonts: Array.isArray(fonts.defaultFonts) ? fonts.defaultFonts : [],
          builtInFonts: Array.isArray(fonts.builtInFonts) ? fonts.builtInFonts : [],
          fallbackFonts:
            fonts.fallbackFonts && typeof fonts.fallbackFonts === 'object'
              ? Object.fromEntries(
                  Object.entries(fonts.fallbackFonts).filter(
                    ([, paths]) => Array.isArray(paths) && paths.every((path) => typeof path === 'string'),
                  ),
                )
              : {},
        };
      })
      .catch((error) => {
        sharedAssetManifestPromise = undefined;
        throw error;
      });
  }
  return sharedAssetManifestPromise;
};

const fetchSharedAsset = async (request, url) => {
  try {
    const { version, revisions, fontAssets, defaultFonts, builtInFonts, fallbackFonts } =
      await resolveSharedAssetManifest();
    const canonicalUrl = new URL(`${url.pathname}${url.search}`, canonicalOfficeOrigin);
    let path = canonicalUrl.pathname.slice(1);
    if (
      path.startsWith('fonts/') &&
      !downloadedFontPaths.has(path) &&
      !defaultFonts.includes(path) &&
      !builtInFonts.includes(path)
    ) {
      const requested = fontAssets.get(path);
      path = selectFallbackFont({ requested, fallbackFonts, defaultFonts, fontAssets }) || path;
      canonicalUrl.pathname = `/${path}`;
    }
    canonicalUrl.searchParams.set(SHARED_ASSET_VERSION_QUERY, revisions.get(path) || version);
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
  } catch (error) {
    if (url.pathname.startsWith('/fonts/')) {
      console.error('[onlyoffice-browser] Shared font fetch failed', error);
      const diagnostic = isLocalEditorHost && error instanceof Error ? `: ${error.message}` : '';
      return new Response(`Font fallback metadata is unavailable${diagnostic}`, {
        status: 503,
        statusText: 'Font Fallback Unavailable',
      });
    }
    // Preserve the Worker redirect path as a safe fallback if version discovery
    // is temporarily unavailable. Font requests fail closed above so an
    // optional font can never bypass the download allowlist.
    return fetch(request);
  }
};

const isOnlyOfficeRuntimeAsset = (url) => ONLYOFFICE_RUNTIME_ASSET_REGEX.test(url.pathname);

const isExcludedApplicationPath = (url) =>
  url.pathname.startsWith('/api/') ||
  url.pathname.startsWith('/internal/') ||
  url.pathname.startsWith('/ws/') ||
  url.pathname.startsWith('/@vite/') ||
  url.pathname.startsWith('/@react-refresh') ||
  url.pathname.startsWith('/@id/') ||
  url.pathname.startsWith('/@fs/') ||
  url.pathname.startsWith('/node_modules/') ||
  url.pathname.startsWith('/src/');

const hasDynamicDocumentParameter = (url) => url.searchParams.has('file') || url.searchParams.has('src');

const isAllowedNavigation = (request, url) =>
  request.mode !== 'navigate' ||
  ONLYOFFICE_NAVIGATION_PATHS.has(url.pathname) ||
  (isCanonicalPwaHost && PWA_APP_NAVIGATION_PATHS.has(url.pathname));

const isFontRequest = (url) =>
  url.pathname.startsWith('/fonts/') || /\.(ttf|tte|ttc|otf|otc|woff2?|eot)$/.test(url.pathname);

const isEligibleGet = (request, url) =>
  request.method === 'GET' &&
  isAllowedNavigation(request, url) &&
  !isExcludedApplicationPath(url) &&
  !hasDynamicDocumentParameter(url);

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

self.addEventListener('message', (event) => {
  if (event.data?.type === 'SKIP_WAITING') {
    void self.skipWaiting();
  } else if (event.data?.type === 'SET_FONT_ALLOWLIST') {
    downloadedFontPaths = new Set(
      Array.isArray(event.data.paths)
        ? event.data.paths.filter((path) => typeof path === 'string' && /^fonts\/[^/]+$/.test(path))
        : [],
    );
    event.ports[0]?.postMessage({ ok: true });
  } else if (event.data?.type === 'GET_VERSION') {
    event.ports[0]?.postMessage({ version: SERVICE_WORKER_VERSION });
  }
});

const expirationPlugin = () =>
  new ExpirationPlugin({
    maxEntries: MAX_CACHE_ITEMS,
    purgeOnQuotaError: true,
  });

const runtimeNetworkFirst = new NetworkFirst({
  cacheName: RUNTIME_CACHE_NAME,
  networkTimeoutSeconds: 3,
  plugins: [expirationPlugin()],
});

const staticStaleWhileRevalidate = new StaleWhileRevalidate({
  cacheName: STATIC_CACHE_NAME,
  plugins: [expirationPlugin()],
});

registerRoute(
  ({ request, sameOrigin, url }) =>
    sameOrigin && isCanonicalPwaHost && request.mode === 'navigate' && url.pathname === '/',
  async () => (await matchPrecache('/index.html')) || fetch('/index.html'),
  'GET',
);

const printRouteHandler = async ({ request, url }) => {
  const cache = await caches.open(PRINT_PDF_CACHE_NAME);
  return responseForCachedPrintPdf(request, await cache.match(url.href));
};

// Temporary print PDFs support both full and Range/HEAD requests.
registerRoute(
  ({ sameOrigin, url }) => sameOrigin && url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX),
  printRouteHandler,
  'GET',
);
registerRoute(
  ({ sameOrigin, url }) => sameOrigin && url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX),
  printRouteHandler,
  'HEAD',
);

// A unique editor origin keeps its host documents and workers origin-bound,
// while every shareable request uses the canonical content-revision URL.
registerRoute(
  ({ request, sameOrigin, url }) => sameOrigin && isEligibleGet(request, url) && shouldProxySharedAsset(request, url),
  ({ request, url }) => fetchSharedAsset(request, url),
  'GET',
);

registerRoute(
  ({ request, sameOrigin, url }) => {
    if (
      !sameOrigin ||
      url.searchParams.has(SHARED_ASSET_VERSION_QUERY) ||
      !isEligibleGet(request, url) ||
      shouldProxySharedAsset(request, url) ||
      isFontRequest(url)
    ) {
      return false;
    }
    const isHtml =
      request.mode === 'navigate' ||
      url.pathname.endsWith('.html') ||
      url.pathname === '/' ||
      url.pathname.endsWith('/');
    return isHtml || isOnlyOfficeRuntimeAsset(url) || url.pathname === ONLYOFFICE_RUNTIME_MANIFEST_PATH;
  },
  runtimeNetworkFirst,
  'GET',
);

registerRoute(
  ({ request, sameOrigin, url }) =>
    sameOrigin &&
    !url.searchParams.has(SHARED_ASSET_VERSION_QUERY) &&
    isEligibleGet(request, url) &&
    !shouldProxySharedAsset(request, url) &&
    !isFontRequest(url),
  staticStaleWhileRevalidate,
  'GET',
);
