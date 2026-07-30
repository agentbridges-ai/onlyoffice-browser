import { clientsClaim, setCacheNameDetails } from 'workbox-core';
import { ExpirationPlugin } from 'workbox-expiration';
import { cleanupOutdatedCaches, matchPrecache, precacheAndRoute } from 'workbox-precaching';
import { registerRoute } from 'workbox-routing';
import { NetworkFirst, StaleWhileRevalidate } from 'workbox-strategies';
import { isProductionOfficeEditorHostname } from './lib/office-origin-pool';

const SERVICE_WORKER_VERSION = 'SW_VERSION_PLACEHOLDER';
const PRINT_PDF_CACHE_NAME = 'onlyoffice-browser-print-pdfs';
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const ONLYOFFICE_RUNTIME_ASSET_REGEX = /(^|\/)(web-apps|sdkjs|wasm\/x2t)\//;
const ONLYOFFICE_NAVIGATION_PATHS = new Set(['/office-host.html', '/reset.html']);
const PWA_APP_NAVIGATION_PATHS = new Set(['/', '/index.html']);
const ONLYOFFICE_RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const FONT_MANIFEST_PATH = '/onlyoffice-browser-font-assets.json';
const RELEASE_CHANNEL_PATH = '/channels/stable.json';
const CANONICAL_OFFICE_HOST = 'onlyoffice.getpi.work';
const LEGACY_EDITOR_HOST_PATTERN = /^office-editor-[a-z0-9-]+\.getpi\.work$/;
const LOCAL_EDITOR_HOST_PATTERN =
  /^host-(?:aries|taurus|gemini|cancer|leo|virgo|libra|scorpio|sagittarius|capricorn|aquarius|pisces)\.office\.localhost$/;
const LOCAL_LEGACY_EDITOR_HOST_PATTERN = /^host-office-editor-[a-z0-9-]+\.office\.localhost$/;
// Match the local PWA origin so the matrix exercises the same cross-editor
// HTTP cache key sharing used by onlyoffice.getpi.work in production.
const LOCAL_CANONICAL_OFFICE_HOST = 'onlyoffice.localhost';
const SHARED_ASSET_VERSION_QUERY = '__oobv';
const ORIGIN_BOUND_DESTINATIONS = new Set(['document', 'iframe', 'worker', 'sharedworker', 'serviceworker']);

const MAX_CACHE_ITEMS = 100;
const MAX_PACKAGE_SEGMENT_BUFFERS = 4;
const PACKAGE_SEGMENT_CACHE_NAME = 'onlyoffice-browser-package-segments-v1';
const RUNTIME_CACHE_NAME = 'onlyoffice-browser-runtime-v1';
const STATIC_CACHE_NAME = 'onlyoffice-browser-static-v1';

const isCanonicalPwaHost = self.location.hostname === CANONICAL_OFFICE_HOST;
const isLocalEditorHost =
  LOCAL_EDITOR_HOST_PATTERN.test(self.location.hostname) ||
  LOCAL_LEGACY_EDITOR_HOST_PATTERN.test(self.location.hostname);
const isIsolatedEditorHost =
  isProductionOfficeEditorHostname(self.location.hostname) ||
  LEGACY_EDITOR_HOST_PATTERN.test(self.location.hostname) ||
  isLocalEditorHost;
const canonicalOfficeOrigin = isLocalEditorHost
  ? `${self.location.protocol}//${LOCAL_CANONICAL_OFFICE_HOST}${self.location.port ? `:${self.location.port}` : ''}`
  : `https://${CANONICAL_OFFICE_HOST}`;
let sharedAssetManifestPromise;
const packageSegmentBuffers = new Map();
const packageSegmentCachePrunes = new Map();
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

const loadCurrentRelease = async () => {
  const channelUrl = new URL(RELEASE_CHANNEL_PATH, canonicalOfficeOrigin);
  const channelResponse = await fetch(channelUrl.href, {
    cache: 'no-store',
    credentials: 'omit',
  });
  if (!channelResponse.ok) {
    throw new Error(`Office release channel request failed (${channelResponse.status})`);
  }
  const channel = await channelResponse.json();
  if (channel?.version !== 1 || typeof channel.releaseId !== 'string') {
    throw new Error('Office release channel is invalid');
  }
  const manifestUrl = new URL(
    `/releases/${encodeURIComponent(channel.releaseId)}/manifest.json`,
    canonicalOfficeOrigin,
  );
  const manifestResponse = await fetch(manifestUrl.href, {
    cache: 'no-store',
    credentials: 'omit',
  });
  if (!manifestResponse.ok) {
    throw new Error(`Office release manifest request failed (${manifestResponse.status})`);
  }
  const manifest = await manifestResponse.json();
  if (
    (manifest?.version !== 3 && manifest?.version !== 4) ||
    manifest.releaseId !== channel.releaseId ||
    !Array.isArray(manifest.assets)
  ) {
    throw new Error('Office release manifest is invalid');
  }
  return manifest;
};

const loadCurrentReleaseForOrigin = async () => {
  try {
    return await loadCurrentRelease();
  } catch (error) {
    // Vite's local assets.office.localhost origin serves generated assets
    // directly and intentionally has no Cloudflare channel/release routes.
    // Keep production fail-closed while allowing the same generated
    // AllFonts.js/font files to be exercised by local editor tests.
    if (isLocalEditorHost) return null;
    throw error;
  }
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
        const [runtime, fonts, release] = await Promise.all([
          loadVersionedJson(ONLYOFFICE_RUNTIME_MANIFEST_PATH, version),
          loadVersionedJson(FONT_MANIFEST_PATH, version),
          loadCurrentReleaseForOrigin(),
        ]);
        const revisions = new Map();
        for (const asset of [...(runtime.assets || []), ...(fonts.assets || [])]) {
          if (typeof asset?.path === 'string' && typeof asset?.revision === 'string') {
            revisions.set(asset.path, asset.revision);
          }
        }
        revisions.set(ONLYOFFICE_RUNTIME_MANIFEST_PATH.slice(1), version);
        revisions.set(FONT_MANIFEST_PATH.slice(1), version);
        return {
          version,
          revisions,
          fonts,
          release,
        };
      })
      .catch((error) => {
        sharedAssetManifestPromise = undefined;
        throw error;
      });
  }
  return sharedAssetManifestPromise;
};

const digestHex = async (bytes) =>
  Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), (byte) =>
    byte.toString(16).padStart(2, '0'),
  ).join('');

const readPersistedOfficePackSegment = async (cache, segmentUrl, segment) => {
  const response = await cache.match(segmentUrl.href);
  if (!response) return null;
  const expectedDigest = response.headers.get('x-onlyoffice-segment-sha256');
  const expectedBytes = Number(response.headers.get('content-length'));
  if (
    !response.ok ||
    expectedDigest !== segment.sha256 ||
    !Number.isSafeInteger(expectedBytes) ||
    expectedBytes !== segment.bytes
  ) {
    await cache.delete(segmentUrl.href);
    return null;
  }
  const bytes = await response.arrayBuffer();
  if (bytes.byteLength === segment.bytes && (await digestHex(bytes)) === segment.sha256) return bytes;
  await cache.delete(segmentUrl.href);
  return null;
};

const prunePersistedOfficePackSegments = (cache, releaseId) => {
  const existing = packageSegmentCachePrunes.get(releaseId);
  if (existing) return existing;
  const currentReleasePrefix = `/p/${encodeURIComponent(releaseId)}/`;
  const prune = cache
    .keys()
    .then((requests) =>
      Promise.all(
        requests
          .filter(({ url }) => {
            const pathname = new URL(url).pathname;
            return pathname.startsWith('/p/') && !pathname.startsWith(currentReleasePrefix);
          })
          .map((request) => cache.delete(request)),
      ),
    )
    .catch((error) => {
      console.warn('[onlyoffice-browser] Unable to prune stale Office Pack segments', {
        releaseId,
        error: error instanceof Error ? error.message : String(error),
      });
    });
  packageSegmentCachePrunes.set(releaseId, prune);
  return prune;
};

const loadOfficePackSegment = (release, segment) => {
  const key = `${release.releaseId}/${segment.id}`;
  const existing = packageSegmentBuffers.get(key);
  if (existing) return existing;
  const promise = (async () => {
    const segmentUrl = new URL(
      `/p/${encodeURIComponent(release.releaseId)}/office-resources.oobpack`,
      canonicalOfficeOrigin,
    );
    segmentUrl.searchParams.set('segment', segment.id);
    const cache = await caches.open(PACKAGE_SEGMENT_CACHE_NAME);
    await prunePersistedOfficePackSegments(cache, release.releaseId);
    const persisted = await readPersistedOfficePackSegment(cache, segmentUrl, segment);
    if (persisted) return persisted;
    const response = await fetch(segmentUrl.href, {
      mode: 'cors',
      credentials: 'omit',
      cache: 'force-cache',
    });
    if (response.status !== 200) {
      response.body?.cancel();
      throw new Error(`Office Pack segment request failed (${response.status}): ${segment.id}`);
    }
    const bytes = await response.arrayBuffer();
    if (bytes.byteLength !== segment.bytes || (await digestHex(bytes)) !== segment.sha256) {
      throw new Error(`Office Pack segment integrity mismatch: ${segment.id}`);
    }
    const headers = new Headers(response.headers);
    headers.delete('content-encoding');
    headers.delete('content-range');
    headers.delete('transfer-encoding');
    headers.set('content-length', String(segment.bytes));
    headers.set('x-onlyoffice-segment-sha256', segment.sha256);
    try {
      await cache.put(
        segmentUrl.href,
        new Response(bytes.slice(0), {
          status: 200,
          headers,
        }),
      );
    } catch (error) {
      console.warn('[onlyoffice-browser] Unable to persist verified Office Pack segment', {
        releaseId: release.releaseId,
        segmentId: segment.id,
        error: error instanceof Error ? error.message : String(error),
      });
    }
    return bytes;
  })().catch((error) => {
    packageSegmentBuffers.delete(key);
    throw error;
  });
  packageSegmentBuffers.set(key, promise);
  while (packageSegmentBuffers.size > MAX_PACKAGE_SEGMENT_BUFFERS) {
    const oldest = packageSegmentBuffers.keys().next().value;
    if (!oldest || oldest === key) break;
    packageSegmentBuffers.delete(oldest);
  }
  return promise;
};

const fetchOfficePackAsset = async (request, asset, release) => {
  const pack = release?.package;
  if (
    release?.version !== 4 ||
    pack?.format !== 'onlyoffice-pack-v1' ||
    pack.path !== 'office-resources.oobpack' ||
    !Array.isArray(pack.segments) ||
    !Number.isSafeInteger(asset.packageOffset) ||
    !Number.isSafeInteger(asset.bytes) ||
    asset.bytes <= 0
  ) {
    return null;
  }
  const requestedRange = request.headers.has('range')
    ? parseRangeHeader(request.headers.get('range'), asset.bytes)
    : null;
  if (request.headers.has('range') && !requestedRange) {
    return new Response(null, {
      status: 416,
      headers: { 'content-range': `bytes */${asset.bytes}` },
    });
  }
  const relativeStart = requestedRange?.start || 0;
  const relativeEnd = requestedRange?.end ?? asset.bytes - 1;
  const start = asset.packageOffset + relativeStart;
  const end = asset.packageOffset + relativeEnd;
  const expectedLength = relativeEnd - relativeStart + 1;
  const overlappingSegments = pack.segments.filter(
    (segment) =>
      Number.isSafeInteger(segment?.offset) &&
      Number.isSafeInteger(segment?.bytes) &&
      segment.bytes > 0 &&
      start < segment.offset + segment.bytes &&
      end >= segment.offset,
  );
  const output = new Uint8Array(expectedLength);
  let outputOffset = 0;
  let receivedLength = 0;
  for (const segment of overlappingSegments) {
    const overlapStart = Math.max(start, segment.offset);
    const overlapEnd = Math.min(end, segment.offset + segment.bytes - 1);
    const rangeLength = overlapEnd - overlapStart + 1;
    const segmentBytes = new Uint8Array(await loadOfficePackSegment(release, segment));
    output.set(segmentBytes.subarray(overlapStart - segment.offset, overlapEnd - segment.offset + 1), outputOffset);
    outputOffset += rangeLength;
    receivedLength += rangeLength;
  }
  if (receivedLength !== expectedLength || overlappingSegments.length === 0) {
    throw new Error(`Office Pack segment coverage mismatch: ${asset.path}`);
  }
  const headers = new Headers();
  headers.set('accept-ranges', 'bytes');
  headers.set('content-length', String(expectedLength));
  headers.set('content-type', asset.mime || 'application/octet-stream');
  headers.set('x-onlyoffice-pack-release', release.releaseId);
  if (requestedRange) {
    headers.set('content-range', `bytes ${relativeStart}-${relativeEnd}/${asset.bytes}`);
  } else {
    headers.delete('content-range');
  }
  return new Response(output, {
    status: requestedRange ? 206 : 200,
    headers,
  });
};

const fetchSharedAsset = async (request, url) => {
  try {
    const { version, revisions, release } = await resolveSharedAssetManifest();
    const canonicalUrl = new URL(`${url.pathname}${url.search}`, canonicalOfficeOrigin);
    const path = canonicalUrl.pathname.slice(1);
    // Keep each font URL bound to its own bytes. The x2t converter mounts the
    // complete generated font set, while editor-only family substitution is
    // handled by AscFonts before glyph selection.
    canonicalUrl.searchParams.set(SHARED_ASSET_VERSION_QUERY, revisions.get(path) || version);
    const headers = new Headers();
    for (const name of ['accept', 'accept-language', 'range']) {
      const value = request.headers.get(name);
      if (value) headers.set(name, value);
    }
    const releaseAsset = release?.assets?.find((asset) => asset?.path === path);
    const packedResponse = releaseAsset ? await fetchOfficePackAsset(request, releaseAsset, release) : null;
    const response =
      packedResponse ||
      (await fetch(canonicalUrl.href, {
        method: request.method,
        headers,
        mode: 'cors',
        credentials: 'omit',
        cache: 'force-cache',
        redirect: 'follow',
      }));
    // AllFonts.js and font_selection.bin are generated by DocumentServer as a
    // matched set. Return the metadata unchanged so its PANOSE, Unicode-range,
    // metrics, weight and width penalties select the closest installed font.
    // Product defaults are applied only when a new document is created.
    return response;
  } catch (error) {
    if (url.pathname.startsWith('/fonts/')) {
      console.error('[onlyoffice-browser] Shared font fetch failed', error);
      const diagnostic = isLocalEditorHost && error instanceof Error ? `: ${error.message}` : '';
      return new Response(`Font metadata is unavailable${diagnostic}`, {
        status: 503,
        statusText: 'Font Metadata Unavailable',
      });
    }
    // Preserve the Worker redirect path as a safe fallback if version discovery
    // is temporarily unavailable. Font requests fail closed above because
    // loading different bytes for a font URL would corrupt glyph selection.
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
    // Retain the v2 message contract for older clients. Release v4 always
    // exposes the complete generated font catalog to native ONLYOFFICE lookup.
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

// A reusable editor-pool origin keeps host documents and workers origin-bound,
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
