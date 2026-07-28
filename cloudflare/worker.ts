const CANONICAL_HOST = 'onlyoffice.getpi.work';
const EDITOR_HOST_PATTERN = /^office-[a-z0-9-]+\.getpi\.work$/;
const VERSION_QUERY = '__oobv';
const PRINT_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';

type R2ObjectLike = {
  body?: ReadableStream;
  httpEtag: string;
  size: number;
  writeHttpMetadata(headers: Headers): void;
};

type RuntimeBucket = {
  get(key: string, options?: { range?: Headers }): Promise<R2ObjectLike | null>;
  head(key: string): Promise<R2ObjectLike | null>;
};

export type WorkerEnv = {
  ASSETS: RuntimeBucket;
  ASSET_VERSION: string;
};

type WorkerExecutionContext = {
  waitUntil(promise: Promise<unknown>): void;
};

export function isOnlyOfficeHost(hostname: string): boolean {
  return hostname === CANONICAL_HOST || EDITOR_HOST_PATTERN.test(hostname);
}

export function isIsolatedEditorHost(hostname: string): boolean {
  return EDITOR_HOST_PATTERN.test(hostname);
}

export function shouldShareAsset(pathname: string, destination: string | null): boolean {
  const normalizedDestination = destination?.trim().toLowerCase() || '';
  if (
    normalizedDestination === 'document' ||
    normalizedDestination === 'iframe' ||
    normalizedDestination === 'worker' ||
    normalizedDestination === 'sharedworker' ||
    normalizedDestination === 'serviceworker'
  ) {
    return false;
  }
  if (pathname.startsWith('/assets/')) return false;
  if (
    pathname === '/office-host.html' ||
    pathname === '/reset.html' ||
    pathname === '/document_editor_service_worker.js' ||
    pathname === '/sw.js'
  ) {
    return false;
  }
  if (pathname.startsWith(PRINT_ROUTE_PREFIX)) return false;
  return true;
}

export function resolveObjectKey(pathname: string): string | null {
  let decoded: string;
  try {
    decoded = decodeURIComponent(pathname);
  } catch {
    return null;
  }
  const key = decoded === '/' ? 'index.html' : decoded.replace(/^\/+/, '');
  if (!key || key.includes('\0') || key.split('/').some((part) => part === '..')) return null;
  return key;
}

function canonicalAssetUrl(url: URL, version: string): string {
  const canonical = new URL(url);
  canonical.hostname = CANONICAL_HOST;
  canonical.searchParams.set(VERSION_QUERY, version);
  return canonical.href;
}

function applySharedHeaders(headers: Headers): void {
  headers.set('Access-Control-Allow-Origin', '*');
  headers.set('Access-Control-Expose-Headers', 'X-OnlyOffice-Asset-Version');
  headers.set('Cross-Origin-Resource-Policy', 'cross-origin');
  headers.set('Timing-Allow-Origin', '*');
}

function matchesEtag(request: Request, etag: string): boolean {
  return (request.headers.get('if-none-match') || '')
    .split(',')
    .map((value) => value.trim())
    .some((value) => value === '*' || value === etag);
}

function parseSingleRange(value: string | null, size: number): { start: number; end: number } | null {
  const match = /^bytes=(\d*)-(\d*)$/.exec(value || '');
  if (!match || size <= 0 || (!match[1] && !match[2])) return null;
  if (!match[1]) {
    const suffix = Number(match[2]);
    if (!Number.isSafeInteger(suffix) || suffix <= 0) return null;
    return { start: Math.max(0, size - suffix), end: size - 1 };
  }
  const start = Number(match[1]);
  const requestedEnd = match[2] ? Number(match[2]) : size - 1;
  if (!Number.isSafeInteger(start) || !Number.isSafeInteger(requestedEnd) || start < 0 || start >= size) {
    return null;
  }
  const end = Math.min(requestedEnd, size - 1);
  return end < start ? null : { start, end };
}

async function serveAsset(
  request: Request,
  env: WorkerEnv,
  ctx: WorkerExecutionContext,
  url: URL,
  isolated: boolean,
): Promise<Response> {
  const key = resolveObjectKey(url.pathname);
  if (!key) return new Response('Invalid asset path', { status: 400 });

  const versioned = url.searchParams.get(VERSION_QUERY) === env.ASSET_VERSION;
  const immutable = !isolated && versioned;
  const cache = (caches as CacheStorage & { default: Cache }).default;
  const cacheKey = new Request(url.href, { method: 'GET' });
  const rangeHeader = request.headers.get('range');
  const cacheable = request.method === 'GET' && immutable && !rangeHeader;
  if (cacheable) {
    const cached = await cache.match(cacheKey);
    if (cached) return cached;
  }

  if (request.method === 'HEAD') {
    const metadata = await env.ASSETS.head(key);
    if (!metadata) return new Response('Not found', { status: 404 });
    const headers = responseHeaders(metadata, immutable, isolated, key, env.ASSET_VERSION);
    if (matchesEtag(request, metadata.httpEtag)) return new Response(null, { status: 304, headers });
    headers.set('Content-Length', String(metadata.size));
    return new Response(null, { status: 200, headers });
  }

  const range = rangeHeader ? parseSingleRange(rangeHeader, (await env.ASSETS.head(key))?.size ?? 0) : null;
  if (rangeHeader && !range) {
    const metadata = await env.ASSETS.head(key);
    return new Response(null, {
      status: 416,
      headers: metadata ? { 'Content-Range': `bytes */${metadata.size}` } : undefined,
    });
  }

  const object = await env.ASSETS.get(
    key,
    range ? { range: new Headers({ Range: `bytes=${range.start}-${range.end}` }) } : undefined,
  );
  if (!object?.body) return new Response('Not found', { status: 404 });
  const headers = responseHeaders(object, immutable, isolated, key, env.ASSET_VERSION);
  if (matchesEtag(request, object.httpEtag)) return new Response(null, { status: 304, headers });

  let status = 200;
  if (range) {
    status = 206;
    headers.set('Accept-Ranges', 'bytes');
    headers.set('Content-Length', String(range.end - range.start + 1));
    headers.set('Content-Range', `bytes ${range.start}-${range.end}/${object.size}`);
  } else {
    headers.set('Content-Length', String(object.size));
  }
  const response = new Response(object.body, { status, headers });
  if (cacheable) ctx.waitUntil(cache.put(cacheKey, response.clone()));
  return response;
}

function responseHeaders(
  object: R2ObjectLike,
  immutable: boolean,
  isolated: boolean,
  key: string,
  version: string,
): Headers {
  const headers = new Headers();
  object.writeHttpMetadata(headers);
  headers.set('ETag', object.httpEtag);
  headers.set(
    'Cache-Control',
    immutable ? 'public, max-age=31536000, immutable' : 'public, max-age=0, must-revalidate',
  );
  applySharedHeaders(headers);
  headers.set('X-OnlyOffice-Asset-Version', version);
  if (isolated && key === 'office-host.html') headers.set('Origin-Agent-Cluster', '?1');
  return headers;
}

export default {
  async fetch(request: Request, env: WorkerEnv, ctx: WorkerExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    const hostname = url.hostname.toLowerCase();
    if (!isOnlyOfficeHost(hostname)) return new Response('Unknown host', { status: 404 });

    if (request.method === 'OPTIONS') {
      const headers = new Headers({
        'Access-Control-Allow-Methods': 'GET, HEAD, OPTIONS',
        'Access-Control-Allow-Headers': request.headers.get('access-control-request-headers') || '*',
        'Access-Control-Max-Age': '86400',
      });
      applySharedHeaders(headers);
      return new Response(null, { status: 204, headers });
    }
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      return new Response('Method not allowed', { status: 405, headers: { Allow: 'GET, HEAD, OPTIONS' } });
    }

    const isolated = isIsolatedEditorHost(hostname);
    if (
      isolated &&
      shouldShareAsset(url.pathname, request.headers.get('sec-fetch-dest'))
    ) {
      return Response.redirect(canonicalAssetUrl(url, env.ASSET_VERSION), 307);
    }
    return serveAsset(request, env, ctx, url, isolated);
  },
};
