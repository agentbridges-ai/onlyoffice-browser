const CANONICAL_HOST = 'onlyoffice.getpi.work';
const EDITOR_HOST_PATTERN = /^office-editor-[a-z0-9-]+\.getpi\.work$/;
const VERSION_QUERY = '__oobv';
const ASSET_REVISION_PATTERN = /^[a-f0-9]{16,64}$/;
const PRINT_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const RELEASE_PATH_PATTERN = /^\/r\/([^/]{1,384})\/(.+)$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;

type R2ObjectLike = {
  body?: ReadableStream;
  httpEtag: string;
  size: number;
  writeHttpMetadata(headers: Headers): void;
};

type ReleaseManifest = {
  version: 3;
  releaseId: string;
  assets: Array<{ path: string; sha256: string; mime: string; bytes: number }>;
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

export function shouldDisableResponseTransform(key: string, isolated: boolean): boolean {
  return isolated && (key.endsWith('.html') || key === 'index.html');
}

export function isAssetRevision(value: string | null): boolean {
  return ASSET_REVISION_PATTERN.test(value || '');
}

export function shouldShareAsset(pathname: string, destination: string | null): boolean {
  const normalizedDestination = destination?.trim().toLowerCase() || '';
  // Worker request metadata is not stable across browsers and intermediary
  // layers. Keep both worker entrypoints and their classic importScripts
  // dependency on the editor origin based on path, even if Sec-Fetch-Dest is
  // missing or reported as "script"/"empty".
  if (pathname.startsWith('/wasm/x2t/') && pathname.endsWith('.js')) return false;
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

export function resolveReleaseRequest(pathname: string): { releaseId: string; path: string } | null {
  const match = RELEASE_PATH_PATTERN.exec(pathname);
  if (!match) return null;
  let releaseId: string;
  try {
    releaseId = decodeURIComponent(match[1]);
  } catch {
    return null;
  }
  if (!RELEASE_ID_PATTERN.test(releaseId)) return null;
  const assetPath = resolveObjectKey(`/${match[2]}`);
  return assetPath ? { releaseId, path: assetPath } : null;
}

export function resolveEditorAssetRoute(pathname: string): { releaseId: string | null; pathname: string } {
  const releaseRequest = resolveReleaseRequest(pathname);
  return releaseRequest
    ? { releaseId: releaseRequest.releaseId, pathname: `/${releaseRequest.path}` }
    : { releaseId: null, pathname };
}

export function canonicalReleasePathname(pathname: string, releaseId: string): string {
  return resolveReleaseRequest(pathname) ? pathname : `/r/${releaseId}/${pathname.replace(/^\/+/, '')}`;
}

async function readJsonObject<T>(env: WorkerEnv, key: string): Promise<T | null> {
  const object = await env.ASSETS.get(key);
  if (!object?.body) return null;
  try {
    return (await new Response(object.body).json()) as T;
  } catch {
    return null;
  }
}

async function readReleaseManifest(
  env: WorkerEnv,
  ctx: WorkerExecutionContext,
  releaseId: string,
): Promise<ReleaseManifest | null> {
  const cache = (caches as CacheStorage & { default: Cache }).default;
  const cacheKey = new Request(`https://${CANONICAL_HOST}/releases/${encodeURIComponent(releaseId)}/manifest.json`);
  const cached = await cache.match(cacheKey);
  if (cached) return (await cached.json()) as ReleaseManifest;
  const object = await env.ASSETS.get(`releases/${releaseId}/manifest.json`);
  if (!object?.body) return null;
  const response = new Response(object.body, {
    headers: { 'Cache-Control': 'public, max-age=31536000, immutable' },
  });
  const cacheResponse = response.clone();
  try {
    const manifest = (await response.json()) as ReleaseManifest;
    ctx.waitUntil(cache.put(cacheKey, cacheResponse));
    return manifest;
  } catch {
    return null;
  }
}

async function resolveImmutableAsset(
  env: WorkerEnv,
  ctx: WorkerExecutionContext,
  request: { releaseId: string; path: string },
): Promise<{ key: string; version: string; publicPath: string; mime: string } | null> {
  const manifest = await readReleaseManifest(env, ctx, request.releaseId);
  if (manifest?.version !== 3 || manifest.releaseId !== request.releaseId || !Array.isArray(manifest.assets)) {
    return null;
  }
  const asset = manifest.assets.find((candidate) => candidate.path === request.path);
  if (!asset || !/^[a-f0-9]{64}$/.test(asset.sha256)) return null;
  return {
    key: `blobs/sha256/${asset.sha256}`,
    version: request.releaseId,
    publicPath: request.path,
    mime: asset.mime,
  };
}

async function stableReleaseId(env: WorkerEnv): Promise<string | null> {
  const channel = await readJsonObject<{ version?: number; releaseId?: string }>(env, 'channels/stable.json');
  return channel?.version === 1 && typeof channel.releaseId === 'string' && RELEASE_ID_PATTERN.test(channel.releaseId)
    ? channel.releaseId
    : null;
}

export function releaseIdFromReferrer(value: string | null): string | null {
  if (!value) return null;
  try {
    return resolveReleaseRequest(new URL(value).pathname)?.releaseId || null;
  } catch {
    return null;
  }
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
  const releaseRequest = resolveReleaseRequest(url.pathname);
  const immutableAsset = releaseRequest ? await resolveImmutableAsset(env, ctx, releaseRequest) : null;
  if (releaseRequest && !immutableAsset) return new Response('Release asset not found', { status: 404 });
  const key = immutableAsset?.key || resolveObjectKey(url.pathname);
  const publicPath = immutableAsset?.publicPath || key;
  if (!key || !publicPath) return new Response('Invalid asset path', { status: 400 });

  const versioned = isAssetRevision(url.searchParams.get(VERSION_QUERY));
  const immutable = Boolean(immutableAsset) || (!isolated && (versioned || key.startsWith('releases/')));
  const assetVersion = immutableAsset?.version || env.ASSET_VERSION;
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
    const headers = responseHeaders(metadata, immutable, isolated, publicPath, assetVersion, immutableAsset?.mime);
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
  const headers = responseHeaders(object, immutable, isolated, publicPath, assetVersion, immutableAsset?.mime);
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
  releaseMime?: string,
): Headers {
  const headers = new Headers();
  object.writeHttpMetadata(headers);
  applyReleaseMime(headers, releaseMime);
  headers.set('ETag', object.httpEtag);
  headers.set(
    'Cache-Control',
    key === 'channels/stable.json'
      ? 'no-store'
      : immutable
        ? 'public, max-age=31536000, immutable'
        : 'public, max-age=0, must-revalidate',
  );
  if (shouldDisableResponseTransform(key, isolated)) {
    headers.set('Cache-Control', `${headers.get('Cache-Control')}, no-transform`);
  }
  applySharedHeaders(headers);
  headers.set('X-OnlyOffice-Asset-Version', version);
  if (isolated && key === 'office-host.html') headers.set('Origin-Agent-Cluster', '?1');
  return headers;
}

export function applyReleaseMime(headers: Headers, releaseMime?: string): void {
  if (releaseMime) headers.set('Content-Type', releaseMime);
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
    const pinnedReleaseId = releaseIdFromReferrer(request.headers.get('referer'));
    const editorAssetRoute = resolveEditorAssetRoute(url.pathname);
    if (isolated && shouldShareAsset(editorAssetRoute.pathname, request.headers.get('sec-fetch-dest'))) {
      const releaseId = editorAssetRoute.releaseId || pinnedReleaseId || (await stableReleaseId(env));
      if (releaseId) {
        const canonical = new URL(url);
        canonical.hostname = CANONICAL_HOST;
        canonical.pathname = canonicalReleasePathname(url.pathname, releaseId);
        canonical.searchParams.delete(VERSION_QUERY);
        return Response.redirect(canonical.href, 307);
      }
      return Response.redirect(canonicalAssetUrl(url, env.ASSET_VERSION), 307);
    }
    if (isolated && pinnedReleaseId && !resolveReleaseRequest(url.pathname)) {
      url.pathname = `/r/${pinnedReleaseId}/${url.pathname.replace(/^\/+/, '')}`;
    }
    return serveAsset(request, env, ctx, url, isolated);
  },
};
