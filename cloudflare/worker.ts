import { isProductionOfficeEditorHostname } from '../src/lib/office-origin-pool';

const CANONICAL_HOST = 'onlyoffice.getpi.work';
const LEGACY_EDITOR_HOST_PATTERN = /^office-editor-[a-z0-9-]+\.getpi\.work$/;
const LOCAL_PWA_HOST = 'onlyoffice.localhost';
const LOCAL_CANONICAL_HOST = 'assets.office.localhost';
const LOCAL_MATRIX_HOST_HEADER = 'x-onlyoffice-matrix-host';
const LOCAL_MATRIX_CONTROL_TOKEN_HEADER = 'x-onlyoffice-matrix-control-token';
const MAX_EDGE_CACHE_OBJECT_BYTES = 8 * 1024 * 1024;
const LOCAL_LEGACY_EDITOR_HOST_PATTERN = /^office-editor-([a-z0-9-]+)\.localhost$/;
const LOCAL_LEGACY_ISOLATED_EDITOR_HOST_PATTERN = /^host-office-editor-([a-z0-9-]+)\.office\.localhost$/;
const LOCAL_EDITOR_HOST_PATTERN =
  /^(aries|taurus|gemini|cancer|leo|virgo|libra|scorpio|sagittarius|capricorn|aquarius|pisces)\.localhost$/;
const LOCAL_ISOLATED_EDITOR_HOST_PATTERN =
  /^host-(aries|taurus|gemini|cancer|leo|virgo|libra|scorpio|sagittarius|capricorn|aquarius|pisces)\.office\.localhost$/;
const VERSION_QUERY = '__oobv';
const ASSET_REVISION_PATTERN = /^[a-f0-9]{16,64}$/;
const PRINT_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const RELEASE_PATH_PATTERN = /^\/r\/([^/]{1,384})\/(.+)$/;
const RELEASE_PACKAGE_PATH_PATTERN = /^\/p\/([^/]{1,384})\/office-resources\.oobpack$/;
const RELEASE_SEGMENT_PATH_PATTERN = /^\/segments\/sha256\/([a-f0-9]{64})$/;
const RELEASE_CONTENT_OBJECT_PATH_PATTERN = /^\/objects\/([^/]{1,384})\/sha256\/([a-f0-9]{64})$/;
const EDITOR_SHELL_ASSET_PATH_PATTERN = /^\/assets\/[a-zA-Z0-9._+-]+\.(?:css|js)$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const FASTCDC_POLICY_ID = 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0';
const OFFICE_EDITOR_SLOTS = [
  'aries',
  'taurus',
  'gemini',
  'cancer',
  'leo',
  'virgo',
  'libra',
  'scorpio',
  'sagittarius',
  'capricorn',
  'aquarius',
  'pisces',
] as const;
export type LocalMatrixCounter = {
  workerRequests: number;
  cacheHits: number;
  r2Heads: number;
  r2Gets: number;
  declaredBytes: number;
  actualBytes: number;
  r2Bytes: number;
  completed: number;
  aborted: number;
  failed: number;
  stalled: number;
  statuses: Record<string, number>;
};
const localMatrixSegmentCounters = new Map<string, LocalMatrixCounter>();
const localMatrixObjectCounters = new Map<string, LocalMatrixCounter>();
const localMatrixRouteCounters = new Map<string, LocalMatrixCounter>();
const localMatrixStalledObjects = new Set<string>();

type R2ObjectLike = {
  body?: ReadableStream;
  httpEtag: string;
  size: number;
  writeHttpMetadata(headers: Headers): void;
};

export type ReleaseManifest = {
  version: 3 | 4 | 5;
  releaseId: string;
  contentProtocol?: {
    version?: number;
    digest?: string;
    cacheKeyFormat?: string;
    storageSetSha256?: string;
    fastcdcPolicyId?: string;
  };
  assets: Array<{
    path: string;
    sha256: string;
    mime: string;
    bytes: number;
    profile?: string;
    chunk?: string;
    packageOffset?: number;
    representations?: {
      whole?: { sha256?: string; bytes?: number };
      fastcdc?: {
        algorithm?: string;
        minBytes?: number;
        averageBytes?: number;
        maxBytes?: number;
        normalization?: number;
        seed?: number;
        chunks?: Array<{ offset?: number; bytes?: number; sha256?: string }>;
      };
    };
  }>;
  package?: {
    format: string;
    path: string;
    bytes: number;
    sha256: string;
    segments?: Array<{ id: string; offset: number; bytes: number; sha256: string }>;
  };
};

type ImmutableAsset = {
  key: string;
  version: string;
  publicPath: string;
  mime: string;
  objectOffset?: number;
  publicSize?: number;
  contentSha256?: string;
  requireExactObjectSize?: boolean;
};

type RuntimeBucket = {
  get(key: string, options?: { range?: Headers }): Promise<R2ObjectLike | null>;
  head(key: string): Promise<R2ObjectLike | null>;
  put?(
    key: string,
    value: ReadableStream | ArrayBuffer | ArrayBufferView | string | null,
    options?: { httpMetadata?: { contentType?: string; cacheControl?: string } },
  ): Promise<unknown>;
};

export type WorkerEnv = {
  ASSETS: RuntimeBucket;
  ASSET_VERSION: string;
  LOCAL_MATRIX_MODE?: string;
  LOCAL_MATRIX_PORT?: string;
  LOCAL_MATRIX_CONTROL_TOKEN?: string;
  LOCAL_MATRIX_R2_DELAY_MS?: string;
  LOCAL_MATRIX_R2_STALL_KEY?: string;
  LOCAL_MATRIX_R2_STALL_AFTER_BYTES?: string;
  LOCAL_MATRIX_R2_STALL_MS?: string;
  LOCAL_MATRIX_R2_STALL_ONCE?: string;
};

type WorkerExecutionContext = {
  waitUntil(promise: Promise<unknown>): void;
};

export function isOnlyOfficeHost(hostname: string): boolean {
  return (
    hostname === CANONICAL_HOST ||
    isProductionOfficeEditorHostname(hostname) ||
    LEGACY_EDITOR_HOST_PATTERN.test(hostname)
  );
}

export function isIsolatedEditorHost(hostname: string): boolean {
  return isProductionOfficeEditorHostname(hostname) || LEGACY_EDITOR_HOST_PATTERN.test(hostname);
}

export function resolveRuntimeHost(
  hostname: string,
  localMatrixMode = false,
): { logicalHostname: string; canonicalHostname: string } {
  if (!localMatrixMode) {
    return { logicalHostname: hostname, canonicalHostname: CANONICAL_HOST };
  }
  if (hostname === LOCAL_PWA_HOST || hostname === LOCAL_CANONICAL_HOST) {
    return { logicalHostname: CANONICAL_HOST, canonicalHostname: LOCAL_CANONICAL_HOST };
  }
  const editor = LOCAL_EDITOR_HOST_PATTERN.exec(hostname) || LOCAL_ISOLATED_EDITOR_HOST_PATTERN.exec(hostname);
  if (editor) {
    return {
      logicalHostname: `${editor[1]}.getpi.work`,
      canonicalHostname: LOCAL_CANONICAL_HOST,
    };
  }
  const legacyEditor =
    LOCAL_LEGACY_EDITOR_HOST_PATTERN.exec(hostname) || LOCAL_LEGACY_ISOLATED_EDITOR_HOST_PATTERN.exec(hostname);
  if (legacyEditor) {
    return {
      logicalHostname: `office-editor-${legacyEditor[1]}.getpi.work`,
      canonicalHostname: LOCAL_CANONICAL_HOST,
    };
  }
  return { logicalHostname: hostname, canonicalHostname: LOCAL_CANONICAL_HOST };
}

function parseHostnameAuthority(value: string | null): string | null {
  if (!value || value.length > 253 || /[/\\@]/.test(value)) return null;
  try {
    const parsed = new URL(`http://${value}`);
    return parsed.username || parsed.password || parsed.pathname !== '/' || parsed.search || parsed.hash
      ? null
      : parsed.hostname.toLowerCase();
  } catch {
    return null;
  }
}

/**
 * Wrangler local development rewrites request.url to the configured production
 * route. The matrix gateway records the browser-visible Host in a private
 * header so the Worker still executes the real canonical/editor isolation
 * branch. Production never trusts this header.
 */
export function resolveRequestHostname(request: Request, localMatrixMode = false): string | null {
  const urlHostname = new URL(request.url).hostname.toLowerCase();
  if (!localMatrixMode) return urlHostname;
  const forwarded = parseHostnameAuthority(request.headers.get(LOCAL_MATRIX_HOST_HEADER));
  if (forwarded) return forwarded;
  // Direct *.localhost requests remain useful for unit tests. A Wrangler
  // request rewritten to a production route must pass through the gateway.
  return urlHostname.endsWith('.localhost') ? urlHostname : null;
}

export function shouldDisableResponseTransform(key: string, isolated: boolean): boolean {
  return isolated && (key.endsWith('.html') || key === 'index.html');
}

export function isAssetRevision(value: string | null): boolean {
  return ASSET_REVISION_PATTERN.test(value || '');
}

export function shouldPopulateEdgeCache(options: {
  method: string;
  immutable: boolean;
  hasRange: boolean;
  publicSize?: number;
  unboundedContentSegment?: boolean;
}): boolean {
  if (options.method !== 'GET' || !options.immutable || options.hasRange || options.unboundedContentSegment) {
    return false;
  }
  return options.publicSize === undefined || options.publicSize <= MAX_EDGE_CACHE_OBJECT_BYTES;
}

export function shouldShareAsset(pathname: string, destination: string | null): boolean {
  const normalizedDestination = destination?.trim().toLowerCase() || '';
  // Worker request metadata is not stable across browsers and intermediary
  // layers. Keep both worker entrypoints and their classic importScripts
  // dependency on the editor origin based on path, even if Sec-Fetch-Dest is
  // missing or reported as "script"/"empty".
  if (pathname.startsWith('/wasm/x2t/') && pathname.endsWith('.js')) return false;
  // The immutable Host and shell-prime HTML import a small, bounded set of
  // Vite chunks from /assets. They must remain same-origin to satisfy the
  // editor CSP and are the only non-document dependencies copied into each
  // editor origin's <=16 MiB shell cache.
  if (EDITOR_SHELL_ASSET_PATH_PATTERN.test(pathname)) return false;
  if (
    normalizedDestination === 'document' ||
    normalizedDestination === 'iframe' ||
    normalizedDestination === 'worker' ||
    normalizedDestination === 'sharedworker' ||
    normalizedDestination === 'serviceworker'
  ) {
    return false;
  }
  if (
    pathname === '/office-host.html' ||
    pathname === '/editor-shell-prime.html' ||
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
  // A URL.pathname assignment or a subsequent request parse can decode dot
  // segments again. Reject residual percent escapes after the single
  // canonical decode so double/triple encoded separators can never escape a
  // release prefix.
  if (/%[0-9a-f]{2}/i.test(decoded)) return null;
  const key = decoded === '/' ? 'index.html' : decoded.replace(/^\/+/, '').replace(/\/{2,}/g, '/');
  if (
    !key ||
    key.includes('\\') ||
    Array.from(key).some((character) => {
      const code = character.charCodeAt(0);
      return code === 0 || code <= 0x1f || code === 0x7f;
    }) ||
    key.split('/').some((part) => part === '.' || part === '..')
  ) {
    return null;
  }
  return key;
}

function canonicalAssetUrl(url: URL, version: string, canonicalHostname = CANONICAL_HOST): string {
  const canonical = new URL(url);
  canonical.hostname = canonicalHostname;
  canonical.searchParams.set(VERSION_QUERY, version);
  return canonical.href;
}

function isCanonicalReleaseMetadataPath(pathname: string): boolean {
  return /^\/(?:channels|releases|segments|objects|blobs|packages)\//.test(pathname) || pathname.startsWith('/p/');
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

export function resolveReleasePackageRequest(pathname: string): { releaseId: string } | null {
  const match = RELEASE_PACKAGE_PATH_PATTERN.exec(pathname);
  if (!match) return null;
  let releaseId: string;
  try {
    releaseId = decodeURIComponent(match[1]);
  } catch {
    return null;
  }
  return RELEASE_ID_PATTERN.test(releaseId) ? { releaseId } : null;
}

export function resolveContentSegmentRequest(pathname: string): { sha256: string } | null {
  const match = RELEASE_SEGMENT_PATH_PATTERN.exec(pathname);
  return match ? { sha256: match[1] } : null;
}

export function resolveReleaseContentObjectRequest(pathname: string): { releaseId: string; sha256: string } | null {
  const match = RELEASE_CONTENT_OBJECT_PATH_PATTERN.exec(pathname);
  if (!match) return null;
  let releaseId: string;
  try {
    releaseId = decodeURIComponent(match[1]);
  } catch {
    return null;
  }
  return RELEASE_ID_PATTERN.test(releaseId) ? { releaseId, sha256: match[2] } : null;
}

export function resolveEditorAssetRoute(pathname: string): { releaseId: string | null; pathname: string } {
  const releaseRequest = resolveReleaseRequest(pathname);
  return releaseRequest
    ? { releaseId: releaseRequest.releaseId, pathname: `/${releaseRequest.path}` }
    : { releaseId: null, pathname };
}

export function canonicalReleasePathname(pathname: string, releaseId: string): string | null {
  if (resolveReleaseRequest(pathname)) return pathname;
  const key = resolveObjectKey(pathname);
  return key ? `/r/${releaseId}/${key}` : null;
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
): Promise<ImmutableAsset | null> {
  const manifest = await readReleaseManifest(env, ctx, request.releaseId);
  if (
    (manifest?.version !== 3 && manifest?.version !== 4 && manifest?.version !== 5) ||
    manifest.releaseId !== request.releaseId ||
    !Array.isArray(manifest.assets)
  ) {
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

async function resolveImmutablePackage(
  env: WorkerEnv,
  ctx: WorkerExecutionContext,
  request: { releaseId: string },
  segmentId: string | null,
): Promise<ImmutableAsset | null> {
  const manifest = await readReleaseManifest(env, ctx, request.releaseId);
  const pack = manifest?.package;
  if (
    (manifest?.version !== 4 && manifest?.version !== 5) ||
    manifest.releaseId !== request.releaseId ||
    pack?.format !== 'onlyoffice-pack-v1' ||
    pack.path !== 'office-resources.oobpack' ||
    !/^[a-f0-9]{64}$/.test(pack.sha256) ||
    !Number.isSafeInteger(pack.bytes) ||
    pack.bytes <= 0 ||
    (segmentId !== null && !Array.isArray(pack.segments))
  ) {
    return null;
  }
  const segment = segmentId === null ? null : pack.segments?.find((candidate) => candidate.id === segmentId);
  if (
    segmentId !== null &&
    (!segment ||
      !Number.isSafeInteger(segment.offset) ||
      segment.offset < 0 ||
      !Number.isSafeInteger(segment.bytes) ||
      segment.bytes <= 0 ||
      segment.offset + segment.bytes > pack.bytes ||
      !/^[a-f0-9]{64}$/.test(segment.sha256))
  ) {
    return null;
  }
  return {
    key: `packages/sha256/${pack.sha256}.oobpack`,
    version: request.releaseId,
    publicPath: pack.path,
    mime: 'application/vnd.onlyoffice.browser-pack',
    ...(segment ? { objectOffset: segment.offset, publicSize: segment.bytes } : {}),
  };
}

function isSafeReleasePath(value: unknown): value is string {
  if (
    typeof value !== 'string' ||
    !value ||
    value.startsWith('/') ||
    value.includes('\\') ||
    value.includes('?') ||
    value.includes('#') ||
    Array.from(value).some((character) => {
      const code = character.charCodeAt(0);
      return code <= 0x1f || code === 0x7f;
    })
  ) {
    return false;
  }
  try {
    const decoded = decodeURIComponent(value);
    return (
      !decoded.startsWith('/') &&
      !decoded.includes('\\') &&
      decoded.split('/').every((segment) => segment && segment !== '.' && segment !== '..')
    );
  } catch {
    return false;
  }
}

type ReleaseContentObjectTarget = {
  sha256: string;
  bytes: number;
  key: string;
};

function resolveReleaseContentObjectTarget(
  manifest: ReleaseManifest | null,
  request: { releaseId: string; sha256: string },
): ReleaseContentObjectTarget | null {
  if (
    manifest?.version !== 5 ||
    manifest.releaseId !== request.releaseId ||
    manifest.contentProtocol?.version !== 1 ||
    manifest.contentProtocol.digest !== 'sha256' ||
    manifest.contentProtocol.cacheKeyFormat !== 'canonical-sha256-v1' ||
    !SHA256_PATTERN.test(manifest.contentProtocol.storageSetSha256 || '') ||
    manifest.contentProtocol.fastcdcPolicyId !== FASTCDC_POLICY_ID ||
    !Array.isArray(manifest.assets)
  ) {
    return null;
  }

  const pack = manifest.package;
  if (
    pack?.format !== 'onlyoffice-pack-v1' ||
    pack.path !== 'office-resources.oobpack' ||
    !Number.isSafeInteger(pack.bytes) ||
    pack.bytes <= 0 ||
    !SHA256_PATTERN.test(pack.sha256 || '') ||
    !Array.isArray(pack.segments) ||
    pack.segments.length === 0
  ) {
    return null;
  }

  const objectSizes = new Map<string, number>();
  const blobObjects = new Set<string>();
  const packageObjects = new Set<string>();
  const addObject = (
    sha256: unknown,
    bytes: unknown,
    storage: 'blob' | 'package' | null,
    allowEmpty = false,
  ): boolean => {
    if (
      typeof sha256 !== 'string' ||
      !SHA256_PATTERN.test(sha256) ||
      !Number.isSafeInteger(bytes) ||
      (allowEmpty ? (bytes as number) < 0 : (bytes as number) <= 0)
    ) {
      return false;
    }
    const existingBytes = objectSizes.get(sha256);
    if (existingBytes !== undefined && existingBytes !== bytes) return false;
    objectSizes.set(sha256, bytes as number);
    if (storage === 'blob') blobObjects.add(sha256);
    else if (storage === 'package') packageObjects.add(sha256);
    return true;
  };

  let expectedPackageOffset = 0;
  for (const segment of pack.segments) {
    if (
      segment?.id !== segment?.sha256 ||
      segment?.offset !== expectedPackageOffset ||
      !addObject(segment?.sha256, segment?.bytes, 'package')
    ) {
      return null;
    }
    expectedPackageOffset += segment.bytes;
  }
  if (expectedPackageOffset !== pack.bytes) return null;

  const assetPaths = new Set<string>();
  for (const asset of manifest.assets) {
    const whole = asset?.representations?.whole;
    if (
      !asset ||
      !isSafeReleasePath(asset.path) ||
      assetPaths.has(asset.path) ||
      !Number.isSafeInteger(asset.bytes) ||
      asset.bytes < 0 ||
      !SHA256_PATTERN.test(asset.sha256 || '') ||
      typeof asset.mime !== 'string' ||
      !asset.mime ||
      !Number.isSafeInteger(asset.packageOffset) ||
      (asset.packageOffset as number) < 0 ||
      (asset.packageOffset as number) + asset.bytes > pack.bytes ||
      !whole ||
      whole.sha256 !== asset.sha256 ||
      whole.bytes !== asset.bytes ||
      !addObject(whole.sha256, whole.bytes, 'blob', true)
    ) {
      return null;
    }
    assetPaths.add(asset.path);

    const fastcdc = asset.representations?.fastcdc;
    if (!fastcdc) continue;
    if (
      fastcdc.algorithm !== 'fastcdc-v2020' ||
      fastcdc.minBytes !== 65_536 ||
      fastcdc.averageBytes !== 262_144 ||
      fastcdc.maxBytes !== 1_048_576 ||
      fastcdc.normalization !== 1 ||
      fastcdc.seed !== 0 ||
      !Array.isArray(fastcdc.chunks) ||
      fastcdc.chunks.length === 0
    ) {
      return null;
    }
    let expectedOffset = 0;
    for (const chunk of fastcdc.chunks) {
      if (
        chunk?.offset !== expectedOffset ||
        !Number.isSafeInteger(chunk.bytes) ||
        (chunk.bytes as number) <= 0 ||
        expectedOffset + (chunk.bytes as number) > asset.bytes ||
        !addObject(chunk.sha256, chunk.bytes, 'blob')
      ) {
        return null;
      }
      expectedOffset += chunk.bytes as number;
    }
    if (expectedOffset !== asset.bytes) return null;
  }

  const bytes = objectSizes.get(request.sha256);
  if (bytes === undefined) return null;
  if (blobObjects.has(request.sha256)) {
    return {
      sha256: request.sha256,
      bytes,
      key: `blobs/sha256/${request.sha256}`,
    };
  }
  if (packageObjects.has(request.sha256)) {
    return {
      sha256: request.sha256,
      bytes,
      key: `segments/sha256/${request.sha256}`,
    };
  }
  return null;
}

export function resolveReleaseContentObject(
  manifest: ReleaseManifest | null,
  request: { releaseId: string; sha256: string },
): { sha256: string; bytes: number } | null {
  const target = resolveReleaseContentObjectTarget(manifest, request);
  return target ? { sha256: target.sha256, bytes: target.bytes } : null;
}

async function resolveImmutableContentObject(
  env: WorkerEnv,
  ctx: WorkerExecutionContext,
  request: { releaseId: string; sha256: string },
): Promise<ImmutableAsset | null> {
  const contentObject = resolveReleaseContentObjectTarget(
    await readReleaseManifest(env, ctx, request.releaseId),
    request,
  );
  if (!contentObject) return null;
  return {
    key: contentObject.key,
    version: request.releaseId,
    publicPath: `objects/${request.releaseId}/sha256/${contentObject.sha256}`,
    mime: 'application/octet-stream',
    publicSize: contentObject.bytes,
    contentSha256: contentObject.sha256,
    requireExactObjectSize: true,
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
  headers.set(
    'Access-Control-Expose-Headers',
    'Accept-Ranges, Content-Length, Content-Range, ETag, X-Content-SHA256, X-OnlyOffice-Asset-Version',
  );
  headers.set('Cross-Origin-Resource-Policy', 'cross-origin');
  headers.set('Timing-Allow-Origin', '*');
}

function sharedRedirect(url: string, status: 307 | 308 = 307): Response {
  const location = new URL(url).href;
  const headers = new Headers({ Location: location });
  applySharedHeaders(headers);
  return new Response(null, { status, headers });
}

function assetError(body: BodyInit | null, status: number, initialHeaders?: HeadersInit): Response {
  const headers = new Headers(initialHeaders);
  applySharedHeaders(headers);
  return new Response(body, { status, headers });
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

function localMatrixCounter(counters: Map<string, LocalMatrixCounter>, key: string) {
  let counter = counters.get(key);
  if (!counter) {
    counter = {
      workerRequests: 0,
      cacheHits: 0,
      r2Heads: 0,
      r2Gets: 0,
      declaredBytes: 0,
      actualBytes: 0,
      r2Bytes: 0,
      completed: 0,
      aborted: 0,
      failed: 0,
      stalled: 0,
      statuses: {},
    };
    counters.set(key, counter);
  }
  return counter;
}

function uniqueMatrixCounters(counters: Array<LocalMatrixCounter | null>): LocalMatrixCounter[] {
  return [...new Set(counters.filter((counter): counter is LocalMatrixCounter => counter !== null))];
}

function recordMatrixStatus(counters: LocalMatrixCounter[], status: number): void {
  const key = String(status);
  for (const counter of counters) counter.statuses[key] = (counter.statuses[key] || 0) + 1;
}

function completeMatrixRequest(counters: LocalMatrixCounter[], status: number): void {
  recordMatrixStatus(counters, status);
  for (const counter of counters) counter.completed += 1;
}

function recordMatrixBytes(counters: LocalMatrixCounter[], bytes: number): void {
  for (const counter of counters) {
    counter.actualBytes += bytes;
    counter.r2Bytes += bytes;
  }
}

function countR2Body(
  body: ReadableStream,
  counters: LocalMatrixCounter[],
  status: number,
  expectedBytes: number,
): ReadableStream<Uint8Array> {
  const FixedLengthStreamConstructor = (
    globalThis as typeof globalThis & {
      FixedLengthStream?: new (length: number) => {
        readable: ReadableStream<Uint8Array>;
        writable: WritableStream<Uint8Array>;
      };
    }
  ).FixedLengthStream;
  if (FixedLengthStreamConstructor) {
    const fixed = new FixedLengthStreamConstructor(expectedBytes);
    // Keep the source-to-response pump as one stream pipeline. A detached
    // reader/writer task can outlive the Worker response and intermittently
    // reset large HTTP/2 streams before the FixedLengthStream reaches its
    // declared byte count. pipeTo() keeps backpressure and the response body
    // lifecycle connected while still allowing the matrix counters to observe
    // every R2 chunk.
    const counted = body.pipeThrough(
      new TransformStream<Uint8Array, Uint8Array>({
        transform(value, controller) {
          const chunk = value instanceof Uint8Array ? value : new Uint8Array(value as unknown as ArrayBuffer);
          recordMatrixBytes(counters, chunk.byteLength);
          controller.enqueue(chunk);
        },
      }),
    );
    void counted.pipeTo(fixed.writable).then(
      () => completeMatrixRequest(counters, status),
      () => {
        recordMatrixStatus(counters, status);
        for (const counter of counters) counter.aborted += 1;
      },
    );
    return fixed.readable;
  }
  const reader = body.getReader();
  let settled = false;
  const complete = () => {
    if (settled) return;
    settled = true;
    completeMatrixRequest(counters, status);
  };
  const fail = () => {
    if (settled) return;
    settled = true;
    recordMatrixStatus(counters, status);
    for (const counter of counters) counter.failed += 1;
  };
  return new ReadableStream<Uint8Array>({
    async pull(controller) {
      try {
        const result = await reader.read();
        if (result.done) {
          complete();
          controller.close();
          return;
        }
        const chunk =
          result.value instanceof Uint8Array ? result.value : new Uint8Array(result.value as unknown as ArrayBuffer);
        recordMatrixBytes(counters, chunk.byteLength);
        controller.enqueue(chunk);
      } catch (error) {
        fail();
        controller.error(error);
      }
    },
    async cancel(reason) {
      if (!settled) {
        settled = true;
        recordMatrixStatus(counters, status);
        for (const counter of counters) counter.aborted += 1;
      }
      await reader.cancel(reason);
    },
  });
}

function maybeStallLocalMatrixBody(
  body: ReadableStream,
  env: WorkerEnv,
  objectKey: string,
  counters: LocalMatrixCounter[],
): ReadableStream {
  if (env.LOCAL_MATRIX_MODE !== '1') return body;
  const configuredKey = env.LOCAL_MATRIX_R2_STALL_KEY?.trim() || '';
  const keyMatches =
    configuredKey === '*' || configuredKey === objectKey || objectKey.endsWith(`/sha256/${configuredKey}`);
  if (!configuredKey || !keyMatches) return body;
  const afterBytes = Number.parseInt(env.LOCAL_MATRIX_R2_STALL_AFTER_BYTES || '1', 10);
  const stallMs = Number.parseInt(env.LOCAL_MATRIX_R2_STALL_MS || '0', 10);
  if (!Number.isSafeInteger(afterBytes) || afterBytes <= 0 || !Number.isSafeInteger(stallMs) || stallMs <= 0)
    return body;
  const once = env.LOCAL_MATRIX_R2_STALL_ONCE === '1';
  const stallToken = `${objectKey}:${configuredKey}:${afterBytes}:${stallMs}`;
  if (once && localMatrixStalledObjects.has(stallToken)) return body;
  if (once) localMatrixStalledObjects.add(stallToken);
  for (const counter of counters) counter.stalled += 1;
  const reader = body.getReader();
  let seenBytes = 0;
  let stalled = false;
  return new ReadableStream({
    async pull(controller) {
      const result = await reader.read();
      if (result.done) {
        controller.close();
        return;
      }
      const value = result.value instanceof Uint8Array ? result.value : new Uint8Array(result.value as ArrayBuffer);
      seenBytes += value.byteLength;
      controller.enqueue(value);
      if (!stalled && seenBytes >= afterBytes) {
        stalled = true;
        await new Promise((resolve) => setTimeout(resolve, stallMs));
      }
    },
    cancel(reason) {
      return reader.cancel(reason);
    },
  });
}

async function serveAsset(
  request: Request,
  env: WorkerEnv,
  ctx: WorkerExecutionContext,
  url: URL,
  isolated: boolean,
): Promise<Response> {
  const releaseRequest = resolveReleaseRequest(url.pathname);
  const packageRequest = resolveReleasePackageRequest(url.pathname);
  const contentSegmentRequest = resolveContentSegmentRequest(url.pathname);
  const contentObjectRequest = resolveReleaseContentObjectRequest(url.pathname);
  const routeCounter =
    env.LOCAL_MATRIX_MODE === '1' && (releaseRequest || packageRequest || contentSegmentRequest || contentObjectRequest)
      ? localMatrixCounter(localMatrixRouteCounters, `${request.method} ${url.pathname}${url.search}`)
      : null;
  const matrixCounter =
    env.LOCAL_MATRIX_MODE !== '1'
      ? null
      : contentSegmentRequest
        ? localMatrixCounter(localMatrixSegmentCounters, contentSegmentRequest.sha256)
        : contentObjectRequest
          ? localMatrixCounter(
              localMatrixObjectCounters,
              `${contentObjectRequest.releaseId}:${contentObjectRequest.sha256}`,
            )
          : null;
  const matrixCounters = uniqueMatrixCounters([matrixCounter, routeCounter]);
  for (const counter of matrixCounters) counter.workerRequests += 1;
  const packageSegmentId = packageRequest ? url.searchParams.get('segment') : null;
  const immutableAsset = releaseRequest
    ? await resolveImmutableAsset(env, ctx, releaseRequest)
    : packageRequest
      ? await resolveImmutablePackage(env, ctx, packageRequest, packageSegmentId)
      : contentSegmentRequest
        ? {
            key: `segments/sha256/${contentSegmentRequest.sha256}`,
            version: contentSegmentRequest.sha256,
            publicPath: `segments/sha256/${contentSegmentRequest.sha256}`,
            mime: 'application/vnd.onlyoffice.browser-pack-segment',
          }
        : contentObjectRequest
          ? await resolveImmutableContentObject(env, ctx, contentObjectRequest)
          : null;
  if ((releaseRequest || packageRequest || contentSegmentRequest || contentObjectRequest) && !immutableAsset) {
    const body = releaseRequest
      ? 'Release asset not found'
      : packageRequest
        ? 'Release package not found'
        : contentObjectRequest
          ? 'Release content object not found'
          : 'Release segment not found';
    completeMatrixRequest(matrixCounters, 404);
    return assetError(body, 404);
  }
  const key = immutableAsset?.key || resolveObjectKey(url.pathname);
  const publicPath = immutableAsset?.publicPath || key;
  if (!key || !publicPath) {
    completeMatrixRequest(matrixCounters, 400);
    return assetError('Invalid asset path', 400);
  }

  const versioned = isAssetRevision(url.searchParams.get(VERSION_QUERY));
  const immutable = Boolean(immutableAsset) || (!isolated && (versioned || key.startsWith('releases/')));
  const assetVersion = immutableAsset?.version || env.ASSET_VERSION;
  const cache = (caches as CacheStorage & { default: Cache }).default;
  const cacheKeyUrl = new URL(url);
  if (env.LOCAL_MATRIX_MODE === '1' && env.LOCAL_MATRIX_CONTROL_TOKEN) {
    cacheKeyUrl.searchParams.set('__onlyoffice_matrix_run', env.LOCAL_MATRIX_CONTROL_TOKEN);
  }
  const cacheKey = new Request(cacheKeyUrl.href, { method: 'GET' });
  const rangeHeader = request.headers.get('range');
  const cacheable = shouldPopulateEdgeCache({
    method: request.method,
    immutable,
    hasRange: Boolean(rangeHeader),
    ...(immutableAsset?.publicSize !== undefined ? { publicSize: immutableAsset.publicSize } : {}),
    unboundedContentSegment: Boolean(contentSegmentRequest),
  });
  if (cacheable) {
    const cached = await cache.match(cacheKey);
    if (cached) {
      for (const counter of matrixCounters) counter.cacheHits += 1;
      completeMatrixRequest(matrixCounters, cached.status);
      return cached;
    }
  }

  if (request.method === 'HEAD') {
    const metadata = await env.ASSETS.head(key);
    for (const counter of matrixCounters) counter.r2Heads += 1;
    if (!metadata) {
      completeMatrixRequest(matrixCounters, 404);
      return assetError('Not found', 404);
    }
    if (immutableAsset?.requireExactObjectSize && metadata.size !== immutableAsset.publicSize) {
      completeMatrixRequest(matrixCounters, 502);
      return assetError('Content object size mismatch', 502);
    }
    const headers = responseHeaders(
      metadata,
      immutable,
      isolated,
      publicPath,
      assetVersion,
      immutableAsset?.mime,
      localMatrixSecurityUrl(url, env),
    );
    applyContentObjectHeaders(headers, immutableAsset ?? undefined);
    if (matchesEtag(request, metadata.httpEtag)) {
      completeMatrixRequest(matrixCounters, 304);
      return new Response(null, { status: 304, headers });
    }
    headers.set('Content-Length', String(immutableAsset?.publicSize ?? metadata.size));
    completeMatrixRequest(matrixCounters, 200);
    return new Response(null, { status: 200, headers });
  }

  const metadata = await env.ASSETS.head(key);
  for (const counter of matrixCounters) counter.r2Heads += 1;
  if (!metadata) {
    completeMatrixRequest(matrixCounters, 404);
    return assetError('Not found', 404);
  }
  if (immutableAsset?.requireExactObjectSize && metadata.size !== immutableAsset.publicSize) {
    completeMatrixRequest(matrixCounters, 502);
    return assetError('Content object size mismatch', 502);
  }
  const publicSize = immutableAsset?.publicSize ?? metadata?.size ?? 0;
  const objectOffset = immutableAsset?.objectOffset ?? 0;
  const range = rangeHeader ? parseSingleRange(rangeHeader, publicSize) : null;
  if (rangeHeader && !range) {
    completeMatrixRequest(matrixCounters, 416);
    return assetError(null, 416, {
      'Accept-Ranges': 'bytes',
      'Content-Range': `bytes */${publicSize}`,
      ...(immutableAsset?.contentSha256 ? { 'X-Content-SHA256': immutableAsset.contentSha256 } : {}),
    });
  }
  if (matchesEtag(request, metadata.httpEtag)) {
    const headers = responseHeaders(
      metadata,
      immutable,
      isolated,
      publicPath,
      assetVersion,
      immutableAsset?.mime,
      localMatrixSecurityUrl(url, env),
    );
    applyContentObjectHeaders(headers, immutableAsset ?? undefined);
    completeMatrixRequest(matrixCounters, 304);
    return new Response(null, { status: 304, headers });
  }

  const objectRange = range
    ? { start: objectOffset + range.start, end: objectOffset + range.end }
    : immutableAsset?.publicSize
      ? { start: objectOffset, end: objectOffset + immutableAsset.publicSize - 1 }
      : null;
  const localDelayMs =
    env.LOCAL_MATRIX_MODE === '1' && contentObjectRequest
      ? Number.parseInt(env.LOCAL_MATRIX_R2_DELAY_MS || '0', 10)
      : 0;
  if (Number.isSafeInteger(localDelayMs) && localDelayMs > 0 && localDelayMs <= 5_000) {
    await new Promise((resolve) => setTimeout(resolve, localDelayMs));
  }
  const object = await env.ASSETS.get(
    key,
    objectRange ? { range: new Headers({ Range: `bytes=${objectRange.start}-${objectRange.end}` }) } : undefined,
  );
  for (const counter of matrixCounters) counter.r2Gets += 1;
  if (!object?.body) {
    completeMatrixRequest(matrixCounters, 404);
    return assetError('Not found', 404);
  }
  const declaredBytes = objectRange ? objectRange.end - objectRange.start + 1 : publicSize;
  for (const counter of matrixCounters) counter.declaredBytes += declaredBytes;
  const headers = responseHeaders(
    object,
    immutable,
    isolated,
    publicPath,
    assetVersion,
    immutableAsset?.mime,
    localMatrixSecurityUrl(url, env),
  );
  applyContentObjectHeaders(headers, immutableAsset ?? undefined);

  let status = 200;
  if (range) {
    status = 206;
    headers.set('Accept-Ranges', 'bytes');
    headers.set('Content-Length', String(range.end - range.start + 1));
    headers.set('Content-Range', `bytes ${range.start}-${range.end}/${publicSize}`);
  } else {
    headers.set('Content-Length', String(publicSize));
  }
  const response = new Response(
    countR2Body(
      maybeStallLocalMatrixBody(object.body, env, key, matrixCounters),
      matrixCounters,
      status,
      declaredBytes,
    ),
    { status, headers },
  );
  if (cacheable) ctx.waitUntil(cache.put(cacheKey, response.clone()).catch(() => undefined));
  return response;
}

function applyContentObjectHeaders(headers: Headers, asset?: ImmutableAsset): void {
  if (!asset?.contentSha256) return;
  headers.set('Accept-Ranges', 'bytes');
  headers.set('X-Content-SHA256', asset.contentSha256);
}

function officeParentOrigins(localMatrixUrl?: URL): string[] {
  const localPort = localMatrixUrl?.port ? `:${localMatrixUrl.port}` : '';
  return [
    'https://piwork.getpi.work',
    'https://onlyoffice.getpi.work',
    ...(localMatrixUrl
      ? [
          `${localMatrixUrl.protocol}//piwork.localhost${localPort}`,
          `${localMatrixUrl.protocol}//onlyoffice.localhost${localPort}`,
        ]
      : []),
  ];
}

function responseHeaders(
  object: R2ObjectLike,
  immutable: boolean,
  isolated: boolean,
  key: string,
  version: string,
  releaseMime?: string,
  localMatrixUrl?: URL,
): Headers {
  const headers = new Headers();
  object.writeHttpMetadata(headers);
  applyReleaseMime(headers, releaseMime);
  headers.set('ETag', object.httpEtag);
  headers.set(
    'Cache-Control',
    key.startsWith('channels/')
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
  if (isolated && key === 'office-host.html') {
    headers.set('Origin-Agent-Cluster', '?1');
    headers.set(
      'Content-Security-Policy',
      [`frame-ancestors ${officeParentOrigins(localMatrixUrl).join(' ')}`, "base-uri 'none'", "object-src 'none'"].join(
        '; ',
      ),
    );
  }
  if (
    (!isolated && (key === 'resource-broker.html' || key === 'resource-installer.html')) ||
    (isolated && key === 'editor-shell-prime.html')
  ) {
    const localPort = localMatrixUrl?.port ? `:${localMatrixUrl.port}` : '';
    const parentOrigins = officeParentOrigins(localMatrixUrl);
    const productionEditorOrigins = OFFICE_EDITOR_SLOTS.map((slot) => `https://${slot}.getpi.work`);
    const localEditorOrigins = localMatrixUrl
      ? OFFICE_EDITOR_SLOTS.flatMap((slot) => [
          `${localMatrixUrl.protocol}//${slot}.localhost${localPort}`,
          `${localMatrixUrl.protocol}//host-${slot}.office.localhost${localPort}`,
        ])
      : [];
    const localFrameAncestors = localMatrixUrl
      ? [...parentOrigins.slice(2), ...(key === 'resource-broker.html' ? localEditorOrigins : [])]
      : [];
    const frameAncestors = [
      ...parentOrigins.slice(0, 2),
      ...(key === 'resource-broker.html' ? productionEditorOrigins : []),
      ...localFrameAncestors,
    ];
    const frameSources =
      key === 'resource-installer.html'
        ? [...productionEditorOrigins, ...localEditorOrigins]
        : key === 'editor-shell-prime.html'
          ? [
              'https://onlyoffice.getpi.work',
              ...(localMatrixUrl ? [`${localMatrixUrl.protocol}//onlyoffice.localhost${localPort}`] : []),
            ]
          : [];
    headers.set(
      'Content-Security-Policy',
      [
        "default-src 'none'",
        "script-src 'self'",
        "connect-src 'self'",
        "worker-src 'self'",
        "base-uri 'none'",
        "form-action 'none'",
        ...(frameSources.length ? [`frame-src ${frameSources.join(' ')}`] : []),
        `frame-ancestors ${frameAncestors.join(' ')}`,
      ].join('; '),
    );
  }
  return headers;
}

function localMatrixSecurityUrl(url: URL, env: WorkerEnv): URL | undefined {
  if (env.LOCAL_MATRIX_MODE !== '1') return undefined;
  const resolved = new URL(url);
  const configuredPort = env.LOCAL_MATRIX_PORT || '';
  if (
    !resolved.port &&
    /^(?:[1-9]\d{0,3}|[1-5]\d{4}|6[0-4]\d{3}|65[0-4]\d{2}|655[0-2]\d|6553[0-5])$/.test(configuredPort)
  ) {
    resolved.port = configuredPort;
  }
  return resolved;
}

type LocalMatrixStablePointer = {
  version: 1;
  releaseId: string;
  manifestUrl: string;
  manifestSha256: string;
};

async function sha256Hex(bytes: ArrayBuffer): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  return Array.from(digest, (value) => value.toString(16).padStart(2, '0')).join('');
}

async function updateLocalMatrixStablePointer(request: Request, env: WorkerEnv): Promise<Response> {
  if (
    !env.LOCAL_MATRIX_CONTROL_TOKEN ||
    request.headers.get(LOCAL_MATRIX_CONTROL_TOKEN_HEADER) !== env.LOCAL_MATRIX_CONTROL_TOKEN
  ) {
    return assetError('Unauthorized', 401);
  }
  if (!env.ASSETS.put) return assetError('Local R2 writes are unavailable', 501);
  const declaredLength = Number(request.headers.get('content-length') || '0');
  if (!Number.isSafeInteger(declaredLength) || declaredLength < 0 || declaredLength > 8 * 1024) {
    return assetError('Invalid pointer body size', 413);
  }
  const bytes = await request.arrayBuffer();
  if (bytes.byteLength > 8 * 1024) return assetError('Invalid pointer body size', 413);
  let pointer: LocalMatrixStablePointer;
  try {
    pointer = JSON.parse(new TextDecoder().decode(bytes)) as LocalMatrixStablePointer;
  } catch {
    return assetError('Invalid pointer JSON', 400);
  }
  if (
    pointer?.version !== 1 ||
    !RELEASE_ID_PATTERN.test(pointer.releaseId || '') ||
    pointer.manifestUrl !== `/releases/${pointer.releaseId}/manifest.json` ||
    !SHA256_PATTERN.test(pointer.manifestSha256 || '')
  ) {
    return assetError('Invalid stable pointer', 400);
  }
  const manifestObject = await env.ASSETS.get(`releases/${pointer.releaseId}/manifest.json`);
  if (!manifestObject?.body) return assetError('Target release manifest not found', 404);
  const manifestBytes = await new Response(manifestObject.body).arrayBuffer();
  if ((await sha256Hex(manifestBytes)) !== pointer.manifestSha256) {
    return assetError('Target release manifest digest mismatch', 409);
  }
  try {
    const manifest = JSON.parse(new TextDecoder().decode(manifestBytes)) as ReleaseManifest;
    if (manifest.version !== 5 || manifest.releaseId !== pointer.releaseId) {
      return assetError('Target release manifest identity mismatch', 409);
    }
  } catch {
    return assetError('Target release manifest is invalid', 409);
  }
  const pointerBytes = `${JSON.stringify(pointer, null, 2)}\n`;
  await env.ASSETS.put('channels/stable-v5.json', pointerBytes, {
    httpMetadata: {
      contentType: 'application/json; charset=utf-8',
      cacheControl: 'no-store',
    },
  });
  const response = Response.json({ ok: true, releaseId: pointer.releaseId });
  response.headers.set('Cache-Control', 'no-store');
  applySharedHeaders(response.headers);
  return response;
}

export function applyReleaseMime(headers: Headers, releaseMime?: string): void {
  if (releaseMime) headers.set('Content-Type', releaseMime);
}

export default {
  async fetch(request: Request, env: WorkerEnv, ctx: WorkerExecutionContext): Promise<Response> {
    const url = new URL(request.url);
    const requestHostname = resolveRequestHostname(request, env.LOCAL_MATRIX_MODE === '1');
    if (!requestHostname) return new Response('Missing local matrix host', { status: 421 });
    if (env.LOCAL_MATRIX_MODE === '1' && request.headers.has(LOCAL_MATRIX_HOST_HEADER)) {
      url.hostname = requestHostname;
      if (env.LOCAL_MATRIX_PORT) url.port = env.LOCAL_MATRIX_PORT;
    }
    const { logicalHostname: hostname, canonicalHostname } = resolveRuntimeHost(
      requestHostname,
      env.LOCAL_MATRIX_MODE === '1',
    );
    if (!isOnlyOfficeHost(hostname)) return new Response('Unknown host', { status: 404 });

    if (
      env.LOCAL_MATRIX_MODE === '1' &&
      hostname === CANONICAL_HOST &&
      request.method === 'POST' &&
      url.pathname === '/__matrix__/stable-v5'
    ) {
      return updateLocalMatrixStablePointer(request, env);
    }

    if (
      env.LOCAL_MATRIX_MODE === '1' &&
      hostname === CANONICAL_HOST &&
      request.method === 'GET' &&
      (url.pathname === '/__matrix__/segment-counters' ||
        url.pathname === '/__matrix__/object-counters' ||
        url.pathname === '/__matrix__/route-counters' ||
        url.pathname === '/__matrix__/content-counters')
    ) {
      const body =
        url.pathname === '/__matrix__/segment-counters'
          ? Object.fromEntries(localMatrixSegmentCounters)
          : url.pathname === '/__matrix__/object-counters'
            ? Object.fromEntries(localMatrixObjectCounters)
            : url.pathname === '/__matrix__/route-counters'
              ? Object.fromEntries(localMatrixRouteCounters)
              : {
                  segments: Object.fromEntries(localMatrixSegmentCounters),
                  objects: Object.fromEntries(localMatrixObjectCounters),
                  routes: Object.fromEntries(localMatrixRouteCounters),
                };
      const response = Response.json(body);
      applySharedHeaders(response.headers);
      return response;
    }

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
    if (isolated && isCanonicalReleaseMetadataPath(url.pathname)) {
      const canonical = new URL(url);
      canonical.hostname = canonicalHostname;
      return sharedRedirect(canonical.href);
    }
    if (isolated && resolveReleaseContentObjectRequest(url.pathname)) {
      const canonical = new URL(url);
      canonical.hostname = canonicalHostname;
      return sharedRedirect(canonical.href);
    }
    const pinnedReleaseId = releaseIdFromReferrer(request.headers.get('referer'));
    const editorAssetRoute = resolveEditorAssetRoute(url.pathname);
    if (isolated && shouldShareAsset(editorAssetRoute.pathname, request.headers.get('sec-fetch-dest'))) {
      const releaseId = editorAssetRoute.releaseId || pinnedReleaseId || (await stableReleaseId(env));
      if (releaseId) {
        const canonical = new URL(url);
        canonical.hostname = canonicalHostname;
        const releasePathname = canonicalReleasePathname(url.pathname, releaseId);
        if (!releasePathname) return new Response('Invalid asset path', { status: 400 });
        canonical.pathname = releasePathname;
        canonical.searchParams.delete(VERSION_QUERY);
        return sharedRedirect(canonical.href);
      }
      return sharedRedirect(canonicalAssetUrl(url, env.ASSET_VERSION, canonicalHostname));
    }
    if (isolated && pinnedReleaseId && !resolveReleaseRequest(url.pathname)) {
      url.pathname = `/r/${pinnedReleaseId}/${url.pathname.replace(/^\/+/, '')}`;
    }
    let response = await serveAsset(request, env, ctx, url, isolated);
    if (env.LOCAL_MATRIX_MODE === '1') {
      response = new Response(response.body, {
        status: response.status,
        statusText: response.statusText,
        headers: response.headers,
      });
      response.headers.set('X-OnlyOffice-Matrix-Host', requestHostname);
      response.headers.set('X-OnlyOffice-Matrix-Isolated', isolated ? '1' : '0');
      response.headers.set(
        'Access-Control-Expose-Headers',
        `${response.headers.get('Access-Control-Expose-Headers') || ''}, Content-Security-Policy, X-OnlyOffice-Matrix-Host, X-OnlyOffice-Matrix-Isolated`.replace(
          /^,\s*/,
          '',
        ),
      );
    }
    return response;
  },
};
