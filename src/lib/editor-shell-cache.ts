import { sha256 } from '@noble/hashes/sha2.js';
import { isResourceBrokerReleaseId } from './resource-broker-protocol';

export const EDITOR_SHELL_CACHE_NAME = 'onlyoffice-editor-shell-v1';
export const EDITOR_SHELL_HOST_PATH = 'office-host.html';
export const EDITOR_SHELL_PRIME_PATH = 'editor-shell-prime.html';
export const EDITOR_SHELL_PRIME_INSTALL_QUERY = 'install';
export const EDITOR_SHELL_MAX_HTML_BYTES = 1024 * 1024;
export const EDITOR_SHELL_MAX_MANIFEST_BYTES = 4 * 1024 * 1024;
export const EDITOR_SHELL_MAX_ASSET_BYTES = 4 * 1024 * 1024;
export const EDITOR_SHELL_MAX_TOTAL_BYTES = 16 * 1024 * 1024;
export const EDITOR_SHELL_MAX_ASSETS = 64;
export const EDITOR_SHELL_CACHE_RESPONSE_HEADER = 'x-onlyoffice-editor-shell-cache';

const EDITOR_SHELL_POINTER_VERSION = 1;
const EDITOR_SHELL_POINTER_PATH = '.onlyoffice-internal/editor-shell';
const EDITOR_SHELL_GENERATION_PREFIX = `${EDITOR_SHELL_CACHE_NAME}:generation:`;
const EDITOR_SHELL_MAX_POINTER_BYTES = 64 * 1024;
const EDITOR_SHELL_MAX_MANIFEST_ASSETS = 50_000;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const GENERATION_ID_PATTERN = /^[a-f0-9]{32}$/;
const SAFE_DEPENDENCY_PATTERN = /^assets\/[a-zA-Z0-9._+-]+\.(?:css|js)$/;
const MIME_BY_PATH = new Map([
  [EDITOR_SHELL_HOST_PATH, 'text/html; charset=utf-8'],
  [EDITOR_SHELL_PRIME_PATH, 'text/html; charset=utf-8'],
]);

type EditorShellAssetDescriptor = {
  path: string;
  bytes: number;
  mime: string;
  sha256: string;
};

type EditorShellPointer = {
  version: typeof EDITOR_SHELL_POINTER_VERSION;
  releaseId: string;
  manifestSha256: string;
  cacheName: string;
  cachedBytes: number;
  assets: EditorShellAssetDescriptor[];
};

export type EditorShellPrimeResult = {
  releaseId: string;
  cachedPaths: string[];
  cachedBytes: number;
};

function expectedMime(path: string): string | null {
  const documentMime = MIME_BY_PATH.get(path);
  if (documentMime) return documentMime;
  if (path.endsWith('.js')) return 'text/javascript; charset=utf-8';
  if (path.endsWith('.css')) return 'text/css; charset=utf-8';
  return null;
}

function normalizedMime(value: string | null): string {
  return (value || '')
    .split(';')
    .map((part) => part.trim().toLowerCase())
    .join('; ');
}

function isSafeManifestPath(path: string): boolean {
  if (
    path.length === 0 ||
    path.length > 2_048 ||
    path.startsWith('/') ||
    path.includes('\\') ||
    path.includes('\0') ||
    path.includes('?') ||
    path.includes('#') ||
    Array.from(path).some((character) => {
      const codePoint = character.codePointAt(0) ?? 0;
      return codePoint <= 0x1f || codePoint === 0x7f;
    }) ||
    /%(?:2e|2f|5c)/i.test(path)
  ) {
    return false;
  }
  try {
    const decoded = decodeURIComponent(path);
    return decoded.split('/').every((part) => part.length > 0 && part !== '.' && part !== '..');
  } catch {
    return false;
  }
}

export function isEditorShellAssetPath(path: string): boolean {
  return path === EDITOR_SHELL_HOST_PATH || path === EDITOR_SHELL_PRIME_PATH || SAFE_DEPENDENCY_PATTERN.test(path);
}

export function editorReleaseAssetUrl(origin: string | URL, releaseId: string, path: string): URL {
  if (!isResourceBrokerReleaseId(releaseId) || !isEditorShellAssetPath(path)) {
    throw new TypeError('Invalid editor shell asset identity');
  }
  return new URL(`/r/${encodeURIComponent(releaseId)}/${path}`, origin);
}

function editorReleaseManifestUrl(origin: string | URL, releaseId: string): URL {
  if (!isResourceBrokerReleaseId(releaseId)) throw new TypeError('Invalid editor shell release identity');
  return new URL(`/releases/${encodeURIComponent(releaseId)}/manifest.json`, origin);
}

function editorShellPointerUrl(origin: string | URL, releaseId: string): URL {
  if (!isResourceBrokerReleaseId(releaseId)) throw new TypeError('Invalid editor shell release identity');
  return new URL(`/${EDITOR_SHELL_POINTER_PATH}/${encodeURIComponent(releaseId)}`, origin);
}

function generationCacheName(releaseId: string, generationId: string): string {
  if (!isResourceBrokerReleaseId(releaseId) || !GENERATION_ID_PATTERN.test(generationId)) {
    throw new TypeError('Invalid editor shell generation identity');
  }
  return `${EDITOR_SHELL_GENERATION_PREFIX}${releaseId}:${generationId}`;
}

function isOwnedGenerationCacheName(cacheName: string, releaseId: string): boolean {
  const prefix = `${EDITOR_SHELL_GENERATION_PREFIX}${releaseId}:`;
  return cacheName.startsWith(prefix) && GENERATION_ID_PATTERN.test(cacheName.slice(prefix.length));
}

function randomGenerationId(): string {
  const value = new Uint8Array(16);
  crypto.getRandomValues(value);
  return Array.from(value, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export function releaseIdFromEditorShellPath(pathname: string, expectedPath: string): string | null {
  if (expectedPath !== EDITOR_SHELL_HOST_PATH && expectedPath !== EDITOR_SHELL_PRIME_PATH) return null;
  const match = new RegExp(`^/r/([^/]+)/${expectedPath.replace('.', '\\.')}$`).exec(pathname);
  if (!match) return null;
  try {
    const releaseId = decodeURIComponent(match[1]);
    return isResourceBrokerReleaseId(releaseId) ? releaseId : null;
  } catch {
    return null;
  }
}

export function extractEditorShellDependencies(html: string): string[] {
  const dependencies = new Set<string>();
  const attribute = /(?:src|href)=["']\.\/(assets\/[^"'?#]+)["']/g;
  for (const match of html.matchAll(attribute)) {
    if (!SAFE_DEPENDENCY_PATTERN.test(match[1])) {
      throw new TypeError('Office host contains an unsafe shell dependency');
    }
    dependencies.add(match[1]);
    if (dependencies.size > EDITOR_SHELL_MAX_ASSETS - 2) {
      throw new TypeError('Office host contains too many shell dependencies');
    }
  }
  if (dependencies.size === 0) throw new TypeError('Office host has no shell dependencies');
  return [...dependencies].sort();
}

function responseLength(response: Response): number | null {
  const value = response.headers.get('content-length');
  if (value === null || !/^(?:0|[1-9][0-9]*)$/.test(value)) return null;
  const bytes = Number(value);
  return Number.isSafeInteger(bytes) ? bytes : null;
}

async function fetchShellAsset(fetchImpl: typeof fetch, url: URL): Promise<Response> {
  const response = await fetchImpl(url, {
    cache: 'no-store',
    credentials: 'omit',
    redirect: 'error',
  });
  if (!response.ok || response.status !== 200 || response.type === 'opaque') {
    await response.body?.cancel().catch(() => undefined);
    throw new Error(`Editor shell asset request failed (${response.status})`);
  }
  return response;
}

async function readBoundedBody(response: Response, maximumBytes: number, label: string): Promise<Uint8Array> {
  const declaredBytes = responseLength(response);
  if (declaredBytes !== null && declaredBytes > maximumBytes) {
    await response.body?.cancel().catch(() => undefined);
    throw new Error(`${label} exceeds its bounded size`);
  }
  if (!response.body) return new Uint8Array();
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.byteLength;
      if (total > maximumBytes) {
        await reader.cancel().catch(() => undefined);
        throw new Error(`${label} exceeds its bounded size`);
      }
      chunks.push(value);
    }
  } catch (error) {
    await reader.cancel().catch(() => undefined);
    throw error;
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

function digestBytes(bytes: Uint8Array): string {
  return Array.from(sha256(bytes), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function decodeUtf8(bytes: Uint8Array, label: string): string {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    throw new Error(`${label} is not valid UTF-8`);
  }
}

function assertResponseIdentity(
  response: Response,
  descriptor: EditorShellAssetDescriptor,
  releaseId: string,
  label: string,
): void {
  const declaredBytes = responseLength(response);
  // Browsers remove Content-Length when they expose a transparently decoded
  // gzip/br response, and HTTP/2/3 responses may not carry it at all. Treat a
  // present value as a strict early check; readAndValidateAsset still enforces
  // the exact decoded byte length, the bounded read, and the manifest SHA-256.
  if (declaredBytes !== null && declaredBytes !== descriptor.bytes) {
    throw new Error(`${label} has an invalid byte length`);
  }
  if (normalizedMime(response.headers.get('content-type')) !== normalizedMime(descriptor.mime)) {
    throw new Error(`${label} has an invalid content type`);
  }
  const responseRelease = response.headers.get('x-onlyoffice-asset-version');
  if (responseRelease !== null && responseRelease !== releaseId) {
    throw new Error(`${label} has an invalid release identity`);
  }
  const responseDigest = response.headers.get('x-content-sha256');
  if (responseDigest !== null && responseDigest !== descriptor.sha256) {
    throw new Error(`${label} has an invalid content digest`);
  }
}

async function readAndValidateAsset(
  response: Response,
  descriptor: EditorShellAssetDescriptor,
  releaseId: string,
  label: string,
): Promise<{ bytes: Uint8Array; headers: Headers }> {
  assertResponseIdentity(response, descriptor, releaseId, label);
  const bytes = await readBoundedBody(response, descriptor.bytes, label);
  if (bytes.byteLength !== descriptor.bytes) throw new Error(`${label} has an invalid byte length`);
  if (digestBytes(bytes) !== descriptor.sha256) throw new Error(`${label} failed SHA-256 verification`);
  return { bytes, headers: new Headers(response.headers) };
}

function responseBody(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

function parseManifestDescriptor(value: unknown, path: string): EditorShellAssetDescriptor {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError(`Release manifest is missing editor shell asset ${path}`);
  }
  const asset = value as Record<string, unknown>;
  const mime = expectedMime(path);
  if (
    asset.path !== path ||
    !Number.isSafeInteger(asset.bytes) ||
    Number(asset.bytes) <= 0 ||
    Number(asset.bytes) > EDITOR_SHELL_MAX_ASSET_BYTES ||
    typeof asset.mime !== 'string' ||
    !mime ||
    normalizedMime(asset.mime) !== normalizedMime(mime) ||
    typeof asset.sha256 !== 'string' ||
    !SHA256_PATTERN.test(asset.sha256)
  ) {
    throw new TypeError(`Release manifest has an invalid editor shell asset ${path}`);
  }
  return {
    path,
    bytes: Number(asset.bytes),
    mime,
    sha256: asset.sha256,
  };
}

async function readReleaseManifest(
  response: Response,
  releaseId: string,
): Promise<{
  manifestSha256: string;
  assets: Map<string, unknown>;
}> {
  if (normalizedMime(response.headers.get('content-type')) !== 'application/json; charset=utf-8') {
    await response.body?.cancel().catch(() => undefined);
    throw new TypeError('Release manifest has an invalid content type');
  }
  const bytes = await readBoundedBody(response, EDITOR_SHELL_MAX_MANIFEST_BYTES, 'Release manifest');
  let value: unknown;
  try {
    value = JSON.parse(decodeUtf8(bytes, 'Release manifest'));
  } catch (error) {
    if (error instanceof SyntaxError) throw new TypeError('Release manifest is not valid JSON');
    throw error;
  }
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('Release manifest has an invalid shape');
  }
  const manifest = value as Record<string, unknown>;
  if (manifest.version !== 5 || manifest.releaseId !== releaseId || !Array.isArray(manifest.assets)) {
    throw new TypeError('Release manifest has an incompatible identity');
  }
  if (manifest.assets.length === 0 || manifest.assets.length > EDITOR_SHELL_MAX_MANIFEST_ASSETS) {
    throw new TypeError('Release manifest has an invalid asset count');
  }
  const assets = new Map<string, unknown>();
  for (const candidate of manifest.assets) {
    if (typeof candidate !== 'object' || candidate === null || Array.isArray(candidate)) {
      throw new TypeError('Release manifest contains an invalid asset');
    }
    const path = (candidate as Record<string, unknown>).path;
    if (typeof path !== 'string' || !isSafeManifestPath(path) || assets.has(path)) {
      throw new TypeError('Release manifest contains an invalid asset path');
    }
    assets.set(path, candidate);
  }
  return { manifestSha256: digestBytes(bytes), assets };
}

async function fetchReleaseManifest(
  fetchImpl: typeof fetch,
  origin: string | URL,
  releaseId: string,
): Promise<{
  manifestSha256: string;
  assets: Map<string, unknown>;
}> {
  return readReleaseManifest(await fetchShellAsset(fetchImpl, editorReleaseManifestUrl(origin, releaseId)), releaseId);
}

function validateDescriptorSet(releaseId: string, descriptors: EditorShellAssetDescriptor[]): number {
  if (!isResourceBrokerReleaseId(releaseId) || descriptors.length < 3 || descriptors.length > EDITOR_SHELL_MAX_ASSETS) {
    throw new TypeError('Invalid editor shell descriptor set');
  }
  const paths = new Set<string>();
  let totalBytes = 0;
  for (const descriptor of descriptors) {
    if (
      paths.has(descriptor.path) ||
      !isEditorShellAssetPath(descriptor.path) ||
      descriptor.mime !== expectedMime(descriptor.path) ||
      !Number.isSafeInteger(descriptor.bytes) ||
      descriptor.bytes <= 0 ||
      descriptor.bytes > EDITOR_SHELL_MAX_ASSET_BYTES ||
      !SHA256_PATTERN.test(descriptor.sha256)
    ) {
      throw new TypeError('Invalid editor shell descriptor');
    }
    paths.add(descriptor.path);
    totalBytes += descriptor.bytes;
    if (!Number.isSafeInteger(totalBytes) || totalBytes > EDITOR_SHELL_MAX_TOTAL_BYTES) {
      throw new TypeError('Editor shell exceeds its bounded total size');
    }
  }
  if (!paths.has(EDITOR_SHELL_HOST_PATH) || !paths.has(EDITOR_SHELL_PRIME_PATH)) {
    throw new TypeError('Editor shell descriptor set is incomplete');
  }
  return totalBytes;
}

function serializePointer(pointer: EditorShellPointer): Uint8Array {
  return new TextEncoder().encode(JSON.stringify(pointer));
}

function parsePointer(value: unknown, releaseId: string): EditorShellPointer {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) {
    throw new TypeError('Editor shell pointer has an invalid shape');
  }
  const pointer = value as Record<string, unknown>;
  if (
    Object.keys(pointer).length !== 6 ||
    pointer.version !== EDITOR_SHELL_POINTER_VERSION ||
    pointer.releaseId !== releaseId ||
    typeof pointer.manifestSha256 !== 'string' ||
    !SHA256_PATTERN.test(pointer.manifestSha256) ||
    typeof pointer.cacheName !== 'string' ||
    !isOwnedGenerationCacheName(pointer.cacheName, releaseId) ||
    !Number.isSafeInteger(pointer.cachedBytes) ||
    !Array.isArray(pointer.assets)
  ) {
    throw new TypeError('Editor shell pointer has an invalid identity');
  }
  const assets = pointer.assets.map((candidate) => {
    if (
      typeof candidate !== 'object' ||
      candidate === null ||
      Array.isArray(candidate) ||
      Object.keys(candidate).length !== 4
    ) {
      throw new TypeError('Editor shell pointer contains an invalid asset');
    }
    const descriptor = candidate as Record<string, unknown>;
    if (
      typeof descriptor.path !== 'string' ||
      typeof descriptor.mime !== 'string' ||
      typeof descriptor.sha256 !== 'string'
    ) {
      throw new TypeError('Editor shell pointer contains an invalid asset');
    }
    return {
      path: descriptor.path,
      bytes: Number(descriptor.bytes),
      mime: descriptor.mime,
      sha256: descriptor.sha256,
    };
  });
  const cachedBytes = validateDescriptorSet(releaseId, assets);
  if (cachedBytes !== pointer.cachedBytes) throw new TypeError('Editor shell pointer has an invalid byte total');
  const sorted = [...assets].sort((left, right) => left.path.localeCompare(right.path));
  if (sorted.some((asset, index) => asset.path !== assets[index]?.path)) {
    throw new TypeError('Editor shell pointer assets are not canonical');
  }
  return {
    version: EDITOR_SHELL_POINTER_VERSION,
    releaseId,
    manifestSha256: pointer.manifestSha256,
    cacheName: pointer.cacheName,
    cachedBytes,
    assets,
  };
}

async function readPointer(
  cacheStorage: CacheStorage,
  origin: string | URL,
  releaseId: string,
): Promise<EditorShellPointer> {
  const metadata = await cacheStorage.open(EDITOR_SHELL_CACHE_NAME);
  const response = await metadata.match(editorShellPointerUrl(origin, releaseId).href);
  if (!response || !response.ok) throw new Error('Editor shell cache has no active generation');
  if (normalizedMime(response.headers.get('content-type')) !== 'application/json; charset=utf-8') {
    await response.body?.cancel().catch(() => undefined);
    throw new Error('Editor shell pointer has an invalid content type');
  }
  const bytes = await readBoundedBody(response, EDITOR_SHELL_MAX_POINTER_BYTES, 'Editor shell pointer');
  let value: unknown;
  try {
    value = JSON.parse(decodeUtf8(bytes, 'Editor shell pointer'));
  } catch {
    throw new Error('Editor shell pointer is invalid');
  }
  return parsePointer(value, releaseId);
}

async function verifyGeneration(
  cacheStorage: CacheStorage,
  origin: string | URL,
  pointer: EditorShellPointer,
): Promise<EditorShellPrimeResult> {
  const cache = await cacheStorage.open(pointer.cacheName);
  const htmlDependencies: string[] = [];
  for (const descriptor of pointer.assets) {
    const url = editorReleaseAssetUrl(origin, pointer.releaseId, descriptor.path);
    const response = await cache.match(url.href);
    if (!response || !response.ok) throw new Error(`Editor shell cache is missing ${descriptor.path}`);
    const asset = await readAndValidateAsset(response, descriptor, pointer.releaseId, `Cached ${descriptor.path}`);
    if (descriptor.path === EDITOR_SHELL_HOST_PATH || descriptor.path === EDITOR_SHELL_PRIME_PATH) {
      htmlDependencies.push(...extractEditorShellDependencies(decodeUtf8(asset.bytes, `Cached ${descriptor.path}`)));
    }
  }
  const referencedPaths = [EDITOR_SHELL_HOST_PATH, EDITOR_SHELL_PRIME_PATH, ...new Set(htmlDependencies)].sort();
  const storedPaths = pointer.assets.map((asset) => asset.path);
  if (
    referencedPaths.length !== storedPaths.length ||
    referencedPaths.some((path, index) => path !== storedPaths[index])
  ) {
    throw new Error('Editor shell cache dependency graph does not match its manifest');
  }
  return {
    releaseId: pointer.releaseId,
    cachedPaths: storedPaths,
    cachedBytes: pointer.cachedBytes,
  };
}

export async function primeEditorShell(options: {
  releaseId: string;
  origin: string | URL;
  manifestOrigin?: string | URL;
  cacheStorage: CacheStorage;
  fetch: typeof fetch;
  createGenerationId?: () => string;
}): Promise<EditorShellPrimeResult> {
  try {
    // Release paths are immutable. A deeply verified same-release shell must
    // never be overwritten or deleted by a later failed re-prime.
    return await verifyEditorShell(options);
  } catch {
    // Missing, legacy or incomplete shells are repaired in an isolated staging
    // cache. They remain invisible until the generation pointer is committed.
  }

  const manifest = await fetchReleaseManifest(
    options.fetch,
    options.manifestOrigin ?? options.origin,
    options.releaseId,
  );
  const hostDescriptor = parseManifestDescriptor(manifest.assets.get(EDITOR_SHELL_HOST_PATH), EDITOR_SHELL_HOST_PATH);
  const primeDescriptor = parseManifestDescriptor(
    manifest.assets.get(EDITOR_SHELL_PRIME_PATH),
    EDITOR_SHELL_PRIME_PATH,
  );
  const hostUrl = editorReleaseAssetUrl(options.origin, options.releaseId, EDITOR_SHELL_HOST_PATH);
  const primeUrl = editorReleaseAssetUrl(options.origin, options.releaseId, EDITOR_SHELL_PRIME_PATH);
  const primeInstallUrl = new URL(primeUrl);
  primeInstallUrl.searchParams.set(EDITOR_SHELL_PRIME_INSTALL_QUERY, '1');
  const host = await readAndValidateAsset(
    await fetchShellAsset(options.fetch, hostUrl),
    hostDescriptor,
    options.releaseId,
    'Editor shell host HTML',
  );
  const prime = await readAndValidateAsset(
    await fetchShellAsset(options.fetch, primeInstallUrl),
    primeDescriptor,
    options.releaseId,
    'Editor shell prime HTML',
  );
  const dependencies = [
    ...new Set([
      ...extractEditorShellDependencies(decodeUtf8(host.bytes, 'Editor shell host HTML')),
      ...extractEditorShellDependencies(decodeUtf8(prime.bytes, 'Editor shell prime HTML')),
    ]),
  ].sort();
  const descriptors = [
    hostDescriptor,
    primeDescriptor,
    ...dependencies.map((path) => parseManifestDescriptor(manifest.assets.get(path), path)),
  ].sort((left, right) => left.path.localeCompare(right.path));
  const cachedBytes = validateDescriptorSet(options.releaseId, descriptors);
  const generationId = (options.createGenerationId ?? randomGenerationId)();
  const cacheName = generationCacheName(options.releaseId, generationId);
  const staging = await options.cacheStorage.open(cacheName);
  let committed = false;

  try {
    const documents = new Map([
      [EDITOR_SHELL_HOST_PATH, host],
      [EDITOR_SHELL_PRIME_PATH, prime],
    ]);
    for (const descriptor of descriptors) {
      const url = editorReleaseAssetUrl(options.origin, options.releaseId, descriptor.path);
      const document = documents.get(descriptor.path);
      const asset =
        document ||
        (await readAndValidateAsset(
          await fetchShellAsset(options.fetch, url),
          descriptor,
          options.releaseId,
          `Editor shell dependency ${descriptor.path}`,
        ));
      await staging.put(
        url,
        new Response(responseBody(asset.bytes), {
          status: 200,
          headers: asset.headers,
        }),
      );
    }

    const pointer: EditorShellPointer = {
      version: EDITOR_SHELL_POINTER_VERSION,
      releaseId: options.releaseId,
      manifestSha256: manifest.manifestSha256,
      cacheName,
      cachedBytes,
      assets: descriptors,
    };
    await verifyGeneration(options.cacheStorage, options.origin, pointer);

    let previous: EditorShellPointer | null = null;
    try {
      previous = await readPointer(options.cacheStorage, options.origin, options.releaseId);
    } catch {
      // Legacy, corrupt and absent pointers are never trusted for cleanup.
    }
    const metadata = await options.cacheStorage.open(EDITOR_SHELL_CACHE_NAME);
    const pointerBytes = serializePointer(pointer);
    if (pointerBytes.byteLength > EDITOR_SHELL_MAX_POINTER_BYTES) {
      throw new Error('Editor shell pointer exceeds its bounded size');
    }
    await metadata.put(
      editorShellPointerUrl(options.origin, options.releaseId),
      new Response(responseBody(pointerBytes), {
        status: 200,
        headers: {
          'content-length': String(pointerBytes.byteLength),
          'content-type': 'application/json; charset=utf-8',
        },
      }),
    );
    committed = true;
    if (previous && previous.cacheName !== cacheName) {
      await options.cacheStorage.delete(previous.cacheName).catch(() => undefined);
    }
    return {
      releaseId: options.releaseId,
      cachedPaths: descriptors.map((descriptor) => descriptor.path),
      cachedBytes,
    };
  } finally {
    if (!committed) await options.cacheStorage.delete(cacheName).catch(() => undefined);
  }
}

export async function verifyEditorShell(options: {
  releaseId: string;
  origin: string | URL;
  cacheStorage: CacheStorage;
}): Promise<EditorShellPrimeResult> {
  const pointer = await readPointer(options.cacheStorage, options.origin, options.releaseId);
  return verifyGeneration(options.cacheStorage, options.origin, pointer);
}

export async function matchEditorShell(
  request: Request,
  releaseId: string,
  cacheStorage: CacheStorage,
): Promise<Response | undefined> {
  if (request.method !== 'GET' && request.method !== 'HEAD') return undefined;
  const url = new URL(request.url);
  const prefix = `/r/${encodeURIComponent(releaseId)}/`;
  if (!url.pathname.startsWith(prefix) || url.search) return undefined;
  const path = url.pathname.slice(prefix.length);
  if (!isEditorShellAssetPath(path)) return undefined;
  let pointer: EditorShellPointer;
  try {
    pointer = await readPointer(cacheStorage, url.origin, releaseId);
  } catch {
    return undefined;
  }
  if (!pointer.assets.some((asset) => asset.path === path)) return undefined;
  const cache = await cacheStorage.open(pointer.cacheName);
  // The Host carries its parent/session identity in a fragment. FetchEvent
  // preserves that fragment even though it is never sent over HTTP, while
  // Cache Storage keys the immutable shell object without it. Resolve the key
  // from the already-validated release and path instead of treating the
  // client-only identity fragment as part of the stored object URL.
  const response = await cache.match(editorReleaseAssetUrl(url.origin, releaseId, path).href);
  if (!response) return undefined;
  const headers = new Headers(response.headers);
  headers.delete('content-encoding');
  headers.set(EDITOR_SHELL_CACHE_RESPONSE_HEADER, '1');
  if (request.method === 'HEAD') {
    return new Response(null, { status: response.status, statusText: response.statusText, headers });
  }
  const bytes = await response.arrayBuffer();
  headers.set('content-length', String(bytes.byteLength));
  return new Response(bytes, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}
