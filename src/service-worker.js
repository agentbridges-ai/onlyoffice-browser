import { clientsClaim, setCacheNameDetails } from 'workbox-core';
import { ExpirationPlugin } from 'workbox-expiration';
import { cleanupOutdatedCaches, matchPrecache, precacheAndRoute } from 'workbox-precaching';
import { registerRoute } from 'workbox-routing';
import { NetworkFirst, StaleWhileRevalidate } from 'workbox-strategies';
import { isProductionOfficeEditorHostname } from './lib/office-origin-pool';
import { CanonicalResourceStore, IndexedDbCanonicalResourceJournal } from './lib/canonical-resource-store';
import {
  CanonicalResourceBrokerService,
  parseCanonicalResourceBrokerClientUrl,
} from './lib/canonical-resource-broker-service';
import {
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  normalizeOnlyOfficeRuntimeRequestPath,
  normalizeResourceBrokerResourcePath,
  parseResourceBrokerClientMessage,
} from './lib/resource-broker-protocol';
import {
  EDITOR_SHELL_CACHE_NAME,
  EDITOR_SHELL_HOST_PATH,
  EDITOR_SHELL_PRIME_INSTALL_QUERY,
  EDITOR_SHELL_PRIME_PATH,
  matchEditorShell,
  primeEditorShell,
  releaseIdFromEditorShellPath,
  verifyEditorShell,
} from './lib/editor-shell-cache';
import {
  EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE,
  EDITOR_RESOURCE_BROKER_BIND_TYPE,
  EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS,
  EditorResourceBrokerClient,
  EditorResourceBrokerError,
  parseEditorResourceBrokerBindMessage,
} from './lib/editor-resource-broker';
import { EditorClientIdentityRegistry } from './lib/editor-client-identity-registry';
import { releaseIdFromOfficeHostUrl } from './lib/office-host-url';

const SERVICE_WORKER_VERSION = 'SW_VERSION_PLACEHOLDER';
const PRINT_PDF_CACHE_NAME = 'onlyoffice-browser-print-pdfs';
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const ONLYOFFICE_RUNTIME_ASSET_REGEX = /(^|\/)(web-apps|sdkjs|wasm\/x2t)\//;
const ONLYOFFICE_NAVIGATION_PATHS = new Set(['/office-host.html', '/reset.html']);
const PWA_APP_NAVIGATION_PATHS = new Set(['/', '/index.html', '/resource-broker.html']);
const ONLYOFFICE_RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const CONTENT_SEGMENT_PATH_REGEX = /^\/segments\/sha256\/[a-f0-9]{64}$/;
const CONTENT_OBJECT_PATH_REGEX = /^\/objects\/[^/]+\/sha256\/[a-f0-9]{64}$/;
const CONTENT_BLOB_PATH_REGEX = /^\/blobs\/sha256\/[a-f0-9]{64}$/;
const CANONICAL_OFFICE_HOST = 'onlyoffice.getpi.work';
const LEGACY_EDITOR_HOST_PATTERN = /^office-editor-[a-z0-9-]+\.getpi\.work$/;
const LOCAL_EDITOR_HOST_PATTERN =
  /^host-(?:aries|taurus|gemini|cancer|leo|virgo|libra|scorpio|sagittarius|capricorn|aquarius|pisces)\.office\.localhost$/;
const LOCAL_LEGACY_EDITOR_HOST_PATTERN = /^host-office-editor-[a-z0-9-]+\.office\.localhost$/;
// Match the local PWA origin so the matrix exercises the same cross-editor
// HTTP cache key sharing used by onlyoffice.getpi.work in production.
const LOCAL_CANONICAL_OFFICE_HOST = 'onlyoffice.localhost';
const SHARED_ASSET_VERSION_QUERY = '__oobv';

const MAX_CACHE_ITEMS = 100;
const RUNTIME_CACHE_NAME = 'onlyoffice-browser-runtime-v1';
const STATIC_CACHE_NAME = 'onlyoffice-browser-static-v1';

const isCanonicalPwaHost =
  self.location.hostname === CANONICAL_OFFICE_HOST || self.location.hostname === LOCAL_CANONICAL_OFFICE_HOST;
const isLocalCanonicalPwaHost = self.location.hostname === LOCAL_CANONICAL_OFFICE_HOST;
const isLocalEditorHost =
  LOCAL_EDITOR_HOST_PATTERN.test(self.location.hostname) ||
  LOCAL_LEGACY_EDITOR_HOST_PATTERN.test(self.location.hostname);
const isIsolatedEditorHost =
  isProductionOfficeEditorHostname(self.location.hostname) ||
  LEGACY_EDITOR_HOST_PATTERN.test(self.location.hostname) ||
  isLocalEditorHost;
const canonicalResourceJournal = isCanonicalPwaHost ? new IndexedDbCanonicalResourceJournal() : null;
const canonicalResourceStore =
  canonicalResourceJournal &&
  new CanonicalResourceStore({
    cacheStorage: caches,
    journal: canonicalResourceJournal,
    fetch: (...args) => fetch(...args),
    objectUrl: (releaseId, object) =>
      new URL(`/objects/${encodeURIComponent(releaseId)}/sha256/${object.sha256}`, self.location.origin),
    cacheKeyOrigin: self.location.origin,
  });
const canonicalResourceBroker =
  canonicalResourceJournal &&
  canonicalResourceStore &&
  new CanonicalResourceBrokerService({
    journal: canonicalResourceJournal,
    store: canonicalResourceStore,
  });

const parseEditorOfficeHostClientIdentity = (value) => {
  let url;
  try {
    url = new URL(value);
  } catch {
    return null;
  }
  if (url.origin !== self.location.origin || url.search) return null;
  const releaseId = releaseIdFromOfficeHostUrl(url);
  const fragment = new URLSearchParams(url.hash.startsWith('#') ? url.hash.slice(1) : url.hash);
  const sessionIds = fragment.getAll('sessionId');
  const sessionId = sessionIds.length === 1 ? sessionIds[0] : '';
  if (!releaseId || !isResourceBrokerSessionId(sessionId)) return null;
  return { releaseId, sessionId };
};

const parseEditorPrimeClientIdentity = (value) => {
  let url;
  try {
    url = new URL(value);
  } catch {
    return null;
  }
  if (url.origin !== self.location.origin) return null;
  const releaseId = releaseIdFromEditorShellPath(url.pathname, EDITOR_SHELL_PRIME_PATH);
  const fragment = new URLSearchParams(url.hash.startsWith('#') ? url.hash.slice(1) : url.hash);
  const sessionIds = fragment.getAll('sessionId');
  const releaseIds = fragment.getAll('releaseId');
  const modes = fragment.getAll('mode');
  const sessionId = sessionIds.length === 1 ? sessionIds[0] : '';
  const mode = modes.length === 1 ? modes[0] : '';
  if (
    !releaseId ||
    releaseIds.length !== 1 ||
    releaseIds[0] !== releaseId ||
    !isResourceBrokerSessionId(sessionId) ||
    (mode !== 'install' && mode !== 'verify') ||
    (mode === 'install'
      ? url.searchParams.get(EDITOR_SHELL_PRIME_INSTALL_QUERY) !== '1' || [...url.searchParams.keys()].length !== 1
      : url.search !== '')
  ) {
    return null;
  }
  return { releaseId, sessionId, mode };
};

const requestEditorBrokerConnection = async (identity) => {
  if (!identity) throw new Error('Editor Resource Broker identity is unavailable');
  const windowClients = await self.clients.matchAll({
    includeUncontrolled: false,
    type: 'window',
  });
  const hostClient = windowClients.find((client) => {
    const candidate = parseEditorOfficeHostClientIdentity(client.url);
    return candidate?.releaseId === identity.releaseId && candidate.sessionId === identity.sessionId;
  });
  if (!hostClient) throw new Error('Versioned Office host client is unavailable');
  hostClient.postMessage({
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'ONLYOFFICE_BROKER_NEEDED',
    releaseId: identity.releaseId,
    sessionId: identity.sessionId,
  });
};

const editorResourceBroker = isIsolatedEditorHost
  ? new EditorResourceBrokerClient({
      requestConnection: requestEditorBrokerConnection,
    })
  : null;
const editorBrokerMatrixDiagnostics = [];
const recordEditorBrokerMatrixDiagnostic = (value) => {
  if (!isLocalEditorHost) return;
  editorBrokerMatrixDiagnostics.push({ at: Date.now(), ...value });
  if (editorBrokerMatrixDiagnostics.length > 40) editorBrokerMatrixDiagnostics.shift();
};
// DevTools can evaluate this read-only snapshot in the worker target during
// production acceptance. It is intentionally not reachable through the
// Service Worker message protocol, so pages and untrusted origins gain no
// diagnostics or control surface.
Object.defineProperty(self, '__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__', {
  configurable: false,
  enumerable: false,
  value: () => ({
    schemaVersion: 1,
    role: isCanonicalPwaHost ? 'canonical-service-worker' : isIsolatedEditorHost ? 'editor-service-worker' : 'shell',
    serviceWorkerVersion: SERVICE_WORKER_VERSION,
    canonical: canonicalResourceBroker?.metrics ?? null,
    editor: editorResourceBroker?.metrics ?? null,
    ...(isLocalEditorHost ? { matrixDiagnostics: [...editorBrokerMatrixDiagnostics] } : {}),
  }),
  writable: false,
});
const editorClientIdentities = isIsolatedEditorHost ? new EditorClientIdentityRegistry() : null;
const editorPrimeBindingCapabilities = new Map();
const editorShellPrimeInFlight = new Map();
let editorBrokerBindQueue = Promise.resolve();
setCacheNameDetails({
  prefix: 'onlyoffice-browser',
  precache: 'shell-v1',
  suffix: '',
});
const pwaShellManifest = self.__WB_MANIFEST || [];
const editorBootstrapAssetPaths = new Set(
  pwaShellManifest
    .map((entry) => (typeof entry === 'string' ? entry : entry.url))
    .map((value) => new URL(value, self.location.origin).pathname)
    .filter((pathname) => pathname.startsWith('/assets/')),
);
if (isCanonicalPwaHost) {
  precacheAndRoute(pwaShellManifest);
  cleanupOutdatedCaches();
}
if (isIsolatedEditorHost) {
  clientsClaim();
}

const isOnlyOfficeRuntimeAsset = (url) => ONLYOFFICE_RUNTIME_ASSET_REGEX.test(url.pathname);
const isContentSegmentRequest = (url) => CONTENT_SEGMENT_PATH_REGEX.test(url.pathname);
const isCanonicalManagedResourceRequest = (url) =>
  isContentSegmentRequest(url) ||
  CONTENT_OBJECT_PATH_REGEX.test(url.pathname) ||
  CONTENT_BLOB_PATH_REGEX.test(url.pathname) ||
  url.pathname.startsWith('/channels/') ||
  url.pathname.startsWith('/releases/') ||
  parseEditorReleaseAssetPath(url.pathname) !== null ||
  isOnlyOfficeRuntimeAsset(url) ||
  url.pathname === ONLYOFFICE_RUNTIME_MANIFEST_PATH;

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

const parseEditorReleaseAssetPath = (pathname) => {
  const match = /^\/r\/([^/]+)\/(.+)$/.exec(pathname);
  if (!match) return null;
  let releaseId;
  try {
    releaseId = decodeURIComponent(match[1]);
  } catch {
    return null;
  }
  const path = normalizeResourceBrokerResourcePath(match[2]);
  return isResourceBrokerReleaseId(releaseId) && path ? { releaseId, path } : null;
};

const isEditorShellAssetPath = (path) =>
  path === EDITOR_SHELL_HOST_PATH || path === EDITOR_SHELL_PRIME_PATH || /^assets\/[a-zA-Z0-9._+-]+$/.test(path);

const isEditorDirectOriginBoundPath = (url) => {
  if (
    url.pathname === '/sw.js' ||
    url.pathname === '/document_editor_service_worker.js' ||
    url.pathname === '/editor-shell-prime.html' ||
    url.pathname === '/reset.html' ||
    url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX)
  ) {
    return true;
  }
  return parseEditorReleaseAssetPath(url.pathname)?.path === 'editor-shell-prime.html';
};

const resolveEditorBrokerAssetPath = (url) => {
  const releaseAsset = parseEditorReleaseAssetPath(url.pathname);
  if (releaseAsset) {
    return isEditorShellAssetPath(releaseAsset.path) || releaseAsset.path === 'editor-shell-prime.html'
      ? null
      : releaseAsset.path;
  }
  return normalizeOnlyOfficeRuntimeRequestPath(url.pathname);
};

const isEditorOfficeRuntimeClientUrl = (value, releaseId) => {
  let url;
  try {
    url = new URL(value);
  } catch {
    return false;
  }
  if (url.origin !== self.location.origin || url.hash) return false;
  const releaseAsset = parseEditorReleaseAssetPath(url.pathname);
  if (releaseAsset) {
    return (
      (!releaseId || releaseAsset.releaseId === releaseId) &&
      !isEditorShellAssetPath(releaseAsset.path) &&
      releaseAsset.path !== 'editor-shell-prime.html'
    );
  }
  const path = normalizeResourceBrokerResourcePath(url.pathname);
  return Boolean(
    path &&
    (path.startsWith('web-apps/') ||
      path.startsWith('sdkjs/') ||
      path.startsWith('wasm/x2t/') ||
      path.startsWith('fonts/')),
  );
};

const resolveEditorFetchIdentity = async (event, releaseId, diagnostics = null) => {
  const clientId = event.clientId || event.resultingClientId;
  if (!clientId || !editorClientIdentities) {
    if (diagnostics)
      Object.assign(diagnostics, { clientId: Boolean(clientId), registry: Boolean(editorClientIdentities) });
    return null;
  }
  // A navigation's resultingClientId names the client that will exist only
  // after respondWith() completes. Waiting on clients.get(resultingClientId)
  // therefore deadlocks iframe navigations. Only inspect an incumbent client;
  // a new navigation is authenticated from its exact referrer and the bound
  // Broker identity, then recorded under resultingClientId below.
  const sourceClient = event.clientId ? await self.clients.get(event.clientId) : null;
  const sourceIdentity = parseEditorOfficeHostClientIdentity(sourceClient?.url || event.request.referrer);
  const connectedIdentity = editorResourceBroker?.connectionState.identity ?? null;
  const isNewEditorIframeNavigation = Boolean(
    !event.clientId &&
    event.resultingClientId &&
    event.request.mode === 'navigate' &&
    event.request.destination === 'iframe',
  );
  const recoveryIdentities = [];
  for (const client of await self.clients.matchAll({ includeUncontrolled: false, type: 'window' })) {
    const identity = parseEditorOfficeHostClientIdentity(client.url);
    if (identity) recoveryIdentities.push(identity);
  }
  const identity = editorClientIdentities.resolve({
    clientId,
    resultingClientId: event.resultingClientId || undefined,
    releaseId,
    exactSourceIdentity: sourceIdentity,
    connectedIdentity,
    sourceIsOfficeRuntime:
      isNewEditorIframeNavigation ||
      isEditorOfficeRuntimeClientUrl(sourceClient?.url || event.request.referrer, releaseId),
    recoveryIdentities,
  });
  if (diagnostics) {
    Object.assign(diagnostics, {
      clientId: true,
      sourceClient: Boolean(sourceClient),
      sourceIdentity: Boolean(sourceIdentity),
      connectedIdentity: Boolean(connectedIdentity),
      sourceIsOfficeRuntime: isEditorOfficeRuntimeClientUrl(sourceClient?.url || event.request.referrer, releaseId),
      isNewEditorIframeNavigation,
      recoveryIdentities: recoveryIdentities.length,
      resolved: Boolean(identity),
    });
  }
  return identity;
};

const fetchEditorBrokerAsset = async (event, request, url) => {
  const path = resolveEditorBrokerAssetPath(url);
  const releaseAsset = parseEditorReleaseAssetPath(url.pathname);
  const diagnosticId = isLocalEditorHost ? `${Date.now()}-${crypto.randomUUID()}` : null;
  if (diagnosticId) {
    recordEditorBrokerMatrixDiagnostic({
      id: diagnosticId,
      event: 'start',
      path,
      mode: request.mode,
      destination: request.destination,
      clientId: Boolean(event.clientId),
      resultingClientId: Boolean(event.resultingClientId),
    });
    request.signal?.addEventListener(
      'abort',
      () => recordEditorBrokerMatrixDiagnostic({ id: diagnosticId, event: 'abort', path }),
      { once: true },
    );
  }
  if (!path || !editorResourceBroker) {
    return new Response('Office resource path is not available', {
      status: 404,
      headers: { 'cache-control': 'no-store' },
    });
  }
  const identityDiagnostics = isLocalEditorHost ? {} : null;
  const identity = await resolveEditorFetchIdentity(event, releaseAsset?.releaseId ?? null, identityDiagnostics);
  if (!identity) {
    if (diagnosticId) recordEditorBrokerMatrixDiagnostic({ id: diagnosticId, event: 'identity-failed', path });
    return new Response('Office resource service is not connected', {
      status: 503,
      headers: {
        'cache-control': 'no-store',
        'retry-after': '1',
        ...(identityDiagnostics ? { 'x-onlyoffice-matrix-broker-identity': JSON.stringify(identityDiagnostics) } : {}),
      },
    });
  }
  let lastBrokerError = null;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    if (request.signal?.aborted) break;
    try {
      const response = await editorResourceBroker.fetchAsset(request, path, {
        identity,
        connectionTimeoutMs: EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS,
      });
      if (diagnosticId) {
        recordEditorBrokerMatrixDiagnostic({ id: diagnosticId, event: 'response', path, status: response.status });
      }
      return response;
    } catch (error) {
      if (diagnosticId) {
        recordEditorBrokerMatrixDiagnostic({
          id: diagnosticId,
          event: 'error',
          path,
          name: error instanceof Error ? error.name : typeof error,
          message: error instanceof Error ? error.message.slice(0, 200) : null,
        });
      }
      lastBrokerError = error;
      const retryableConnectionFailure =
        error instanceof EditorResourceBrokerError &&
        error.stage === 'connection' &&
        (error.code === 'connection' || error.code === 'replaced');
      if (attempt === 0 && retryableConnectionFailure) continue;
      break;
    }
  }
  return new Response('Office resource service is unavailable', {
    status: 503,
    headers: {
      'cache-control': 'no-store',
      'retry-after': '1',
      ...(isLocalEditorHost
        ? {
            'x-onlyoffice-matrix-broker-error': JSON.stringify({
              name: lastBrokerError instanceof Error ? lastBrokerError.name : typeof lastBrokerError,
              message: lastBrokerError instanceof Error ? lastBrokerError.message.slice(0, 240) : null,
              stack:
                lastBrokerError instanceof Error
                  ? (lastBrokerError.stack || '').split('\n').slice(0, 3).join(' | ').slice(0, 480)
                  : null,
              code: lastBrokerError instanceof EditorResourceBrokerError ? lastBrokerError.code : null,
              stage: lastBrokerError instanceof EditorResourceBrokerError ? lastBrokerError.stage : null,
              connection: editorResourceBroker.connectionState.status,
            }),
          }
        : {}),
    },
  });
};

const rejectEditorBrokerBind = (ports) => {
  const replyPort = ports[1];
  try {
    replyPort?.postMessage({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE,
      ok: false,
      code: 'identity',
    });
  } catch {
    // Untrusted ports are closed below.
  }
  for (const port of ports) {
    try {
      port.close();
    } catch {
      // Ignore malformed transferables.
    }
  }
};

self.addEventListener('message', (event) => {
  const brokerRequest = parseResourceBrokerClientMessage(event.data);
  if (brokerRequest?.type === 'PROBE' || brokerRequest?.type === 'READ') {
    const port = event.ports[0];
    const sourceUrl = typeof event.source?.url === 'string' ? event.source.url : '';
    const identity = parseCanonicalResourceBrokerClientUrl(sourceUrl, self.location.origin, isLocalCanonicalPwaHost);
    const trusted =
      canonicalResourceBroker &&
      port &&
      event.ports.length === 1 &&
      identity &&
      identity.releaseId === brokerRequest.releaseId &&
      identity.sessionId === brokerRequest.sessionId;
    if (!trusted) {
      if (port) {
        port.postMessage({
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'ERROR',
          id: brokerRequest.id,
          code: 'protocol',
        });
        port.close();
      }
      return;
    }
    event.waitUntil(canonicalResourceBroker.handle(brokerRequest, port));
  } else if (
    isIsolatedEditorHost &&
    event.data?.protocol === RESOURCE_BROKER_PROTOCOL &&
    event.data?.type === EDITOR_RESOURCE_BROKER_BIND_TYPE
  ) {
    const bind = parseEditorResourceBrokerBindMessage(event.data);
    const sourceUrl = typeof event.source?.url === 'string' ? event.source.url : '';
    const sourceClientId = typeof event.source?.id === 'string' ? event.source.id : '';
    const sourceIdentity = parseEditorOfficeHostClientIdentity(sourceUrl);
    const primeIdentity = parseEditorPrimeClientIdentity(sourceUrl);
    const primeCapabilityKey =
      primeIdentity && sourceClientId
        ? `${sourceClientId}\n${primeIdentity.releaseId}\n${primeIdentity.sessionId}`
        : '';
    const primeCapability = primeCapabilityKey ? editorPrimeBindingCapabilities.get(primeCapabilityKey) : undefined;
    const trustedPrime =
      primeIdentity &&
      primeCapability &&
      primeCapability > Date.now() &&
      primeIdentity.releaseId === bind?.releaseId &&
      primeIdentity.sessionId === bind?.sessionId;
    const trusted =
      editorResourceBroker &&
      bind &&
      sourceClientId &&
      ((sourceIdentity && sourceIdentity.releaseId === bind.releaseId && sourceIdentity.sessionId === bind.sessionId) ||
        trustedPrime) &&
      event.ports.length === 2;
    if (!trusted) {
      rejectEditorBrokerBind(event.ports);
      return;
    }
    const ports = [...event.ports];
    const bindTask = editorBrokerBindQueue.then(async () => {
      const windowClients = await self.clients.matchAll({
        includeUncontrolled: true,
        type: 'window',
      });
      const currentSource = windowClients.find((client) => client.id === sourceClientId);
      const currentSourceIdentity = parseEditorOfficeHostClientIdentity(currentSource?.url || '');
      const currentPrimeIdentity = parseEditorPrimeClientIdentity(currentSource?.url || '');
      const liveHosts = windowClients.flatMap((client) => {
        const identity = parseEditorOfficeHostClientIdentity(client.url);
        return identity ? [{ clientId: client.id, identity }] : [];
      });
      const currentPrimeCapabilityKey =
        currentPrimeIdentity && currentSource
          ? `${currentSource.id}\n${currentPrimeIdentity.releaseId}\n${currentPrimeIdentity.sessionId}`
          : '';
      const currentPrimeCapability = currentPrimeCapabilityKey
        ? editorPrimeBindingCapabilities.get(currentPrimeCapabilityKey)
        : undefined;
      const hostAuthorized =
        currentSourceIdentity &&
        currentSourceIdentity.releaseId === bind.releaseId &&
        currentSourceIdentity.sessionId === bind.sessionId &&
        editorClientIdentities?.canBindHost(sourceClientId, bind, liveHosts);
      const primeAuthorized =
        currentPrimeIdentity &&
        currentPrimeCapability &&
        currentPrimeCapability > Date.now() &&
        currentPrimeIdentity.releaseId === bind.releaseId &&
        currentPrimeIdentity.sessionId === bind.sessionId &&
        editorClientIdentities?.canBindPrime(bind, liveHosts);
      const commitHostIdentity = () => {
        if (hostAuthorized) editorClientIdentities.bindHost(sourceClientId, bind);
      };
      if (
        (!hostAuthorized && !primeAuthorized) ||
        !editorResourceBroker.handleBindMessage(bind, ports, commitHostIdentity)
      ) {
        rejectEditorBrokerBind(ports);
        return;
      }
      if (currentPrimeCapabilityKey) editorPrimeBindingCapabilities.delete(currentPrimeCapabilityKey);
    });
    editorBrokerBindQueue = bindTask.catch(() => {
      rejectEditorBrokerBind(ports);
    });
    event.waitUntil(bindTask);
  } else if (
    isIsolatedEditorHost &&
    event.data?.protocol === RESOURCE_BROKER_PROTOCOL &&
    (event.data?.type === 'ONLYOFFICE_PRIME_EDITOR_SHELL' || event.data?.type === 'ONLYOFFICE_VERIFY_EDITOR_SHELL') &&
    Object.keys(event.data).length === 4 &&
    isResourceBrokerReleaseId(event.data.releaseId)
  ) {
    const port = event.ports[0];
    let sourceUrl;
    try {
      sourceUrl = new URL(typeof event.source?.url === 'string' ? event.source.url : '');
    } catch {
      sourceUrl = null;
    }
    let canonicalOrigin = null;
    try {
      const candidate = new URL(event.data.canonicalOrigin);
      const localCanonical =
        isLocalEditorHost &&
        candidate.protocol === self.location.protocol &&
        candidate.hostname === 'onlyoffice.localhost' &&
        candidate.port === self.location.port;
      if (candidate.origin === 'https://onlyoffice.getpi.work' || localCanonical) {
        canonicalOrigin = candidate.origin;
      }
    } catch {
      canonicalOrigin = null;
    }
    const trusted =
      port &&
      event.ports.length === 1 &&
      canonicalOrigin &&
      sourceUrl?.origin === self.location.origin &&
      releaseIdFromEditorShellPath(sourceUrl.pathname, EDITOR_SHELL_PRIME_PATH) === event.data.releaseId &&
      (event.data.type === 'ONLYOFFICE_PRIME_EDITOR_SHELL'
        ? sourceUrl.searchParams.get(EDITOR_SHELL_PRIME_INSTALL_QUERY) === '1' &&
          [...sourceUrl.searchParams.keys()].length === 1
        : sourceUrl.search === '');
    if (!trusted) {
      port?.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_ERROR',
        releaseId: event.data.releaseId,
        code: 'identity',
      });
      port?.close();
      return;
    }
    let prepare;
    const existingPrime = editorShellPrimeInFlight.get(event.data.releaseId);
    if (event.data.type === 'ONLYOFFICE_PRIME_EDITOR_SHELL') {
      if (existingPrime) {
        prepare = existingPrime;
      } else {
        prepare = primeEditorShell({
          releaseId: event.data.releaseId,
          origin: self.location.origin,
          manifestOrigin: canonicalOrigin,
          cacheStorage: caches,
          fetch: (...args) => fetch(...args),
        });
        editorShellPrimeInFlight.set(event.data.releaseId, prepare);
        const clearPrime = () => {
          if (editorShellPrimeInFlight.get(event.data.releaseId) === prepare) {
            editorShellPrimeInFlight.delete(event.data.releaseId);
          }
        };
        void prepare.then(clearPrime, clearPrime);
      }
    } else {
      prepare = (existingPrime ?? Promise.resolve()).then(() =>
        verifyEditorShell({
          releaseId: event.data.releaseId,
          origin: self.location.origin,
          cacheStorage: caches,
        }),
      );
    }
    event.waitUntil(
      prepare
        .then((result) => {
          port.postMessage({
            protocol: RESOURCE_BROKER_PROTOCOL,
            type: 'ONLYOFFICE_EDITOR_SHELL_PRIMED',
            releaseId: result.releaseId,
            origin: self.location.origin,
            serviceWorkerVersion: SERVICE_WORKER_VERSION,
            cachedPaths: result.cachedPaths,
            cachedBytes: result.cachedBytes,
          });
        })
        .catch((error) => {
          const detail = isLocalEditorHost && error instanceof Error ? error.message.slice(0, 200) : undefined;
          port.postMessage({
            protocol: RESOURCE_BROKER_PROTOCOL,
            type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_ERROR',
            releaseId: event.data.releaseId,
            code: 'storage',
            ...(detail ? { detail } : {}),
          });
        })
        .finally(() => port.close()),
    );
  } else if (event.data?.type === 'SKIP_WAITING') {
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

// The stable pointer remains network-only and no-store. When Chrome has not
// yet updated navigator.onLine during an offline restart, convert the failed
// subrequest into a typed response so the installed release stays usable and
// the page does not emit ERR_INTERNET_DISCONNECTED console noise.
registerRoute(
  ({ request, sameOrigin, url }) =>
    sameOrigin &&
    isCanonicalPwaHost &&
    request.method === 'GET' &&
    !url.search &&
    url.pathname === '/channels/stable-v5.json',
  async ({ request }) => {
    try {
      return await fetch(request);
    } catch {
      return new Response(JSON.stringify({ code: 'offline' }), {
        status: 503,
        headers: {
          'cache-control': 'no-store',
          'content-type': 'application/json; charset=utf-8',
          'retry-after': '1',
        },
      });
    }
  },
  'GET',
);

// Broker capabilities live in the query string, so Workbox's default precache
// matcher cannot resolve this navigation to the query-less shell while
// offline. Serve the immutable HTML shell explicitly; resource-broker.ts still
// validates every query field, referrer, physical origin, and capability.
registerRoute(
  ({ request, sameOrigin, url }) =>
    sameOrigin && isCanonicalPwaHost && request.mode === 'navigate' && url.pathname === '/resource-broker.html',
  async () => (await matchPrecache('/resource-broker.html')) || fetch('/resource-broker.html'),
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

const editorShellRouteMatcher = ({ sameOrigin, url }) => {
  const releaseAsset = parseEditorReleaseAssetPath(url.pathname);
  return Boolean(
    sameOrigin && isIsolatedEditorHost && !url.search && releaseAsset && isEditorShellAssetPath(releaseAsset.path),
  );
};

const editorShellRouteHandler = async ({ event, request, url }) => {
  const releaseAsset = parseEditorReleaseAssetPath(url.pathname);
  const cached = releaseAsset && (await matchEditorShell(request, releaseAsset.releaseId, caches));
  if (cached) return cached;
  if (!releaseAsset) {
    return new Response('Editor shell path is invalid', {
      status: 404,
      headers: { 'cache-control': 'no-store' },
    });
  }
  if (releaseAsset.path === EDITOR_SHELL_HOST_PATH || releaseAsset.path === EDITOR_SHELL_PRIME_PATH) {
    let detail = '';
    if (isLocalEditorHost) {
      try {
        await verifyEditorShell({
          releaseId: releaseAsset.releaseId,
          origin: self.location.origin,
          cacheStorage: caches,
        });
        detail = ' (verified generation did not match the request)';
      } catch (error) {
        detail = ` (${error instanceof Error ? error.message : String(error)})`.slice(0, 220);
      }
    }
    return new Response(`Editor shell is not prepared${detail}`, {
      status: 503,
      headers: { 'cache-control': 'no-store', 'retry-after': '1' },
    });
  }
  if (editorBootstrapAssetPaths.has(`/${releaseAsset.path}`)) {
    const response = await fetch(request, {
      cache: 'no-store',
      credentials: 'omit',
      redirect: 'error',
    });
    if (response.ok && response.status === 200) {
      const cache = await caches.open(EDITOR_SHELL_CACHE_NAME);
      event.waitUntil(cache.put(url.href, response.clone()));
    }
    return response;
  }
  return fetchEditorBrokerAsset(event, request, url);
};

registerRoute(editorShellRouteMatcher, editorShellRouteHandler, 'GET');
registerRoute(editorShellRouteMatcher, editorShellRouteHandler, 'HEAD');

const editorBrokerRouteMatcher = ({ sameOrigin, url }) =>
  sameOrigin &&
  isIsolatedEditorHost &&
  !isEditorDirectOriginBoundPath(url) &&
  !(isLocalEditorHost && isExcludedApplicationPath(url));

const editorBrokerRouteHandler = ({ event, request, url }) => fetchEditorBrokerAsset(event, request, url);

// Every release-manifest Office asset is streamed from the canonical broker.
// There is deliberately no direct-network fallback on an editor origin.
registerRoute(editorBrokerRouteMatcher, editorBrokerRouteHandler, 'GET');
registerRoute(editorBrokerRouteMatcher, editorBrokerRouteHandler, 'HEAD');

registerRoute(
  ({ request, sameOrigin, url }) => {
    if (
      !sameOrigin ||
      isIsolatedEditorHost ||
      url.searchParams.has(SHARED_ASSET_VERSION_QUERY) ||
      !isEligibleGet(request, url) ||
      isFontRequest(url) ||
      isCanonicalManagedResourceRequest(url)
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
    !isIsolatedEditorHost &&
    !url.searchParams.has(SHARED_ASSET_VERSION_QUERY) &&
    isEligibleGet(request, url) &&
    !isFontRequest(url) &&
    !isCanonicalManagedResourceRequest(url),
  staticStaleWhileRevalidate,
  'GET',
);
