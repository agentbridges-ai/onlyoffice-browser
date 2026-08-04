import {
  EDITOR_SHELL_CACHE_RESPONSE_HEADER,
  EDITOR_SHELL_HOST_PATH,
  releaseIdFromEditorShellPath,
} from './lib/editor-shell-cache';
import {
  ResourceBrokerFrameClient,
  resolveResourceBrokerPhysicalEditorIdentity,
} from './lib/resource-broker-frame-client';
import {
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  parseResourceBrokerServerMessage,
} from './lib/resource-broker-protocol';
import { ensureControlledEditorServiceWorker } from './lib/editor-service-worker-control';

const SERVICE_WORKER_PATH = '/document_editor_service_worker.js';
const TIMEOUT_MS = 30_000;
const SHELL_ROUTE_PROBE_TIMEOUT_MS = 5_000;

async function probeInstalledShellRoute(releaseId: string, storageMode: 'cache' | 'network'): Promise<void> {
  const target = new URL(`/r/${encodeURIComponent(releaseId)}/${EDITOR_SHELL_HOST_PATH}`, location.origin);
  const deadline = Date.now() + SHELL_ROUTE_PROBE_TIMEOUT_MS;
  let delayMs = 25;
  let lastStatus = 0;
  do {
    const response = await fetch(target, {
      cache: 'no-store',
      credentials: 'omit',
      redirect: 'error',
    });
    lastStatus = response.status;
    const servedByEditorCache = response.headers.get(EDITOR_SHELL_CACHE_RESPONSE_HEADER) === '1';
    const servedByNetworkFallback = response.headers.get('x-onlyoffice-editor-shell-storage') === 'network';
    const responseRelease = response.headers.get('x-onlyoffice-asset-version');
    if (
      response.ok &&
      response.status === 200 &&
      (storageMode === 'cache' ? servedByEditorCache : servedByNetworkFallback) &&
      responseRelease === releaseId
    ) {
      await response.body?.cancel().catch(() => undefined);
      return;
    }
    await response.body?.cancel().catch(() => undefined);
    if (Date.now() >= deadline) break;
    await new Promise((resolve) => window.setTimeout(resolve, delayMs));
    delayMs = Math.min(delayMs * 2, 500);
  } while (Date.now() < deadline);
  throw new Error(`Editor shell route probe failed (${lastStatus || 'unavailable'})`);
}

function bootstrap(): {
  parentOrigin: string;
  releaseId: string;
  sessionId: string;
  localMatrix: boolean;
  mode: 'install' | 'verify';
} | null {
  const fragment = new URLSearchParams(location.hash.startsWith('#') ? location.hash.slice(1) : location.hash);
  if ([...fragment.keys()].some((key) => !['parentOrigin', 'releaseId', 'sessionId', 'mode'].includes(key))) {
    return null;
  }
  if (
    fragment.getAll('parentOrigin').length !== 1 ||
    fragment.getAll('releaseId').length !== 1 ||
    fragment.getAll('sessionId').length !== 1 ||
    fragment.getAll('mode').length !== 1
  ) {
    return null;
  }
  const parentOrigin = fragment.get('parentOrigin');
  const releaseId = fragment.get('releaseId');
  const sessionId = fragment.get('sessionId');
  const mode = fragment.get('mode');
  const pathReleaseId = releaseIdFromEditorShellPath(location.pathname, 'editor-shell-prime.html');
  if (
    !parentOrigin ||
    !releaseId ||
    !isResourceBrokerReleaseId(releaseId) ||
    !isResourceBrokerSessionId(sessionId) ||
    (mode !== 'install' && mode !== 'verify') ||
    releaseId !== pathReleaseId
  ) {
    return null;
  }

  let parent: URL;
  try {
    parent = new URL(parentOrigin);
  } catch {
    return null;
  }
  const localMatrix =
    parent.hostname === 'onlyoffice.localhost' &&
    parent.protocol === location.protocol &&
    parent.port === location.port;
  if (parent.origin !== 'https://onlyoffice.getpi.work' && !localMatrix) return null;
  if (!resolveResourceBrokerPhysicalEditorIdentity(location.origin, parent.origin, localMatrix)) return null;
  try {
    if (!document.referrer || new URL(document.referrer).origin !== parent.origin) return null;
  } catch {
    return null;
  }
  return { parentOrigin: parent.origin, releaseId, sessionId, localMatrix, mode };
}

async function controlledServiceWorker(): Promise<ServiceWorker> {
  return ensureControlledEditorServiceWorker(navigator.serviceWorker, SERVICE_WORKER_PATH, {
    timeoutMs: TIMEOUT_MS,
  });
}

async function prime(): Promise<void> {
  const identity = bootstrap();
  if (!identity || !('serviceWorker' in navigator)) throw new Error('Invalid editor shell prime identity');
  const progress = (stage: 'service-worker' | 'shell-cache' | 'shell-route' | 'broker-probe') =>
    window.parent.postMessage(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_PROGRESS',
        origin: location.origin,
        releaseId: identity.releaseId,
        sessionId: identity.sessionId,
        stage,
      },
      identity.parentOrigin,
    );
  progress('service-worker');
  const worker = await controlledServiceWorker();
  progress('shell-cache');
  const result = await new Promise<Record<string, unknown>>((resolve, reject) => {
    const channel = new MessageChannel();
    const timeout = window.setTimeout(() => {
      channel.port1.close();
      reject(new Error('Editor shell prime request timed out'));
    }, TIMEOUT_MS);
    channel.port1.onmessage = (event) => {
      window.clearTimeout(timeout);
      channel.port1.close();
      if (!event.data || typeof event.data !== 'object') {
        reject(new Error('Invalid editor shell prime response'));
        return;
      }
      resolve(event.data as Record<string, unknown>);
    };
    channel.port1.onmessageerror = () => {
      window.clearTimeout(timeout);
      channel.port1.close();
      reject(new Error('Editor shell prime response failed'));
    };
    channel.port1.start();
    worker.postMessage(
      {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: identity.mode === 'install' ? 'ONLYOFFICE_PRIME_EDITOR_SHELL' : 'ONLYOFFICE_VERIFY_EDITOR_SHELL',
        releaseId: identity.releaseId,
        canonicalOrigin: identity.parentOrigin,
      },
      [channel.port2],
    );
  });
  if (
    result.protocol !== RESOURCE_BROKER_PROTOCOL ||
    result.type !== 'ONLYOFFICE_EDITOR_SHELL_PRIMED' ||
    result.releaseId !== identity.releaseId ||
    result.origin !== location.origin ||
    typeof result.serviceWorkerVersion !== 'string' ||
    result.serviceWorkerVersion.length === 0 ||
    !Array.isArray(result.cachedPaths) ||
    !Number.isSafeInteger(result.cachedBytes) ||
    (result.storageMode !== 'cache' && result.storageMode !== 'network')
  ) {
    const code = typeof result.code === 'string' ? result.code : 'invalid-response';
    const detail = typeof result.detail === 'string' ? `: ${result.detail}` : '';
    throw new Error(`Editor shell prime did not complete (${code}${detail})`);
  }

  progress('shell-route');
  await probeInstalledShellRoute(identity.releaseId, result.storageMode as 'cache' | 'network');

  const broker = new ResourceBrokerFrameClient({
    releaseId: identity.releaseId,
    sessionId: identity.sessionId,
    parentOrigin: identity.parentOrigin,
    canonicalOrigin: identity.parentOrigin,
    allowLocalMatrix: identity.localMatrix,
  });
  try {
    progress('broker-probe');
    const connection = await broker.connect();
    const probeId = `prime-${crypto.randomUUID()}`;
    await new Promise<void>((resolve, reject) => {
      const timeout = window.setTimeout(() => {
        connection.port.close();
        reject(new Error('Resource Broker read probe timed out'));
      }, TIMEOUT_MS);
      const finish = (callback: () => void) => {
        window.clearTimeout(timeout);
        connection.port.onmessage = null;
        connection.port.onmessageerror = null;
        callback();
      };
      connection.port.onmessage = (event) => {
        const message = parseResourceBrokerServerMessage(event.data);
        if (
          !message ||
          message.id !== probeId ||
          (message.type !== 'RESULT' && message.type !== 'ERROR') ||
          (message.type === 'RESULT' && message.value.releaseId !== identity.releaseId)
        ) {
          finish(() => reject(new Error('Invalid Resource Broker read probe response')));
          return;
        }
        if (message.type === 'ERROR') {
          finish(() => reject(new Error(`Resource Broker read probe failed: ${message.code}`)));
          return;
        }
        finish(resolve);
      };
      connection.port.onmessageerror = () =>
        finish(() => reject(new Error('Resource Broker read probe channel failed')));
      connection.port.start();
      connection.port.postMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PROBE',
        id: probeId,
        releaseId: identity.releaseId,
        sessionId: identity.sessionId,
      });
    });
  } finally {
    broker.destroy();
  }

  window.parent.postMessage(
    {
      ...result,
      sessionId: identity.sessionId,
      brokerReady: true,
      occupied: false,
      storageMode: result.storageMode,
    },
    identity.parentOrigin,
  );
}

void prime().catch((error) => {
  console.error('[onlyoffice-browser] Failed to prime the editor origin shell', error);
  const identity = bootstrap();
  if (!identity) return;
  const message = error instanceof Error ? error.message : String(error);
  const stage = /Service Worker/i.test(message)
    ? 'service-worker'
    : /Resource Broker/i.test(message)
      ? 'broker-probe'
      : /shell route/i.test(message)
        ? 'shell-route'
        : /shell|cache/i.test(message)
          ? 'shell-cache'
          : 'service-worker';
  const code = /cancelled/i.test(message)
    ? 'cancelled'
    : /timed out/i.test(message)
      ? 'timeout'
      : /identity/i.test(message)
        ? 'identity'
        : 'storage';
  window.parent.postMessage(
    {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_FAILED',
      origin: location.origin,
      releaseId: identity.releaseId,
      sessionId: identity.sessionId,
      code,
      stage,
      ...(identity.localMatrix ? { detail: message.slice(0, 200) } : {}),
    },
    identity.parentOrigin,
  );
});
