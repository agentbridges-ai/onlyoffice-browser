const params = new URLSearchParams(location.search);
const fault = params.get('fault') || 'none';
const releaseId = 'matrix-shell-release-v1';
const timeoutMs = 800;
const resultElement = document.querySelector('#result');
const stages = [];

const publish = (value) => {
  const result = { ...value, fault, origin: location.origin, stages: [...stages] };
  if (resultElement) resultElement.textContent = JSON.stringify(result);
  globalThis.__ONLYOFFICE_EDITOR_SHELL_FAULT_RESULT__ = result;
  return result;
};

const failure = (stage, code, detail) => {
  const error = new Error(detail || `${stage}/${code}`);
  error.stage = stage;
  error.code = code;
  return error;
};

async function controlledWorker() {
  stages.push('service-worker');
  const scriptUrl = new URL('/__matrix__/editor-shell-fault/sw.js', location.origin);
  scriptUrl.searchParams.set('fault', fault);
  const registration = await navigator.serviceWorker.register(scriptUrl.href, {
    scope: '/__matrix__/editor-shell-fault/',
    updateViaCache: 'none',
  });
  await navigator.serviceWorker.ready;
  if (navigator.serviceWorker.controller) return navigator.serviceWorker.controller;
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      navigator.serviceWorker.removeEventListener('controllerchange', changed);
      reject(failure('service-worker', 'timeout', 'controllerchange timeout'));
    }, timeoutMs);
    const changed = () => {
      const worker = navigator.serviceWorker.controller || registration.active;
      if (!worker) return;
      clearTimeout(timer);
      resolve(worker);
    };
    navigator.serviceWorker.addEventListener('controllerchange', changed, { once: true });
  });
}

function request(worker, type) {
  const channel = new MessageChannel();
  const promise = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      channel.port1.close();
      reject(failure(type === 'ONLYOFFICE_PRIME_EDITOR_SHELL' ? 'shell-cache' : 'broker-probe', 'timeout'));
    }, timeoutMs);
    channel.port1.onmessage = (event) => {
      clearTimeout(timer);
      channel.port1.close();
      const value = event.data;
      if (!value || typeof value !== 'object') {
        reject(failure('shell-cache', 'storage', 'invalid worker response'));
        return;
      }
      resolve(value);
    };
    channel.port1.start();
  });
  worker.postMessage({ protocol: 'onlyoffice-editor-shell-fault-v1', type, releaseId }, [channel.port2]);
  return promise;
}

async function verifyShellRoute(storageMode) {
  stages.push('shell-route');
  const response = await fetch(`${location.origin}/__matrix__/editor-shell-fault/office-host.html`, {
    cache: 'no-store',
    credentials: 'omit',
  });
  const responseRelease = response.headers.get('x-onlyoffice-asset-version');
  const expectedStorage =
    storageMode === 'cache'
      ? response.headers.get('x-onlyoffice-editor-shell-cache') === '1'
      : response.headers.get('x-onlyoffice-editor-shell-storage') === 'network';
  await response.body?.cancel().catch(() => undefined);
  if (response.status !== 200 || responseRelease !== releaseId || !expectedStorage) {
    throw failure('shell-route', response.status === 404 ? 'network' : 'integrity', 'shell route mismatch');
  }
}

async function run() {
  try {
    const worker = await controlledWorker();
    stages.push('shell-cache');
    const prime = await request(worker, 'ONLYOFFICE_PRIME_EDITOR_SHELL');
    if (
      prime?.type !== 'ONLYOFFICE_EDITOR_SHELL_PRIMED' ||
      (prime.storageMode !== 'cache' && prime.storageMode !== 'network')
    ) {
      throw failure('shell-cache', prime?.code || 'storage', 'shell prime failed');
    }
    await verifyShellRoute(prime.storageMode);
    stages.push('broker-probe');
    const broker = await request(worker, 'ONLYOFFICE_BROKER_PROBE');
    if (broker?.type !== 'ONLYOFFICE_BROKER_PROBE_RESULT' || broker.ok !== true) {
      throw failure('broker-probe', broker?.code || 'storage', 'Broker probe failed');
    }
    return publish({ ready: true, storageMode: prime.storageMode });
  } catch (error) {
    const value = error && typeof error === 'object' ? error : {};
    const stage = ['service-worker', 'shell-cache', 'shell-route', 'broker-probe'].includes(value.stage)
      ? value.stage
      : 'service-worker';
    return publish({ ready: false, code: value.code || 'storage', stage, detail: value.message || String(error) });
  }
}

globalThis.__ONLYOFFICE_EDITOR_SHELL_FAULT_RESULT_PROMISE__ = run();
