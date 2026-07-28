export {};

const PREWARM_MESSAGE = 'onlyoffice-slot-prewarm-v1';
const RUNTIME_MANIFEST_PATH = '/onlyoffice-runtime-assets.json';
const HOST_DOCUMENT_PATH = '/office-host.html';
const statusElement = document.querySelector<HTMLElement>('#status');

type RuntimeManifest = {
  assets?: Array<{ path?: string }>;
};

function report(state: 'ready' | 'error', detail?: string): void {
  if (statusElement) {
    statusElement.textContent =
      state === 'ready' ? 'Offline editor slot ready.' : detail || 'Prewarm failed.';
  }
  const parentOrigin = new URLSearchParams(window.location.search).get('parentOrigin');
  if (!parentOrigin) return;
  window.parent.postMessage(
    {
      type: PREWARM_MESSAGE,
      state,
      origin: window.location.origin,
      detail,
    },
    parentOrigin,
  );
}

function safePath(path: unknown): path is string {
  return typeof path === 'string' && path.length > 0 && !path.startsWith('/') && !path.includes('..');
}

async function hostShellPaths(): Promise<string[]> {
  const hostUrl = new URL(HOST_DOCUMENT_PATH, window.location.origin);
  hostUrl.searchParams.set('__oobslot', 'shell');
  const response = await fetch(hostUrl, { cache: 'reload' });
  if (!response.ok) throw new Error(`Office host prewarm failed (${response.status}).`);
  const html = await response.text();
  const parsed = new DOMParser().parseFromString(html, 'text/html');
  const paths = [HOST_DOCUMENT_PATH, '/reset.html', '/document_editor_service_worker.js', '/sw.js', '/plugins.json', '/themes.json'];
  for (const element of parsed.querySelectorAll<HTMLScriptElement | HTMLLinkElement>('script[src], link[href]')) {
    const value = element.getAttribute(element instanceof HTMLScriptElement ? 'src' : 'href');
    if (!value) continue;
    const url = new URL(value, hostUrl);
    if (url.origin === window.location.origin) paths.push(url.pathname);
  }
  return paths;
}

async function runtimeShellPaths(): Promise<{ paths: string[]; version: string }> {
  const manifestUrl = new URL(RUNTIME_MANIFEST_PATH, window.location.origin);
  manifestUrl.searchParams.set('__cache_status', String(Date.now()));
  const response = await fetch(manifestUrl, { cache: 'reload' });
  if (!response.ok) throw new Error(`Runtime manifest prewarm failed (${response.status}).`);
  const manifest = (await response.json()) as RuntimeManifest;
  const version =
    response.headers.get('X-OnlyOffice-Asset-Version') ||
    response.headers.get('etag')?.replaceAll('"', '');
  if (!version) throw new Error('Runtime asset version is unavailable.');
  const paths = (manifest.assets || [])
    .map((asset) => asset.path)
    .filter(safePath)
    .filter(
      (path) =>
        path.endsWith('.html') ||
        /^wasm\/x2t\/(?:conversion-worker|startup-heartbeat-worker)-.+\.js$/.test(path),
    )
    .map((path) => `/${path}`);
  return { paths, version };
}

async function sendPrewarm(
  worker: ServiceWorker,
  version: string,
  paths: string[],
): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const channel = new MessageChannel();
    const timeoutId = window.setTimeout(() => reject(new Error('Offline slot prewarm timed out.')), 60_000);
    channel.port1.onmessage = (event) => {
      window.clearTimeout(timeoutId);
      if (event.data?.ok === true) resolve();
      else reject(new Error(event.data?.error || 'Offline slot prewarm failed.'));
    };
    worker.postMessage({ type: PREWARM_MESSAGE, version, paths }, [channel.port2]);
  });
}

async function prewarm(): Promise<void> {
  if (!('serviceWorker' in navigator)) throw new Error('Service Worker is unavailable.');
  const [{ paths: runtimePaths, version }, shellPaths] = await Promise.all([
    runtimeShellPaths(),
    hostShellPaths(),
  ]);
  const registration = await navigator.serviceWorker.register('/document_editor_service_worker.js', {
    scope: '/',
    updateViaCache: 'none',
  });
  await navigator.serviceWorker.ready;
  const worker = registration.active || registration.waiting || registration.installing;
  if (!worker) throw new Error('Offline slot Service Worker is unavailable.');
  await sendPrewarm(worker, version, [...new Set([...shellPaths, ...runtimePaths])]);
}

void prewarm()
  .then(() => report('ready'))
  .catch((error) => report('error', error instanceof Error ? error.message : String(error)));
