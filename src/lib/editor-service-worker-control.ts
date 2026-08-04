const DEFAULT_CONTROL_TIMEOUT_MS = 30_000;

export type EditorServiceWorkerRegistration = Pick<
  ServiceWorkerRegistration,
  'active' | 'waiting' | 'installing' | 'update' | 'addEventListener' | 'removeEventListener'
>;

export type EditorServiceWorkerContainer = Pick<
  ServiceWorkerContainer,
  'controller' | 'register' | 'ready' | 'addEventListener' | 'removeEventListener'
>;

function waitForInstalledWorker(
  worker: ServiceWorker,
  timeoutMs: number,
  setTimer: typeof window.setTimeout,
  clearTimer: typeof window.clearTimeout,
): Promise<ServiceWorker | null> {
  if (worker.state === 'installed') return Promise.resolve(worker);
  if (worker.state === 'redundant') return Promise.resolve(null);

  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      clearTimer(timeoutId);
      worker.removeEventListener('statechange', onStateChange);
      callback();
    };
    const onStateChange = () => {
      if (worker.state === 'installed') finish(() => resolve(worker));
      else if (worker.state === 'redundant') finish(() => resolve(null));
    };
    const timeoutId = setTimer(() => {
      finish(() => reject(new Error('Editor Service Worker installation timed out')));
    }, timeoutMs);
    worker.addEventListener('statechange', onStateChange);
  });
}

function waitForControllerChange(
  serviceWorker: EditorServiceWorkerContainer,
  previousController: ServiceWorker | null,
  timeoutMs: number,
  setTimer: typeof window.setTimeout,
  clearTimer: typeof window.clearTimeout,
): Promise<ServiceWorker> {
  const current = serviceWorker.controller;
  if (current && current !== previousController) return Promise.resolve(current);

  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (callback: () => void) => {
      if (settled) return;
      settled = true;
      clearTimer(timeoutId);
      serviceWorker.removeEventListener('controllerchange', onControllerChange);
      callback();
    };
    const onControllerChange = () => {
      const controller = serviceWorker.controller;
      if (controller && controller !== previousController) finish(() => resolve(controller));
    };
    const timeoutId = setTimer(() => {
      finish(() => reject(new Error('Editor Service Worker control timed out')));
    }, timeoutMs);
    serviceWorker.addEventListener('controllerchange', onControllerChange);
    onControllerChange();
  });
}

async function activateWaitingWorker(
  registration: EditorServiceWorkerRegistration,
  serviceWorker: EditorServiceWorkerContainer,
  timeoutMs: number,
  setTimer: typeof window.setTimeout,
  clearTimer: typeof window.clearTimeout,
): Promise<ServiceWorker | null> {
  let waiting = registration.waiting;
  if (!waiting && registration.installing) {
    const installed = await waitForInstalledWorker(registration.installing, timeoutMs, setTimer, clearTimer);
    waiting = registration.waiting ?? installed;
  }
  if (!waiting) return null;

  const previousController = serviceWorker.controller;
  const controllerChange = waitForControllerChange(serviceWorker, previousController, timeoutMs, setTimer, clearTimer);
  waiting.postMessage({ type: 'SKIP_WAITING' });
  return controllerChange;
}

/**
 * Registers the isolated editor Service Worker, checks for a newer script,
 * and activates a waiting worker before returning a controller. Without this
 * handshake a freshly deployed editor origin can run one prime against the
 * old worker and report a misleading shell-cache/storage failure.
 */
export async function ensureControlledEditorServiceWorker(
  serviceWorker: EditorServiceWorkerContainer,
  scriptUrl: string,
  options: {
    scope?: string;
    updateViaCache?: RegistrationOptions['updateViaCache'];
    timeoutMs?: number;
    setTimer?: typeof window.setTimeout;
    clearTimer?: typeof window.clearTimeout;
  } = {},
): Promise<ServiceWorker> {
  const timeoutMs = options.timeoutMs ?? DEFAULT_CONTROL_TIMEOUT_MS;
  const setTimer = options.setTimer ?? window.setTimeout.bind(window);
  const clearTimer = options.clearTimer ?? window.clearTimeout.bind(window);
  const registration = await serviceWorker.register(scriptUrl, {
    scope: options.scope ?? '/',
    updateViaCache: options.updateViaCache ?? 'none',
  });
  await serviceWorker.ready;

  // An editor prime iframe is short-lived. Force the update check here rather
  // than waiting for a future navigation, then activate the waiting worker so
  // the first request after deployment cannot use stale routing code.
  try {
    await registration.update();
  } catch {
    // An existing active controller remains usable while offline. If there is
    // no controller, the control wait below still reports a typed timeout.
  }

  const activated = await activateWaitingWorker(registration, serviceWorker, timeoutMs, setTimer, clearTimer);
  if (activated) return activated;
  if (serviceWorker.controller) return serviceWorker.controller;
  return waitForControllerChange(serviceWorker, null, timeoutMs, setTimer, clearTimer);
}
