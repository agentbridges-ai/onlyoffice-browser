const PROTOCOL = 1;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const SESSION_ID_PATTERN = /^[a-zA-Z0-9._-]{1,128}$/;
const REQUEST_ID_PATTERN = /^[a-zA-Z0-9._-]{1,128}$/;
const ALLOWED_EDITOR_HOSTS = new Set(['aries.localhost', 'taurus.localhost', 'gemini.localhost']);
const params = new URLSearchParams(location.search);
const releaseId = params.get('releaseId') || '';
const sessionId = params.get('sessionId') || '';
const parentOrigin = (() => {
  try {
    return new URL(document.referrer).origin;
  } catch {
    return '';
  }
})();
const allowedParent = (origin) => {
  try {
    const url = new URL(origin);
    return url.protocol === location.protocol && url.port === location.port && ALLOWED_EDITOR_HOSTS.has(url.hostname);
  } catch {
    return false;
  }
};
const randomCapability = () =>
  Array.from(crypto.getRandomValues(new Uint8Array(32)), (value) => value.toString(16).padStart(2, '0')).join('');
const activeRequests = new Map();
let connectionPort = null;
let challenge = randomCapability();

const ensureCanonicalWorker = async () => {
  if (!('serviceWorker' in navigator)) throw new Error('worker');
  const registration = navigator.onLine
    ? await navigator.serviceWorker.register('/__matrix__/cache-broker-shell-sw.js', {
        scope: '/__matrix__/',
        updateViaCache: 'none',
      })
    : await navigator.serviceWorker.getRegistration('/__matrix__/');
  if (!registration) throw new Error('worker-offline');
  await navigator.serviceWorker.ready;
  if (!navigator.serviceWorker.controller) {
    await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error('worker-control')), 5_000);
      navigator.serviceWorker.addEventListener(
        'controllerchange',
        () => {
          clearTimeout(timeout);
          resolve();
        },
        { once: true },
      );
    });
  }
  return navigator.serviceWorker.controller;
};

const canonicalRequest = async (message) => {
  const controller = await ensureCanonicalWorker();
  const channel = new MessageChannel();
  const result = new Promise((resolve, reject) => {
    const timeout = setTimeout(() => reject(new Error('worker-timeout')), 5_000);
    channel.port1.onmessage = (event) => {
      if (event.data?.type !== 'RESULT' && event.data?.type !== 'ERROR') return;
      clearTimeout(timeout);
      channel.port1.close();
      if (event.data.type === 'RESULT') resolve(event.data.value);
      else reject(new Error(event.data.error));
    };
    channel.port1.start();
  });
  controller.postMessage({ ...message, releaseId, sessionId }, [channel.port2]);
  return result;
};

const forwardRequest = async (parentPort, message) => {
  if (!REQUEST_ID_PATTERN.test(message?.id || '')) return;
  if (message.releaseId !== undefined && message.releaseId !== releaseId) {
    parentPort.postMessage({ type: 'ERROR', id: message.id, error: 'release' });
    return;
  }
  const state = { pending: [], port: null };
  activeRequests.set(message.id, state);
  try {
    const controller = await ensureCanonicalWorker();
    const channel = new MessageChannel();
    state.port = channel.port1;
    channel.port1.onmessage = (event) => {
      const transfer = event.data?.bytes instanceof ArrayBuffer ? [event.data.bytes] : [];
      parentPort.postMessage(event.data, transfer);
      if (['RESULT', 'ERROR', 'END', 'CANCELLED'].includes(event.data?.type)) {
        activeRequests.delete(message.id);
        channel.port1.close();
      }
    };
    channel.port1.start();
    controller.postMessage({ ...message, releaseId, sessionId }, [channel.port2]);
    for (const pendingMessage of state.pending) channel.port1.postMessage(pendingMessage);
    state.pending.length = 0;
  } catch (error) {
    activeRequests.delete(message.id);
    throw error;
  }
};

const handlePortMessage = (parentPort, message) => {
  if (message?.type === 'PULL' || message?.type === 'CANCEL') {
    const state = activeRequests.get(message.id);
    if (state?.port) state.port.postMessage(message);
    else state?.pending.push(message);
    return;
  }
  forwardRequest(parentPort, message).catch((error) => {
    parentPort.postMessage({
      type: 'ERROR',
      id: message?.id,
      error: error instanceof Error ? error.message : String(error),
    });
  });
};

addEventListener('message', (event) => {
  const port = event.ports[0];
  if (
    !port ||
    event.source !== parent ||
    event.origin !== parentOrigin ||
    !allowedParent(event.origin) ||
    event.data?.type !== 'CONNECT'
  ) {
    return;
  }
  if (
    connectionPort ||
    !challenge ||
    event.data.protocol !== PROTOCOL ||
    event.data.challenge !== challenge ||
    event.data.releaseId !== releaseId ||
    event.data.sessionId !== sessionId
  ) {
    port.postMessage({ type: 'CONNECT_ERROR', error: 'capability' });
    port.close();
    return;
  }
  challenge = '';
  connectionPort = port;
  port.onmessage = (message) => handlePortMessage(port, message.data);
  port.start();
  port.postMessage({ type: 'CONNECTED', protocol: PROTOCOL, releaseId, sessionId });
});

const identityPromise = (async () => {
  if (!allowedParent(parentOrigin) || !RELEASE_ID_PATTERN.test(releaseId) || !SESSION_ID_PATTERN.test(sessionId)) {
    throw new Error('identity');
  }
  await canonicalRequest({ type: 'PROBE', id: 'bootstrap-probe' });
})();

identityPromise
  .then(() => {
    parent.postMessage(
      {
        type: 'ONLYOFFICE_BROKER_CHALLENGE',
        protocol: PROTOCOL,
        releaseId,
        sessionId,
        challenge,
      },
      parentOrigin,
    );
  })
  .catch(() => undefined);
