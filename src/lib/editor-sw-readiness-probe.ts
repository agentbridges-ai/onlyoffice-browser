import {
  EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE,
  EDITOR_RESOURCE_BROKER_BIND_TYPE,
  EDITOR_RESOURCE_BROKER_BOUND_TYPE,
} from './editor-resource-broker';
import {
  RESOURCE_BROKER_PROTOCOL,
  parseResourceBrokerServerMessage,
  type ResourceBrokerProbeResultMessage,
} from './resource-broker-protocol';

export const EDITOR_SW_READINESS_TIMEOUT_MS = 30_000;

type Timer = ReturnType<typeof globalThis.setTimeout>;

type ReadinessPort = {
  onmessage: ((event: MessageEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  close(): void;
  postMessage(message: unknown, transfer?: Transferable[]): void;
  start(): void;
};

type ReadinessWorker = {
  postMessage(message: unknown, transfer?: Transferable[]): void;
};

export type EditorSwReadinessIdentity = {
  releaseId: string;
  sessionId: string;
};

export type EditorSwReadinessProbe = ResourceBrokerProbeResultMessage['value'];

function exactRecord(value: unknown, keys: readonly string[]): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const actual = Object.keys(value);
  return actual.length === keys.length && keys.every((key) => Object.hasOwn(value, key));
}

function readinessTimeout(
  callback: () => void,
  timeoutMs: number,
  setTimeoutImpl: typeof globalThis.setTimeout,
): Timer {
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > EDITOR_SW_READINESS_TIMEOUT_MS) {
    throw new TypeError('Invalid Editor Service Worker readiness timeout');
  }
  return setTimeoutImpl(callback, timeoutMs);
}

export function probeCanonicalResourceBroker(
  port: ReadinessPort,
  identity: EditorSwReadinessIdentity,
  options: {
    requestId?: string;
    timeoutMs?: number;
    setTimeout?: typeof globalThis.setTimeout;
    clearTimeout?: typeof globalThis.clearTimeout;
  } = {},
): Promise<EditorSwReadinessProbe> {
  const requestId = options.requestId ?? `prime-${crypto.randomUUID()}`;
  const setTimeoutImpl = options.setTimeout ?? globalThis.setTimeout;
  const clearTimeoutImpl = options.clearTimeout ?? globalThis.clearTimeout;
  const timeoutMs = options.timeoutMs ?? EDITOR_SW_READINESS_TIMEOUT_MS;

  return new Promise((resolve, reject) => {
    const timeout = readinessTimeout(
      () => {
        port.onmessage = null;
        port.onmessageerror = null;
        reject(new Error('Resource Broker readiness probe timed out'));
      },
      timeoutMs,
      setTimeoutImpl,
    );
    const finish = (callback: () => void) => {
      clearTimeoutImpl(timeout);
      port.onmessage = null;
      port.onmessageerror = null;
      callback();
    };
    port.onmessage = (event) => {
      const message = parseResourceBrokerServerMessage(event.data);
      if (
        !message ||
        message.id !== requestId ||
        (message.type !== 'RESULT' && message.type !== 'ERROR') ||
        (message.type === 'RESULT' && message.value.releaseId !== identity.releaseId)
      ) {
        finish(() => reject(new Error('Invalid Resource Broker readiness response')));
        return;
      }
      if (message.type === 'ERROR') {
        finish(() => reject(new Error(`Resource Broker readiness probe failed: ${message.code}`)));
        return;
      }
      finish(() => resolve(message.value));
    };
    port.onmessageerror = () => finish(() => reject(new Error('Resource Broker readiness probe channel failed')));
    port.start();
    port.postMessage({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'PROBE',
      id: requestId,
      releaseId: identity.releaseId,
      sessionId: identity.sessionId,
    });
  });
}

export function bindCanonicalBrokerToEditorServiceWorker(
  worker: ReadinessWorker,
  brokerPort: ReadinessPort,
  identity: EditorSwReadinessIdentity,
  options: {
    timeoutMs?: number;
    createMessageChannel?: () => MessageChannel;
    setTimeout?: typeof globalThis.setTimeout;
    clearTimeout?: typeof globalThis.clearTimeout;
  } = {},
): Promise<void> {
  const createMessageChannel = options.createMessageChannel ?? (() => new MessageChannel());
  const setTimeoutImpl = options.setTimeout ?? globalThis.setTimeout;
  const clearTimeoutImpl = options.clearTimeout ?? globalThis.clearTimeout;
  const timeoutMs = options.timeoutMs ?? EDITOR_SW_READINESS_TIMEOUT_MS;

  return new Promise((resolve, reject) => {
    const channel = createMessageChannel();
    const timeout = readinessTimeout(
      () => {
        channel.port1.close();
        reject(new Error('Editor Service Worker broker binding timed out'));
      },
      timeoutMs,
      setTimeoutImpl,
    );
    const finish = (callback: () => void) => {
      clearTimeoutImpl(timeout);
      channel.port1.onmessage = null;
      channel.port1.onmessageerror = null;
      channel.port1.close();
      callback();
    };
    channel.port1.onmessage = (event) => {
      const value = event.data;
      if (
        exactRecord(value, ['protocol', 'type', 'ok', 'releaseId', 'sessionId']) &&
        value.protocol === RESOURCE_BROKER_PROTOCOL &&
        value.type === EDITOR_RESOURCE_BROKER_BOUND_TYPE &&
        value.ok === true &&
        value.releaseId === identity.releaseId &&
        value.sessionId === identity.sessionId
      ) {
        finish(resolve);
        return;
      }
      if (
        exactRecord(value, ['protocol', 'type', 'ok', 'code']) &&
        value.protocol === RESOURCE_BROKER_PROTOCOL &&
        value.type === EDITOR_RESOURCE_BROKER_BIND_ERROR_TYPE &&
        value.ok === false
      ) {
        finish(() => reject(new Error(`Editor Service Worker broker binding failed: ${String(value.code)}`)));
        return;
      }
      finish(() => reject(new Error('Invalid Editor Service Worker broker binding response')));
    };
    channel.port1.onmessageerror = () =>
      finish(() => reject(new Error('Editor Service Worker broker binding response failed')));
    channel.port1.start();
    try {
      worker.postMessage(
        {
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: EDITOR_RESOURCE_BROKER_BIND_TYPE,
          releaseId: identity.releaseId,
          sessionId: identity.sessionId,
        },
        [brokerPort as unknown as Transferable, channel.port2],
      );
    } catch (error) {
      finish(() => reject(error));
    }
  });
}

export async function verifyEditorServiceWorkerRead(
  fetchImpl: typeof fetch,
  origin: string,
  probe: EditorSwReadinessProbe,
): Promise<void> {
  const url = new URL(`/r/${encodeURIComponent(probe.releaseId)}/${probe.probePath}`, origin);
  const response = await fetchImpl(url, {
    method: 'GET',
    headers: { Range: 'bytes=0-0' },
    cache: 'no-store',
    credentials: 'omit',
    redirect: 'error',
  });
  const expectedContentRange = `bytes 0-0/${probe.probeBytes}`;
  if (
    response.status !== 206 ||
    response.headers.get('accept-ranges')?.toLowerCase() !== 'bytes' ||
    response.headers.get('content-length') !== '1' ||
    response.headers.get('content-range') !== expectedContentRange
  ) {
    await response.body?.cancel().catch(() => undefined);
    throw new Error('Editor Service Worker readiness Range response is invalid');
  }
  const bytes = await response.arrayBuffer();
  if (bytes.byteLength !== 1) throw new Error('Editor Service Worker readiness Range body is invalid');
}
