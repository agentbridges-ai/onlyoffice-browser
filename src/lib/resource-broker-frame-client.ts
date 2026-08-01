import {
  RESOURCE_BROKER_MAX_CAPABILITY_TTL_MS,
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
  parseResourceBrokerChallengeMessage,
  type ResourceBrokerCapabilityClaim,
  type ResourceBrokerChallengeMessage,
  type ResourceBrokerEditorOrigin,
} from './resource-broker-protocol';
import { isOfficeEditorOriginSlot } from './office-origin-pool';

export const RESOURCE_BROKER_CANONICAL_ORIGIN = 'https://onlyoffice.getpi.work';
export const RESOURCE_BROKER_ALLOWED_PARENT_ORIGINS = [
  'https://piwork.getpi.work',
  'https://onlyoffice.getpi.work',
] as const;
export const RESOURCE_BROKER_FRAME_PATH = '/resource-broker.html';
export const RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS = 30_000;

const LOCAL_CANONICAL_HOSTNAME = 'onlyoffice.localhost';
const LOCAL_PARENT_HOSTNAMES = new Map([
  ['piwork.localhost', 'https://piwork.getpi.work'],
  ['onlyoffice.localhost', 'https://onlyoffice.getpi.work'],
]);

export interface ResourceBrokerPhysicalEditorIdentity {
  physicalOrigin: string;
  logicalOrigin: ResourceBrokerEditorOrigin;
  localMatrix: boolean;
}

export interface ResourceBrokerTrustedParentIdentity {
  physicalOrigin: string;
  logicalOrigin: (typeof RESOURCE_BROKER_ALLOWED_PARENT_ORIGINS)[number];
  localMatrix: boolean;
}

export interface ResourceBrokerFrameExpectedIdentity {
  canonicalOrigin: string;
  parentOrigin: (typeof RESOURCE_BROKER_ALLOWED_PARENT_ORIGINS)[number];
  editorOrigin: ResourceBrokerEditorOrigin;
  releaseId: string;
  sessionId: string;
}

export interface ResourceBrokerFrameConnection {
  identity: ResourceBrokerFrameExpectedIdentity;
  port: MessagePort;
}

export interface ResourceBrokerFrameClientOptions {
  releaseId: string;
  sessionId: string;
  parentOrigin: string;
  canonicalOrigin?: string;
  allowLocalMatrix?: boolean;
  connectTimeoutMs?: number;
  window?: Window;
  document?: Document;
}

type ResourceBrokerFrameConnectedMessage = {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'CONNECTED';
  releaseId: string;
  sessionId: string;
};

type ResourceBrokerFrameConnectErrorMessage = {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'CONNECT_ERROR';
  code: 'capability' | 'identity' | 'worker' | 'timeout';
};

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function hasExactKeys(value: Record<string, unknown>, expected: readonly string[]): boolean {
  const keys = Object.keys(value);
  return keys.length === expected.length && expected.every((key) => Object.hasOwn(value, key));
}

function parseExactOrigin(value: string): URL | null {
  try {
    const url = new URL(value);
    if (url.username || url.password || url.pathname !== '/' || url.search || url.hash || value !== url.origin) {
      return null;
    }
    return url;
  } catch {
    return null;
  }
}

function localOriginMatchesCanonical(origin: URL, canonical: URL): boolean {
  return origin.protocol === canonical.protocol && origin.port === canonical.port;
}

export function resolveResourceBrokerCanonicalOrigin(
  value: string,
  allowLocalMatrix = false,
): { origin: string; localMatrix: boolean } | null {
  const url = parseExactOrigin(value);
  if (!url) return null;
  if (url.origin === RESOURCE_BROKER_CANONICAL_ORIGIN) {
    return { origin: url.origin, localMatrix: false };
  }
  if (
    allowLocalMatrix &&
    (url.protocol === 'http:' || url.protocol === 'https:') &&
    url.hostname === LOCAL_CANONICAL_HOSTNAME
  ) {
    return { origin: url.origin, localMatrix: true };
  }
  return null;
}

export function resolveResourceBrokerPhysicalEditorIdentity(
  value: string,
  canonicalOrigin: string,
  allowLocalMatrix = false,
): ResourceBrokerPhysicalEditorIdentity | null {
  const origin = parseExactOrigin(value);
  const canonical = resolveResourceBrokerCanonicalOrigin(canonicalOrigin, allowLocalMatrix);
  if (!origin || !canonical) return null;

  if (!canonical.localMatrix) {
    const hostnameMatch = /^(?<slot>[^.]+)\.getpi\.work$/.exec(origin.hostname);
    const slot = hostnameMatch?.groups?.slot;
    if (
      origin.protocol !== 'https:' ||
      origin.port ||
      !slot ||
      !isOfficeEditorOriginSlot(slot) ||
      origin.origin !== `https://${slot}.getpi.work`
    ) {
      return null;
    }
    return {
      physicalOrigin: origin.origin,
      logicalOrigin: origin.origin as ResourceBrokerEditorOrigin,
      localMatrix: false,
    };
  }

  const canonicalUrl = parseExactOrigin(canonical.origin);
  if (!canonicalUrl || !localOriginMatchesCanonical(origin, canonicalUrl)) return null;
  const localMatch =
    /^(?<slot>[^.]+)\.localhost$/.exec(origin.hostname) ??
    /^host-(?<slot>[^.]+)\.office\.localhost$/.exec(origin.hostname);
  const slot = localMatch?.groups?.slot;
  if (!slot || !isOfficeEditorOriginSlot(slot)) return null;
  return {
    physicalOrigin: origin.origin,
    logicalOrigin: `https://${slot}.getpi.work`,
    localMatrix: true,
  };
}

export function resolveResourceBrokerTrustedParentIdentity(
  value: string,
  canonicalOrigin: string,
  allowLocalMatrix = false,
): ResourceBrokerTrustedParentIdentity | null {
  const origin = parseExactOrigin(value);
  const canonical = resolveResourceBrokerCanonicalOrigin(canonicalOrigin, allowLocalMatrix);
  if (!origin || !canonical) return null;

  if (!canonical.localMatrix) {
    if (!RESOURCE_BROKER_ALLOWED_PARENT_ORIGINS.includes(origin.origin as never)) return null;
    return {
      physicalOrigin: origin.origin,
      logicalOrigin: origin.origin as ResourceBrokerTrustedParentIdentity['logicalOrigin'],
      localMatrix: false,
    };
  }

  const canonicalUrl = parseExactOrigin(canonical.origin);
  const logicalOrigin = LOCAL_PARENT_HOSTNAMES.get(origin.hostname);
  if (!canonicalUrl || !logicalOrigin || !localOriginMatchesCanonical(origin, canonicalUrl)) return null;
  return {
    physicalOrigin: origin.origin,
    logicalOrigin: logicalOrigin as ResourceBrokerTrustedParentIdentity['logicalOrigin'],
    localMatrix: true,
  };
}

export function parseResourceBrokerFrameConnectedMessage(
  value: unknown,
): ResourceBrokerFrameConnectedMessage | ResourceBrokerFrameConnectErrorMessage | null {
  if (!isRecord(value) || value.protocol !== RESOURCE_BROKER_PROTOCOL || typeof value.type !== 'string') {
    return null;
  }
  if (
    value.type === 'CONNECTED' &&
    hasExactKeys(value, ['protocol', 'type', 'releaseId', 'sessionId']) &&
    isResourceBrokerReleaseId(value.releaseId) &&
    isResourceBrokerSessionId(value.sessionId)
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CONNECTED',
      releaseId: value.releaseId,
      sessionId: value.sessionId,
    };
  }
  if (
    value.type === 'CONNECT_ERROR' &&
    hasExactKeys(value, ['protocol', 'type', 'code']) &&
    ['capability', 'identity', 'worker', 'timeout'].includes(String(value.code))
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CONNECT_ERROR',
      code: value.code as ResourceBrokerFrameConnectErrorMessage['code'],
    };
  }
  return null;
}

export function validateResourceBrokerFrameChallenge(
  value: unknown,
  expected: ResourceBrokerFrameExpectedIdentity,
  nowMs = Date.now(),
): ResourceBrokerChallengeMessage | null {
  const message = parseResourceBrokerChallengeMessage(value);
  if (
    !message ||
    message.capability.parentOrigin !== expected.parentOrigin ||
    message.capability.editorOrigin !== expected.editorOrigin ||
    message.capability.releaseId !== expected.releaseId ||
    message.capability.sessionId !== expected.sessionId ||
    message.capability.issuedAtMs > nowMs ||
    message.capability.expiresAtMs <= nowMs ||
    message.capability.expiresAtMs - nowMs > RESOURCE_BROKER_MAX_CAPABILITY_TTL_MS
  ) {
    return null;
  }
  return message;
}

export function validateResourceBrokerFrameChallengeEvent(
  event: Pick<MessageEvent, 'data' | 'origin' | 'source'>,
  expectedSource: MessageEventSource | null,
  expected: ResourceBrokerFrameExpectedIdentity,
  nowMs = Date.now(),
): ResourceBrokerChallengeMessage | null {
  if (event.source !== expectedSource || event.origin !== expected.canonicalOrigin) return null;
  return validateResourceBrokerFrameChallenge(event.data, expected, nowMs);
}

export function createResourceBrokerFrameUrl(identity: ResourceBrokerFrameExpectedIdentity, localMatrix = false): URL {
  const url = new URL(RESOURCE_BROKER_FRAME_PATH, identity.canonicalOrigin);
  url.searchParams.set('releaseId', identity.releaseId);
  url.searchParams.set('sessionId', identity.sessionId);
  url.searchParams.set('parentOrigin', identity.parentOrigin);
  if (localMatrix) url.searchParams.set('localMatrix', '1');
  return url;
}

export class ResourceBrokerFrameClient {
  readonly #options: ResourceBrokerFrameClientOptions;
  readonly #window: Window;
  readonly #document: Document;
  #frame: HTMLIFrameElement | null = null;
  #connectionPort: MessagePort | null = null;
  #connectPromise: Promise<ResourceBrokerFrameConnection> | null = null;
  #destroyed = false;
  #messageListener: ((event: MessageEvent) => void) | null = null;
  #rejectConnect: ((reason?: unknown) => void) | null = null;

  constructor(options: ResourceBrokerFrameClientOptions) {
    this.#options = options;
    this.#window = options.window ?? window;
    this.#document = options.document ?? document;
  }

  connect(): Promise<ResourceBrokerFrameConnection> {
    if (this.#destroyed) return Promise.reject(new Error('Resource Broker frame client was destroyed'));
    if (this.#connectPromise) return this.#connectPromise;
    this.#connectPromise = this.#connect();
    return this.#connectPromise;
  }

  async #connect(): Promise<ResourceBrokerFrameConnection> {
    const canonical = resolveResourceBrokerCanonicalOrigin(
      this.#options.canonicalOrigin ?? RESOURCE_BROKER_CANONICAL_ORIGIN,
      this.#options.allowLocalMatrix,
    );
    if (!canonical) throw new Error('Invalid Resource Broker canonical origin');
    const editor = resolveResourceBrokerPhysicalEditorIdentity(
      this.#window.location.origin,
      canonical.origin,
      this.#options.allowLocalMatrix,
    );
    const trustedParent = resolveResourceBrokerTrustedParentIdentity(
      this.#options.parentOrigin,
      canonical.origin,
      this.#options.allowLocalMatrix,
    );
    if (
      !editor ||
      !trustedParent ||
      !isResourceBrokerReleaseId(this.#options.releaseId) ||
      !isResourceBrokerSessionId(this.#options.sessionId)
    ) {
      throw new Error('Invalid Resource Broker frame identity');
    }

    const identity: ResourceBrokerFrameExpectedIdentity = {
      canonicalOrigin: canonical.origin,
      parentOrigin: trustedParent.logicalOrigin,
      editorOrigin: editor.logicalOrigin,
      releaseId: this.#options.releaseId,
      sessionId: this.#options.sessionId,
    };
    const frame = this.#document.createElement('iframe');
    frame.hidden = true;
    frame.tabIndex = -1;
    frame.setAttribute('aria-hidden', 'true');
    frame.setAttribute('sandbox', 'allow-same-origin allow-scripts');
    frame.referrerPolicy = 'strict-origin';
    frame.src = createResourceBrokerFrameUrl(identity, canonical.localMatrix).href;
    this.#frame = frame;

    const timeoutMs = this.#options.connectTimeoutMs ?? RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS;
    if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0 || timeoutMs > RESOURCE_BROKER_FRAME_CONNECT_TIMEOUT_MS) {
      throw new TypeError('Invalid Resource Broker frame timeout');
    }

    const result = new Promise<ResourceBrokerFrameConnection>((resolve, reject) => {
      this.#rejectConnect = reject;
      const timeout = this.#window.setTimeout(() => {
        reject(new Error('Resource Broker frame connection timed out'));
        this.destroy();
      }, timeoutMs);

      const settleReject = (error: Error) => {
        this.#window.clearTimeout(timeout);
        reject(error);
        this.destroy();
      };
      this.#messageListener = (event: MessageEvent) => {
        if (this.#destroyed || event.source !== frame.contentWindow || event.origin !== canonical.origin) {
          return;
        }
        const challenge = validateResourceBrokerFrameChallengeEvent(event, frame.contentWindow, identity);
        if (!challenge) {
          settleReject(new Error('Invalid Resource Broker challenge'));
          return;
        }

        const channel = new MessageChannel();
        const port = channel.port1;
        let connected = false;
        port.onmessage = (portEvent: MessageEvent) => {
          const message = parseResourceBrokerFrameConnectedMessage(portEvent.data);
          if (!message) {
            settleReject(new Error('Invalid Resource Broker connection response'));
            port.close();
            return;
          }
          if (message.type === 'CONNECT_ERROR') {
            settleReject(new Error(`Resource Broker connection failed: ${message.code}`));
            port.close();
            return;
          }
          if (message.releaseId !== identity.releaseId || message.sessionId !== identity.sessionId) {
            settleReject(new Error('Resource Broker connected with a mismatched identity'));
            port.close();
            return;
          }
          connected = true;
          this.#window.clearTimeout(timeout);
          port.onmessage = null;
          this.#connectionPort = port;
          this.#rejectConnect = null;
          if (this.#messageListener) {
            this.#window.removeEventListener('message', this.#messageListener);
            this.#messageListener = null;
          }
          resolve({ identity, port });
        };
        port.onmessageerror = () => {
          if (!connected) settleReject(new Error('Resource Broker connection channel failed'));
        };
        port.start();

        const claim: ResourceBrokerCapabilityClaim = {
          token: challenge.capability.token,
          parentOrigin: identity.parentOrigin,
          editorOrigin: identity.editorOrigin,
          releaseId: identity.releaseId,
          sessionId: identity.sessionId,
        };
        frame.contentWindow?.postMessage(
          {
            protocol: RESOURCE_BROKER_PROTOCOL,
            type: 'CONNECT',
            capability: claim,
          },
          canonical.origin,
          [channel.port2],
        );
      };
      this.#window.addEventListener('message', this.#messageListener);
    });

    this.#document.body.append(frame);
    return result;
  }

  /**
   * Transfers ownership of the connected port to the Editor Service Worker.
   * The canonical iframe remains alive because it owns the relay endpoint at
   * the other side of the transferred port.
   */
  takeConnectionPort(): MessagePort {
    if (this.#destroyed || !this.#connectionPort) {
      throw new Error('Resource Broker connection is not available for transfer');
    }
    const port = this.#connectionPort;
    this.#connectionPort = null;
    return port;
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    if (this.#messageListener) {
      this.#window.removeEventListener('message', this.#messageListener);
      this.#messageListener = null;
    }
    this.#rejectConnect?.(new Error('Resource Broker frame client was destroyed'));
    this.#rejectConnect = null;
    this.#connectionPort?.close();
    this.#connectionPort = null;
    this.#frame?.remove();
    this.#frame = null;
  }
}

export function createResourceBrokerFrameClient(options: ResourceBrokerFrameClientOptions): ResourceBrokerFrameClient {
  return new ResourceBrokerFrameClient(options);
}
