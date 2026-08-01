import { OFFICE_EDITOR_ORIGIN_SLOTS, type OfficeEditorOriginSlot } from './office-origin-pool';
import {
  EDITOR_SHELL_MAX_ASSETS,
  EDITOR_SHELL_MAX_TOTAL_BYTES,
  EDITOR_SHELL_PRIME_INSTALL_QUERY,
  EDITOR_SHELL_PRIME_PATH,
  isEditorShellAssetPath,
} from './editor-shell-cache';
import {
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerReleaseId,
  isResourceBrokerSessionId,
} from './resource-broker-protocol';

export const EDITOR_ORIGIN_PRIME_TIMEOUT_MS = 30_000;
export const EDITOR_ORIGIN_PRIME_CONCURRENCY = 3;

export type EditorOriginPrimeResult = {
  origin: string;
  releaseId: string;
  sessionId: string;
  brokerReady: boolean;
  occupied: boolean;
  serviceWorkerVersion: string;
  cachedPaths: string[];
  cachedBytes: number;
};

export type EditorOriginPrimeProgress = (result: EditorOriginPrimeResult) => void;

type EditorOriginPrimeMessage = EditorOriginPrimeResult & {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'ONLYOFFICE_EDITOR_SHELL_PRIMED';
};

type EditorOriginPrimeFailureMessage = {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_FAILED';
  origin: string;
  releaseId: string;
  sessionId: string;
  code: 'identity' | 'storage' | 'timeout' | 'cancelled';
  stage: 'service-worker' | 'shell-cache' | 'shell-route' | 'broker-probe' | 'unknown';
  detail?: string;
};

type EditorOriginPrimeProgressMessage = {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_PROGRESS';
  origin: string;
  releaseId: string;
  sessionId: string;
  stage: 'service-worker' | 'shell-cache' | 'shell-route' | 'broker-probe';
};

export class EditorOriginPrimeError extends Error {
  readonly origin: string;
  readonly code: EditorOriginPrimeFailureMessage['code'];
  readonly stage: EditorOriginPrimeFailureMessage['stage'];
  readonly detail?: string;

  constructor(message: EditorOriginPrimeFailureMessage) {
    super(
      `Editor origin prime failed (${message.stage}/${message.code}${message.detail ? `: ${message.detail}` : ''})`,
    );
    this.name = 'EditorOriginPrimeError';
    this.origin = message.origin;
    this.code = message.code;
    this.stage = message.stage;
    this.detail = message.detail;
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function parsePrimeResult(
  value: unknown,
): EditorOriginPrimeMessage | EditorOriginPrimeFailureMessage | EditorOriginPrimeProgressMessage | null {
  if (
    isRecord(value) &&
    Object.keys(value).length === 6 &&
    value.protocol === RESOURCE_BROKER_PROTOCOL &&
    value.type === 'ONLYOFFICE_EDITOR_SHELL_PRIME_PROGRESS' &&
    typeof value.origin === 'string' &&
    isResourceBrokerReleaseId(value.releaseId) &&
    isResourceBrokerSessionId(value.sessionId) &&
    ['service-worker', 'shell-cache', 'shell-route', 'broker-probe'].includes(String(value.stage))
  ) {
    return value as EditorOriginPrimeProgressMessage;
  }
  if (
    isRecord(value) &&
    (Object.keys(value).length === 7 || Object.keys(value).length === 8) &&
    value.protocol === RESOURCE_BROKER_PROTOCOL &&
    value.type === 'ONLYOFFICE_EDITOR_SHELL_PRIME_FAILED' &&
    typeof value.origin === 'string' &&
    isResourceBrokerReleaseId(value.releaseId) &&
    isResourceBrokerSessionId(value.sessionId) &&
    ['identity', 'storage', 'timeout', 'cancelled'].includes(String(value.code)) &&
    ['service-worker', 'shell-cache', 'shell-route', 'broker-probe', 'unknown'].includes(String(value.stage)) &&
    (value.detail === undefined || (typeof value.detail === 'string' && value.detail.length <= 200))
  ) {
    return value as EditorOriginPrimeFailureMessage;
  }
  if (
    !isRecord(value) ||
    Object.keys(value).length !== 10 ||
    value.protocol !== RESOURCE_BROKER_PROTOCOL ||
    value.type !== 'ONLYOFFICE_EDITOR_SHELL_PRIMED' ||
    typeof value.origin !== 'string' ||
    !isResourceBrokerReleaseId(value.releaseId) ||
    !isResourceBrokerSessionId(value.sessionId) ||
    typeof value.brokerReady !== 'boolean' ||
    typeof value.occupied !== 'boolean' ||
    value.brokerReady === value.occupied ||
    typeof value.serviceWorkerVersion !== 'string' ||
    value.serviceWorkerVersion.length === 0 ||
    !Array.isArray(value.cachedPaths) ||
    value.cachedPaths.length < 3 ||
    value.cachedPaths.length > EDITOR_SHELL_MAX_ASSETS ||
    new Set(value.cachedPaths).size !== value.cachedPaths.length ||
    !value.cachedPaths.every((path) => typeof path === 'string' && isEditorShellAssetPath(path)) ||
    !Number.isSafeInteger(value.cachedBytes) ||
    Number(value.cachedBytes) <= 0 ||
    Number(value.cachedBytes) > EDITOR_SHELL_MAX_TOTAL_BYTES
  ) {
    return null;
  }
  return value as EditorOriginPrimeMessage;
}
export function editorOriginForSlot(
  slot: OfficeEditorOriginSlot,
  canonicalOrigin: string,
  localMatrix = false,
): string {
  const canonical = new URL(canonicalOrigin);
  if (!localMatrix) return `https://${slot}.getpi.work`;
  if (
    canonical.hostname !== 'onlyoffice.localhost' ||
    (canonical.protocol !== 'http:' && canonical.protocol !== 'https:')
  ) {
    throw new TypeError('Invalid local canonical editor-prime origin');
  }
  return `${canonical.protocol}//host-${slot}.office.localhost${canonical.port ? `:${canonical.port}` : ''}`;
}

export class EditorOriginPrimer {
  readonly #window: Window;
  readonly #document: Document;
  readonly #canonicalOrigin: string;
  readonly #localMatrix: boolean;
  readonly #timeoutMs: number;
  readonly #randomUUID: () => string;

  constructor(
    options: {
      canonicalOrigin?: string;
      localMatrix?: boolean;
      timeoutMs?: number;
      randomUUID?: () => string;
      window?: Window;
      document?: Document;
    } = {},
  ) {
    this.#window = options.window ?? window;
    this.#document = options.document ?? document;
    const canonical = new URL(options.canonicalOrigin ?? this.#window.location.origin);
    this.#canonicalOrigin = canonical.origin;
    this.#localMatrix = options.localMatrix ?? canonical.hostname === 'onlyoffice.localhost';
    this.#timeoutMs = options.timeoutMs ?? EDITOR_ORIGIN_PRIME_TIMEOUT_MS;
    this.#randomUUID = options.randomUUID ?? (() => crypto.randomUUID());
    if (
      !Number.isSafeInteger(this.#timeoutMs) ||
      this.#timeoutMs <= 0 ||
      this.#timeoutMs > EDITOR_ORIGIN_PRIME_TIMEOUT_MS
    ) {
      throw new TypeError('Invalid editor origin prime timeout');
    }
  }

  async primeRelease(
    releaseId: string,
    slots: readonly OfficeEditorOriginSlot[] = OFFICE_EDITOR_ORIGIN_SLOTS,
    onProgress?: EditorOriginPrimeProgress,
  ): Promise<EditorOriginPrimeResult[]> {
    return this.#runRelease(releaseId, slots, 'install', onProgress);
  }

  async verifyRelease(
    releaseId: string,
    slots: readonly OfficeEditorOriginSlot[] = OFFICE_EDITOR_ORIGIN_SLOTS,
    onProgress?: EditorOriginPrimeProgress,
  ): Promise<EditorOriginPrimeResult[]> {
    return this.#runRelease(releaseId, slots, 'verify', onProgress);
  }

  async #runRelease(
    releaseId: string,
    slots: readonly OfficeEditorOriginSlot[],
    mode: 'install' | 'verify',
    onProgress?: EditorOriginPrimeProgress,
  ): Promise<EditorOriginPrimeResult[]> {
    if (!isResourceBrokerReleaseId(releaseId)) throw new TypeError('Invalid editor shell release');
    const uniqueSlots = [...new Set(slots)];
    const results: EditorOriginPrimeResult[] = [];
    let cursor = 0;
    const worker = async () => {
      while (cursor < uniqueSlots.length) {
        const slot = uniqueSlots[cursor++];
        const result = await this.#primeOrigin(
          releaseId,
          editorOriginForSlot(slot, this.#canonicalOrigin, this.#localMatrix),
          mode,
        );
        results.push(result);
        onProgress?.(result);
      }
    };
    await Promise.all(
      Array.from({ length: Math.min(EDITOR_ORIGIN_PRIME_CONCURRENCY, uniqueSlots.length) }, () => worker()),
    );
    return results.sort((left, right) => left.origin.localeCompare(right.origin));
  }

  async #primeOrigin(releaseId: string, origin: string, mode: 'install' | 'verify'): Promise<EditorOriginPrimeResult> {
    const sessionId = `prime-${this.#randomUUID()}`;
    if (!isResourceBrokerSessionId(sessionId)) throw new TypeError('Invalid editor origin prime session');
    const frame = this.#document.createElement('iframe');
    frame.hidden = true;
    frame.tabIndex = -1;
    frame.setAttribute('aria-hidden', 'true');
    frame.setAttribute('sandbox', 'allow-same-origin allow-scripts');
    frame.referrerPolicy = 'strict-origin';
    const url = new URL(`/r/${encodeURIComponent(releaseId)}/${EDITOR_SHELL_PRIME_PATH}`, origin);
    if (mode === 'install') url.searchParams.set(EDITOR_SHELL_PRIME_INSTALL_QUERY, '1');
    const fragment = new URLSearchParams({
      parentOrigin: this.#canonicalOrigin,
      releaseId,
      sessionId,
      mode,
    });
    url.hash = fragment.toString();
    frame.src = url.href;

    return new Promise<EditorOriginPrimeResult>((resolve, reject) => {
      let settled = false;
      let lastStage: EditorOriginPrimeFailureMessage['stage'] = 'unknown';
      const finish = (callback: () => void) => {
        if (settled) return;
        settled = true;
        this.#window.clearTimeout(timeout);
        this.#window.removeEventListener('message', onMessage);
        frame.remove();
        callback();
      };
      const timeout = this.#window.setTimeout(
        () =>
          finish(() =>
            reject(
              new EditorOriginPrimeError({
                protocol: RESOURCE_BROKER_PROTOCOL,
                type: 'ONLYOFFICE_EDITOR_SHELL_PRIME_FAILED',
                origin,
                releaseId,
                sessionId,
                code: 'timeout',
                stage: lastStage,
              }),
            ),
          ),
        this.#timeoutMs,
      );
      const onMessage = (event: MessageEvent) => {
        if (event.source !== frame.contentWindow || event.origin !== origin) return;
        const message = parsePrimeResult(event.data);
        if (
          !message ||
          message.releaseId !== releaseId ||
          message.sessionId !== sessionId ||
          message.origin !== origin
        ) {
          finish(() => reject(new Error('Invalid editor origin prime response')));
          return;
        }
        if (message.type === 'ONLYOFFICE_EDITOR_SHELL_PRIME_PROGRESS') {
          lastStage = message.stage;
          return;
        }
        if (message.type === 'ONLYOFFICE_EDITOR_SHELL_PRIME_FAILED') {
          finish(() => reject(new EditorOriginPrimeError(message)));
          return;
        }
        finish(() =>
          resolve({
            origin: message.origin,
            releaseId: message.releaseId,
            sessionId: message.sessionId,
            brokerReady: message.brokerReady,
            occupied: message.occupied,
            serviceWorkerVersion: message.serviceWorkerVersion,
            cachedPaths: [...message.cachedPaths],
            cachedBytes: message.cachedBytes,
          }),
        );
      };
      this.#window.addEventListener('message', onMessage);
      this.#document.body.append(frame);
    });
  }
}
