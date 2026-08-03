import {
  parseRequiredReleaseIdentity,
  requiredReleaseIdentitiesEqual,
  type RequiredReleaseIdentity,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
  type ResourcePlanRequest,
} from './lib/release-resources';
import { CanonicalResourceInstaller } from './lib/canonical-resource-installer';
import { CanonicalResourceReadinessGate } from './lib/canonical-resource-readiness-gate';
import {
  RESOURCE_INSTALLER_FRAME_PROTOCOL,
  RESOURCE_INSTALLER_RPC_TIMEOUT_MS,
  ResourceInstallerCapabilityGate,
  isResourceInstallerFrameCommand,
  isResourceInstallerPlan,
  isResourceInstallerRequestId,
  isResourceInstallerSnapshot,
  parseResourceInstallerInitMessage,
  parseResourceInstallerReconnectMessage,
  parseResourceInstallerRequestMessage,
  resolveResourceInstallerFrameIdentity,
  resourceInstallerErrorPayload,
  type ResourceInstallerErrorMessage,
  type ResourceInstallerFrameCommand,
  type ResourceInstallerFrameIdentity,
  type ResourceInstallerRequestMessage,
  type ResourceInstallerServerMessage,
} from './lib/resource-installer-frame-protocol';

export interface ResourceInstallerFrameManager {
  getInstallerSnapshot(): ResourceInstallerSnapshot;
  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void;
  plan(request: ResourcePlanRequest): Promise<ResourcePlan>;
  apply(plan: ResourcePlan): Promise<void>;
  checkForUpdates(): Promise<void>;
  checkHealth(options?: { deep?: boolean }): Promise<void>;
  repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void>;
  pause(): void;
  resume(): Promise<void>;
  cancel(): void;
}

export type CreateResourceInstallerFrameManager = (
  requiredReleaseIdentity: RequiredReleaseIdentity,
  reportProgress?: () => void,
) => ResourceInstallerFrameManager | Promise<ResourceInstallerFrameManager>;

export interface ResourceInstallerFramePort {
  onmessage: ((event: MessageEvent) => void) | null;
  onmessageerror: ((event: MessageEvent) => void) | null;
  close(): void;
  postMessage(message: unknown): void;
  start(): void;
}

export interface StartResourceInstallerFrameOptions {
  createManager: CreateResourceInstallerFrameManager;
  window?: Window;
  document?: Document;
  now?: () => number;
  randomFill?: (bytes: Uint8Array<ArrayBuffer>) => Uint8Array<ArrayBuffer>;
  requestTimeoutMs?: number;
}

export type CreateCanonicalResourceInstallerFrameManagerOptions = {
  requiredReleaseIdentity: RequiredReleaseIdentity;
  window?: Window;
  document?: Document;
  fetch?: typeof fetch;
  cacheStorage?: CacheStorage;
  indexedDb?: IDBFactory;
  locks?: LockManager;
  broadcast?: BroadcastChannel;
  reportProgress?: () => void;
};

type ActiveRequest = {
  command: ResourceInstallerFrameCommand;
  running: boolean;
  settled: boolean;
  timeout: ReturnType<typeof globalThis.setTimeout> | null;
};

const MUTATION_COMMANDS = new Set<ResourceInstallerFrameCommand>(['APPLY', 'CHECK_UPDATES', 'CHECK_HEALTH', 'REPAIR']);
const ACTIVITY_COMMANDS = new Set<ResourceInstallerFrameCommand>([
  'APPLY',
  'CHECK_UPDATES',
  'CHECK_HEALTH',
  'REPAIR',
  'RESUME',
]);
const PAUSABLE_COMMANDS = new Set<ResourceInstallerFrameCommand>(['APPLY', 'REPAIR']);

function validTimeout(value: number | undefined): number {
  const timeout = value ?? RESOURCE_INSTALLER_RPC_TIMEOUT_MS;
  if (!Number.isSafeInteger(timeout) || timeout <= 0 || timeout > RESOURCE_INSTALLER_RPC_TIMEOUT_MS) {
    throw new TypeError('Invalid Resource Installer frame timeout');
  }
  return timeout;
}

function initError(id: string, error: ResourceInstallerErrorMessage['error']): ResourceInstallerErrorMessage {
  return {
    protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
    type: 'ERROR',
    id,
    command: 'INIT',
    error,
  };
}

export class ResourceInstallerFrameSession {
  readonly #manager: ResourceInstallerFrameManager;
  readonly #port: ResourceInstallerFramePort;
  readonly #requestTimeoutMs: number;
  readonly #onDisconnect: (() => void) | null;
  readonly #active = new Map<string, ActiveRequest>();
  #mutationQueue: Promise<unknown> = Promise.resolve();
  #unsubscribe: (() => void) | null = null;
  #lastSnapshot: ResourceInstallerSnapshot | null = null;
  #destroyed = false;

  constructor(options: {
    manager: ResourceInstallerFrameManager;
    port: ResourceInstallerFramePort;
    requestTimeoutMs?: number;
    onDisconnect?: () => void;
  }) {
    this.#manager = options.manager;
    this.#port = options.port;
    this.#requestTimeoutMs = validTimeout(options.requestTimeoutMs);
    this.#onDisconnect = options.onDisconnect ?? null;
  }

  start(initId: string, subscribe: boolean): void {
    if (this.#destroyed || !isResourceInstallerRequestId(initId)) {
      throw new Error('Invalid Resource Installer frame session');
    }
    this.#port.onmessage = (event) => this.handle(event.data);
    this.#port.onmessageerror = () => this.disconnect();
    this.#port.start();
    if (subscribe) {
      this.#unsubscribe = this.#manager.subscribeInstaller((next) => {
        if (this.#destroyed || !isResourceInstallerSnapshot(next)) return;
        this.#lastSnapshot = next;
        this.#refreshActivity(next);
        this.#post({
          protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
          type: 'EVENT',
          event: 'SNAPSHOT',
          snapshot: next,
        });
      });
    }
    const snapshot = this.#snapshot();
    this.#lastSnapshot = snapshot;
    if (
      !this.#post({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'READY',
        id: initId,
        command: 'INIT',
        snapshot,
      })
    ) {
      throw new Error('Resource Installer parent port is unavailable');
    }
  }

  handle(value: unknown): boolean {
    if (this.#destroyed) return false;
    const request = parseResourceInstallerRequestMessage(value);
    if (!request) {
      if (
        typeof value === 'object' &&
        value !== null &&
        'id' in value &&
        isResourceInstallerRequestId(value.id) &&
        'command' in value &&
        isResourceInstallerFrameCommand(value.command)
      ) {
        this.#postError(value.id, value.command, { code: 'protocol', retryable: false });
      } else {
        this.destroy();
      }
      return false;
    }
    if (this.#active.has(request.id)) {
      const original = this.#active.get(request.id)!;
      original.settled = true;
      if (original.timeout !== null) globalThis.clearTimeout(original.timeout);
      this.#active.delete(request.id);
      this.#postError(request.id, request.command, { code: 'protocol', retryable: false });
      return false;
    }

    const active: ActiveRequest = {
      command: request.command,
      running: false,
      settled: false,
      timeout: null,
    };
    this.#active.set(request.id, active);
    this.#armTimeout(request.id, active);

    const execute = (): Promise<ResourceInstallerServerMessage | null> => {
      if (active.settled || this.#destroyed) return Promise.resolve(null);
      active.running = true;
      this.#armTimeout(request.id, active);
      return this.#execute(request);
    };
    const operation: Promise<ResourceInstallerServerMessage | null> = MUTATION_COMMANDS.has(request.command)
      ? this.#enqueueMutation(execute)
      : execute();

    void operation.then(
      (message) => {
        if (!message || active.settled || this.#destroyed) return;
        active.settled = true;
        if (active.timeout !== null) globalThis.clearTimeout(active.timeout);
        this.#active.delete(request.id);
        this.#post(message);
      },
      (error) => {
        if (active.settled || this.#destroyed) return;
        active.settled = true;
        if (active.timeout !== null) globalThis.clearTimeout(active.timeout);
        this.#active.delete(request.id);
        this.#postError(request.id, request.command, resourceInstallerErrorPayload(error));
      },
    );
    return true;
  }

  disconnect(): void {
    this.#destroy(this.#onDisconnect === null);
    this.#onDisconnect?.();
  }

  destroy(): void {
    this.#destroy(true);
  }

  #destroy(cancelManager: boolean): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    this.#unsubscribe?.();
    this.#unsubscribe = null;
    for (const request of this.#active.values()) {
      if (request.timeout !== null) globalThis.clearTimeout(request.timeout);
    }
    this.#active.clear();
    this.#port.onmessage = null;
    this.#port.onmessageerror = null;
    this.#port.close();
    if (cancelManager) {
      try {
        this.#manager.cancel();
      } catch {
        // Page teardown is best-effort; no response can be delivered after the port closes.
      }
    }
  }

  async #execute(request: ResourceInstallerRequestMessage): Promise<ResourceInstallerServerMessage> {
    if (request.command === 'PLAN') {
      const plan = await this.#manager.plan(request.request);
      if (!isResourceInstallerPlan(plan)) throw { code: 'protocol' };
      return {
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'RESULT',
        id: request.id,
        command: 'PLAN',
        plan,
      };
    }
    if (request.command === 'APPLY') await this.#manager.apply(request.plan);
    else if (request.command === 'CHECK_UPDATES') await this.#manager.checkForUpdates();
    else if (request.command === 'CHECK_HEALTH') await this.#manager.checkHealth({ deep: false });
    else if (request.command === 'REPAIR') await this.#manager.repair({ scope: request.scope });
    else if (request.command === 'PAUSE') this.#manager.pause();
    else if (request.command === 'RESUME') await this.#manager.resume();
    else if (request.command === 'CANCEL') this.#manager.cancel();

    return {
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'RESULT',
      id: request.id,
      command: request.command,
      snapshot: this.#snapshot(),
    };
  }

  #snapshot(): ResourceInstallerSnapshot {
    const snapshot = this.#manager.getInstallerSnapshot();
    if (!isResourceInstallerSnapshot(snapshot)) {
      throw { code: 'protocol' };
    }
    return snapshot;
  }

  #enqueueMutation<T>(task: () => Promise<T>): Promise<T> {
    const queued = this.#mutationQueue.catch(() => undefined).then(task);
    this.#mutationQueue = queued;
    return queued;
  }

  #refreshActivity(snapshot: ResourceInstallerSnapshot): void {
    const paused = snapshot.phase === 'paused' || snapshot.readiness === 'paused';
    for (const [id, active] of this.#active) {
      if (!active.running || !ACTIVITY_COMMANDS.has(active.command)) continue;
      if (paused && PAUSABLE_COMMANDS.has(active.command)) {
        if (active.timeout !== null) globalThis.clearTimeout(active.timeout);
        active.timeout = null;
        continue;
      }
      this.#armTimeout(id, active);
    }
  }

  #armTimeout(id: string, active: ActiveRequest): void {
    if (active.settled || this.#destroyed) return;
    if (active.timeout !== null) globalThis.clearTimeout(active.timeout);
    if (
      active.running &&
      PAUSABLE_COMMANDS.has(active.command) &&
      (this.#lastSnapshot?.phase === 'paused' || this.#lastSnapshot?.readiness === 'paused')
    ) {
      active.timeout = null;
      return;
    }
    active.timeout = globalThis.setTimeout(() => {
      if (active.settled || this.#destroyed) return;
      active.settled = true;
      active.timeout = null;
      this.#active.delete(id);
      if (active.running && (MUTATION_COMMANDS.has(active.command) || active.command === 'RESUME')) {
        try {
          this.#manager.cancel();
        } catch {
          // The structured timeout remains authoritative even if cancellation is already complete.
        }
      }
      this.#postError(id, active.command, { code: 'timeout', retryable: true });
    }, this.#requestTimeoutMs);
  }

  #post(message: ResourceInstallerServerMessage): boolean {
    if (this.#destroyed) return false;
    try {
      this.#port.postMessage(message);
      return true;
    } catch {
      this.disconnect();
      return false;
    }
  }

  #postError(id: string, command: ResourceInstallerFrameCommand, error: ResourceInstallerErrorMessage['error']): void {
    if (this.#destroyed || !isResourceInstallerRequestId(id)) return;
    this.#post({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'ERROR',
      id,
      command,
      error,
    });
  }
}

export function startResourceInstallerFrame(options: StartResourceInstallerFrameOptions): () => void {
  const browserWindow = options.window ?? window;
  const browserDocument = options.document ?? document;
  const requestTimeoutMs = validTimeout(options.requestTimeoutMs);
  const identity = resolveResourceInstallerFrameIdentity({
    locationOrigin: browserWindow.location.origin,
    documentReferrer: browserDocument.referrer,
    search: browserWindow.location.search,
  });
  if (!identity || browserWindow.parent === browserWindow) {
    throw new Error('Resource Installer frame identity is not trusted');
  }

  let destroyed = false;
  let connectionStarted = false;
  let session: ResourceInstallerFrameSession | null = null;
  let manager: ResourceInstallerFrameManager | null = null;
  let managerPromise: Promise<ResourceInstallerFrameManager> | null = null;
  let boundReleaseIdentity: RequiredReleaseIdentity | null = null;
  let gate: ResourceInstallerCapabilityGate | null = null;
  let pendingPort: MessagePort | null = null;
  let pendingInitTimeout: number | null = null;
  let pendingInitHeartbeat: (() => void) | null = null;
  let connectionGeneration = 0;

  const issueChallenge = () => {
    if (destroyed || connectionStarted) return;
    gate = new ResourceInstallerCapabilityGate({
      parentOrigin: identity.physicalParentOrigin,
      now: options.now,
      randomFill: options.randomFill,
    });
    const capability = gate.issue();
    browserWindow.parent.postMessage(
      {
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'CHALLENGE',
        capability,
      },
      identity.physicalParentOrigin,
    );
  };

  const getManager = (requiredReleaseIdentity: RequiredReleaseIdentity): Promise<ResourceInstallerFrameManager> => {
    if (boundReleaseIdentity && !requiredReleaseIdentitiesEqual(boundReleaseIdentity, requiredReleaseIdentity)) {
      return Promise.reject({ code: 'incompatible', path: 'release/identity' });
    }
    if (!boundReleaseIdentity) boundReleaseIdentity = { ...requiredReleaseIdentity };
    if (manager) return Promise.resolve(manager);
    if (managerPromise) return managerPromise;
    managerPromise = Promise.resolve()
      .then(() => options.createManager({ ...requiredReleaseIdentity }, () => pendingInitHeartbeat?.()))
      .then(
        (created) => {
          if (destroyed) {
            try {
              created.cancel();
            } catch {
              // A manager that completed after page teardown cannot be retained.
            }
            throw { code: 'unavailable' };
          }
          manager = created;
          return created;
        },
        (error) => {
          managerPromise = null;
          throw error;
        },
      );
    return managerPromise;
  };

  const onWindowMessage = (event: MessageEvent) => {
    if (destroyed || event.source !== browserWindow.parent || event.origin !== identity.physicalParentOrigin) {
      return;
    }
    if (parseResourceInstallerReconnectMessage(event.data)) {
      if (event.ports.length !== 0) {
        for (const port of event.ports) port.close();
        return;
      }
      connectionGeneration += 1;
      if (pendingInitTimeout !== null) browserWindow.clearTimeout(pendingInitTimeout);
      pendingInitTimeout = null;
      pendingInitHeartbeat = null;
      pendingPort?.close();
      pendingPort = null;
      connectionStarted = false;
      if (session) {
        session.disconnect();
      } else {
        issueChallenge();
      }
      return;
    }
    const port = event.ports[0];
    if (!port || event.ports.length !== 1 || connectionStarted || !gate) {
      port?.close();
      return;
    }
    connectionStarted = true;
    pendingPort = port;
    const parsedInit = parseResourceInstallerInitMessage(event.data);
    const consumed = gate.consume(event.data, event.origin);
    gate = null;
    if (!parsedInit || !consumed.ok) {
      if (parsedInit) port.postMessage(initError(parsedInit.id, { code: 'capability', retryable: false }));
      port.close();
      pendingPort = null;
      return;
    }
    const init = consumed.message;

    const generation = ++connectionGeneration;
    let timedOut = false;
    const armPendingInitTimeout = () => {
      if (pendingInitTimeout !== null) browserWindow.clearTimeout(pendingInitTimeout);
      pendingInitTimeout = browserWindow.setTimeout(() => {
        if (destroyed || generation !== connectionGeneration) return;
        timedOut = true;
        connectionGeneration += 1;
        pendingInitTimeout = null;
        pendingInitHeartbeat = null;
        try {
          port.postMessage(initError(init.id, { code: 'timeout', retryable: true }));
        } catch {
          // The parent may have gone away while the manager was being created.
        }
        port.close();
        pendingPort = null;
        connectionStarted = false;
        issueChallenge();
      }, requestTimeoutMs);
    };
    pendingInitHeartbeat = () => {
      if (destroyed || timedOut || generation !== connectionGeneration) return;
      armPendingInitTimeout();
    };
    armPendingInitTimeout();
    void getManager(init.requiredReleaseIdentity).then(
      (connectedManager) => {
        if (destroyed || timedOut || generation !== connectionGeneration) return;
        if (pendingInitTimeout !== null) browserWindow.clearTimeout(pendingInitTimeout);
        pendingInitTimeout = null;
        pendingInitHeartbeat = null;
        pendingPort = null;
        const connectedSession = new ResourceInstallerFrameSession({
          manager: connectedManager,
          port,
          requestTimeoutMs,
          onDisconnect: () => {
            if (session !== connectedSession || destroyed) return;
            session = null;
            connectionStarted = false;
            issueChallenge();
          },
        });
        session = connectedSession;
        try {
          connectedSession.start(init.id, init.subscribe);
        } catch (error) {
          try {
            port.postMessage(initError(init.id, resourceInstallerErrorPayload(error)));
          } catch {
            // The connection failure is reported through the next capability handshake.
          }
          if (session === connectedSession) connectedSession.disconnect();
        }
      },
      (error) => {
        if (destroyed || timedOut || generation !== connectionGeneration) return;
        if (pendingInitTimeout !== null) browserWindow.clearTimeout(pendingInitTimeout);
        pendingInitTimeout = null;
        pendingInitHeartbeat = null;
        pendingPort = null;
        connectionStarted = false;
        try {
          port.postMessage(initError(init.id, resourceInstallerErrorPayload(error)));
        } catch {
          // The parent may have gone away while the manager was being created.
        }
        port.close();
        issueChallenge();
      },
    );
  };
  browserWindow.addEventListener('message', onWindowMessage);
  issueChallenge();

  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    browserWindow.removeEventListener('message', onWindowMessage);
    browserWindow.removeEventListener('pagehide', destroy);
    if (pendingInitTimeout !== null) browserWindow.clearTimeout(pendingInitTimeout);
    pendingInitTimeout = null;
    pendingInitHeartbeat = null;
    const hadSession = session !== null;
    session?.destroy();
    session = null;
    if (!hadSession && manager) {
      try {
        manager.cancel();
      } catch {
        // Page teardown cancellation is best-effort.
      }
    }
    manager = null;
    pendingPort?.close();
    pendingPort = null;
  };
  browserWindow.addEventListener('pagehide', destroy, { once: true });
  return destroy;
}

export async function createCanonicalResourceInstallerFrameManager(
  options: CreateCanonicalResourceInstallerFrameManagerOptions,
): Promise<ResourceInstallerFrameManager> {
  const requiredReleaseIdentity = parseRequiredReleaseIdentity(options.requiredReleaseIdentity);
  if (!requiredReleaseIdentity) throw { code: 'identity', path: 'release/identity' };
  const browserWindow = options.window ?? window;
  const browserDocument = options.document ?? document;
  const identity = resolveResourceInstallerFrameIdentity({
    locationOrigin: browserWindow.location.origin,
    documentReferrer: browserDocument.referrer,
    search: browserWindow.location.search,
  });
  if (!identity || browserWindow.parent === browserWindow) throw { code: 'identity' };
  const cacheStorage = options.cacheStorage ?? (typeof caches === 'undefined' ? undefined : caches);
  const indexedDb = options.indexedDb ?? (typeof indexedDB === 'undefined' ? undefined : indexedDB);
  if (!cacheStorage || !indexedDb) throw { code: 'storage' };

  const fetchImpl = options.fetch ?? browserWindow.fetch.bind(browserWindow);
  const installer = new CanonicalResourceInstaller({
    assetBaseUrl: identity.canonicalOrigin,
    fetch: fetchImpl,
    cacheStorage,
    indexedDb,
    locks: options.locks ?? browserWindow.navigator.locks,
    requiredReleaseIdentity,
    broadcast:
      options.broadcast ??
      (typeof BroadcastChannel === 'undefined' ? undefined : new BroadcastChannel('onlyoffice-canonical-resources-v1')),
  });
  const gated = new CanonicalResourceReadinessGate({
    installer,
    canonicalOrigin: identity.canonicalOrigin,
    localTestMode: identity.localTestMode,
    onProgress: options.reportProgress,
  });
  await gated.initialize();
  return gated;
}

declare global {
  interface Window {
    __ONLYOFFICE_CREATE_CANONICAL_RESOURCE_MANAGER__?: CreateResourceInstallerFrameManager;
  }
}

if (
  typeof window !== 'undefined' &&
  typeof document !== 'undefined' &&
  document.documentElement.dataset.onlyofficeResourceInstaller === 'true'
) {
  const createManager: CreateResourceInstallerFrameManager =
    window.__ONLYOFFICE_CREATE_CANONICAL_RESOURCE_MANAGER__ ??
    ((requiredReleaseIdentity, reportProgress) =>
      createCanonicalResourceInstallerFrameManager({
        requiredReleaseIdentity,
        reportProgress,
      }));
  try {
    startResourceInstallerFrame({ createManager });
  } catch {
    // An untrusted top-level navigation or embedding origin receives no capability.
  }
}

export type { ResourceInstallerFrameIdentity };
