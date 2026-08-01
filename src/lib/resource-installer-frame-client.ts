import {
  parseRequiredReleaseIdentity,
  type RequiredReleaseIdentity,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
  type ResourcePlanRequest,
} from './release-resources';
import {
  RESOURCE_INSTALLER_FRAME_PROTOCOL,
  RESOURCE_INSTALLER_RPC_TIMEOUT_MS,
  createResourceInstallerFrameUrl,
  parseResourceInstallerChallengeMessage,
  parseResourceInstallerServerMessage,
  resolveResourceInstallerClientIdentity,
  type ResourceInstallerFrameIdentity,
  type ResourceInstallerFrameCommand,
  type ResourceInstallerFrameErrorCode,
  type ResourceInstallerRequestMessage,
  type ResourceInstallerServerMessage,
} from './resource-installer-frame-protocol';

export interface ResourceInstallerFrameClientOptions {
  canonicalOrigin?: string;
  allowLocalTestMode?: boolean;
  requiredReleaseIdentity?: RequiredReleaseIdentity;
  connectTimeoutMs?: number;
  requestTimeoutMs?: number;
  window?: Window;
  document?: Document;
  createMessageChannel?: () => MessageChannel;
  randomUUID?: () => string;
  now?: () => number;
}

export type ResourceInstallerFrameSnapshotListener = (snapshot: ResourceInstallerSnapshot) => void;

type PendingRequest = {
  command: ResourceInstallerFrameCommand;
  message: ResourceInstallerRequestMessage | null;
  resolve: (message: ResourceInstallerServerMessage) => void;
  reject: (error: unknown) => void;
  timeout: number | null;
  recovering: boolean;
};

const ACTIVITY_COMMANDS = new Set<ResourceInstallerFrameCommand>(['APPLY', 'REPAIR', 'RESUME']);
const PAUSABLE_COMMANDS = new Set<ResourceInstallerFrameCommand>(['APPLY', 'REPAIR']);

function validTimeout(value: number | undefined): number {
  const timeout = value ?? RESOURCE_INSTALLER_RPC_TIMEOUT_MS;
  if (!Number.isSafeInteger(timeout) || timeout <= 0 || timeout > RESOURCE_INSTALLER_RPC_TIMEOUT_MS) {
    throw new TypeError('Invalid Resource Installer RPC timeout');
  }
  return timeout;
}

export class ResourceInstallerFrameRpcError extends Error {
  readonly code: ResourceInstallerFrameErrorCode;
  readonly retryable: boolean;
  readonly path?: string;

  constructor(error: { code: ResourceInstallerFrameErrorCode; retryable: boolean; path?: string }) {
    super(`Resource Installer RPC failed: ${error.code}`);
    this.name = 'ResourceInstallerFrameRpcError';
    this.code = error.code;
    this.retryable = error.retryable;
    this.path = error.path;
  }
}

export class ResourceInstallerFrameClient {
  readonly #options: ResourceInstallerFrameClientOptions;
  readonly #window: Window;
  readonly #document: Document;
  readonly #connectTimeoutMs: number;
  readonly #requestTimeoutMs: number;
  readonly #createMessageChannel: () => MessageChannel;
  readonly #randomUUID: () => string;
  readonly #now: () => number;
  readonly #requiredReleaseIdentity: RequiredReleaseIdentity | null;
  readonly #listeners = new Set<ResourceInstallerFrameSnapshotListener>();
  readonly #pending = new Map<string, PendingRequest>();
  readonly #expiredIds = new Set<string>();
  #frame: HTMLIFrameElement | null = null;
  #port: MessagePort | null = null;
  #snapshot: ResourceInstallerSnapshot | null = null;
  #connectPromise: Promise<ResourceInstallerSnapshot> | null = null;
  #recovering = false;
  #mutationQueue: Promise<unknown> = Promise.resolve();
  #windowMessageListener: ((event: MessageEvent) => void) | null = null;
  #pageHideListener: (() => void) | null = null;
  #destroyed = false;

  constructor(options: ResourceInstallerFrameClientOptions = {}) {
    this.#options = options;
    this.#window = options.window ?? window;
    this.#document = options.document ?? document;
    this.#connectTimeoutMs = validTimeout(options.connectTimeoutMs);
    this.#requestTimeoutMs = validTimeout(options.requestTimeoutMs);
    this.#createMessageChannel = options.createMessageChannel ?? (() => new MessageChannel());
    this.#randomUUID = options.randomUUID ?? (() => crypto.randomUUID());
    this.#now = options.now ?? Date.now;
    this.#requiredReleaseIdentity = parseRequiredReleaseIdentity(options.requiredReleaseIdentity);
  }

  getSnapshot(): ResourceInstallerSnapshot | null {
    return this.#snapshot;
  }

  subscribe(listener: ResourceInstallerFrameSnapshotListener): () => void {
    this.#listeners.add(listener);
    if (this.#snapshot) listener(this.#snapshot);
    return () => this.#listeners.delete(listener);
  }

  connect(): Promise<ResourceInstallerSnapshot> {
    if (this.#destroyed) return Promise.reject(new Error('Resource Installer frame client was destroyed'));
    if (this.#connectPromise) return this.#connectPromise;
    const connection = this.#openFrame();
    this.#connectPromise = connection;
    void connection.catch(() => {
      if (this.#connectPromise !== connection) return;
      this.#connectPromise = null;
      this.#discardTransport();
    });
    return connection;
  }

  refreshSnapshot(): Promise<ResourceInstallerSnapshot> {
    return this.#requestSnapshot({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: this.#nextId(),
      command: 'SNAPSHOT',
    });
  }

  plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    return this.#requestPlan({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: this.#nextId(),
      command: 'PLAN',
      request,
    });
  }

  apply(plan: ResourcePlan): Promise<void> {
    return this.#enqueueMutation(async () => {
      await this.#requestSnapshot({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: this.#nextId(),
        command: 'APPLY',
        plan,
      });
    });
  }

  checkForUpdates(): Promise<void> {
    return this.#enqueueMutation(async () => {
      await this.#requestSnapshot({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: this.#nextId(),
        command: 'CHECK_UPDATES',
      });
    });
  }

  checkHealth(): Promise<void> {
    return this.#enqueueMutation(async () => {
      await this.#requestSnapshot({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: this.#nextId(),
        command: 'CHECK_HEALTH',
      });
    });
  }

  repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void> {
    return this.#enqueueMutation(async () => {
      await this.#requestSnapshot({
        protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
        type: 'REQUEST',
        id: this.#nextId(),
        command: 'REPAIR',
        scope: options.scope,
      });
    });
  }

  async pause(): Promise<void> {
    await this.#requestSnapshot({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: this.#nextId(),
      command: 'PAUSE',
    });
  }

  async resume(): Promise<void> {
    await this.#requestSnapshot({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: this.#nextId(),
      command: 'RESUME',
    });
  }

  async cancel(): Promise<void> {
    await this.#requestSnapshot({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'REQUEST',
      id: this.#nextId(),
      command: 'CANCEL',
    });
  }

  destroy(): void {
    if (this.#destroyed) return;
    this.#destroyed = true;
    if (this.#windowMessageListener) {
      this.#window.removeEventListener('message', this.#windowMessageListener);
      this.#windowMessageListener = null;
    }
    if (this.#pageHideListener) {
      this.#window.removeEventListener('pagehide', this.#pageHideListener);
      this.#pageHideListener = null;
    }
    this.#rejectAll(new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true }));
    this.#discardTransport();
    this.#expiredIds.clear();
    this.#listeners.clear();
  }

  async #openFrame(): Promise<ResourceInstallerSnapshot> {
    const identity = this.#resolveIdentity();
    if (!identity || !this.#requiredReleaseIdentity) {
      throw new ResourceInstallerFrameRpcError({ code: 'identity', retryable: false });
    }
    if (!this.#document.body) {
      throw new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true });
    }
    const frame = this.#document.createElement('iframe');
    frame.hidden = true;
    frame.tabIndex = -1;
    frame.title = 'OnlyOffice resource installer';
    frame.setAttribute('aria-hidden', 'true');
    frame.setAttribute('sandbox', 'allow-same-origin allow-scripts');
    frame.setAttribute('allow', '');
    frame.referrerPolicy = 'origin';
    frame.src = createResourceInstallerFrameUrl(identity).href;
    this.#frame = frame;

    if (!this.#pageHideListener) {
      this.#pageHideListener = () => this.destroy();
      this.#window.addEventListener('pagehide', this.#pageHideListener, { once: true });
    }
    const result = this.#openPort(frame, identity, false);
    this.#document.body.append(frame);
    return result;
  }

  #openPort(
    frame: HTMLIFrameElement,
    identity: ResourceInstallerFrameIdentity,
    reconnect: boolean,
  ): Promise<ResourceInstallerSnapshot> {
    const result = new Promise<ResourceInstallerSnapshot>((resolve, reject) => {
      const connectTimer = this.#window.setTimeout(() => {
        reject(new ResourceInstallerFrameRpcError({ code: 'timeout', retryable: true }));
        if (this.#windowMessageListener) {
          this.#window.removeEventListener('message', this.#windowMessageListener);
          this.#windowMessageListener = null;
        }
      }, this.#connectTimeoutMs);
      let challengeAccepted = false;

      const fail = (error: ResourceInstallerFrameRpcError) => {
        this.#window.clearTimeout(connectTimer);
        reject(error);
        if (this.#windowMessageListener) {
          this.#window.removeEventListener('message', this.#windowMessageListener);
          this.#windowMessageListener = null;
        }
      };
      this.#windowMessageListener = (event: MessageEvent) => {
        if (this.#destroyed || event.source !== frame.contentWindow || event.origin !== identity.canonicalOrigin) {
          return;
        }
        if (challengeAccepted) {
          fail(new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false }));
          return;
        }
        const challenge = parseResourceInstallerChallengeMessage(
          event.data,
          identity.physicalParentOrigin,
          this.#now(),
        );
        if (!challenge) {
          fail(new ResourceInstallerFrameRpcError({ code: 'capability', retryable: false }));
          return;
        }
        challengeAccepted = true;
        const channel = this.#createMessageChannel();
        this.#port = channel.port1;
        channel.port1.onmessage = (portEvent) => this.#handlePortMessage(portEvent.data);
        channel.port1.onmessageerror = () =>
          this.#recoverConnection(new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true }));
        channel.port1.start();

        const id = this.#nextId();
        const ready = new Promise<ResourceInstallerServerMessage>((readyResolve, readyReject) => {
          const pending: PendingRequest = {
            command: 'INIT',
            message: null,
            resolve: readyResolve,
            reject: readyReject,
            timeout: null,
            recovering: false,
          };
          this.#pending.set(id, pending);
          this.#armPendingTimeout(id, pending, this.#connectTimeoutMs);
        });

        frame.contentWindow?.postMessage(
          {
            protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
            type: 'REQUEST',
            id,
            command: 'INIT',
            capability: {
              token: challenge.capability.token,
              parentOrigin: identity.physicalParentOrigin,
            },
            requiredReleaseIdentity: { ...this.#requiredReleaseIdentity! },
            subscribe: true,
          },
          identity.canonicalOrigin,
          [channel.port2],
        );

        void ready.then(
          (message) => {
            if (message.type !== 'READY') {
              fail(new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false }));
              return;
            }
            this.#window.clearTimeout(connectTimer);
            if (this.#windowMessageListener) {
              this.#window.removeEventListener('message', this.#windowMessageListener);
              this.#windowMessageListener = null;
            }
            this.#publish(message.snapshot);
            resolve(message.snapshot);
          },
          (error) =>
            fail(
              error instanceof ResourceInstallerFrameRpcError
                ? error
                : new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true }),
            ),
        );
      };
      this.#window.addEventListener('message', this.#windowMessageListener);
      if (reconnect) {
        frame.contentWindow?.postMessage(
          {
            protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
            type: 'RECONNECT',
          },
          identity.canonicalOrigin,
        );
      }
    });
    return result;
  }

  async #requestPlan(message: Extract<ResourceInstallerRequestMessage, { command: 'PLAN' }>): Promise<ResourcePlan> {
    const response = await this.#send(message);
    if (response.type !== 'RESULT' || response.command !== 'PLAN') {
      throw new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false });
    }
    return response.plan;
  }

  async #requestSnapshot(
    message: Exclude<ResourceInstallerRequestMessage, { command: 'PLAN' }>,
  ): Promise<ResourceInstallerSnapshot> {
    const response = await this.#send(message);
    if (response.type !== 'RESULT' || response.command !== message.command || !('snapshot' in response)) {
      throw new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false });
    }
    this.#publish(response.snapshot);
    return response.snapshot;
  }

  async #send(message: ResourceInstallerRequestMessage): Promise<ResourceInstallerServerMessage> {
    await this.connect();
    if (this.#destroyed || !this.#port) {
      throw new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true });
    }
    if (this.#pending.has(message.id)) {
      throw new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false });
    }
    return new Promise((resolve, reject) => {
      const pending: PendingRequest = {
        command: message.command,
        message,
        resolve,
        reject,
        timeout: null,
        recovering: false,
      };
      this.#pending.set(message.id, pending);
      this.#armPendingTimeout(message.id, pending);
      try {
        this.#port?.postMessage(message);
      } catch {
        this.#recoverConnection(new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true }));
      }
    });
  }

  #handlePortMessage(value: unknown): void {
    const message = parseResourceInstallerServerMessage(value);
    if (!message) {
      this.#failConnection(new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false }));
      return;
    }
    if (message.type === 'EVENT') {
      this.#publish(message.snapshot);
      this.#refreshActivity(message.snapshot);
      this.#reconcileRecovered(message.snapshot);
      return;
    }
    const pending = this.#pending.get(message.id);
    if (!pending || pending.command !== message.command) {
      if (this.#expiredIds.delete(message.id)) return;
      this.#failConnection(new ResourceInstallerFrameRpcError({ code: 'protocol', retryable: false }));
      return;
    }
    this.#pending.delete(message.id);
    if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
    if (message.type === 'ERROR') {
      pending.reject(new ResourceInstallerFrameRpcError(message.error));
      return;
    }
    pending.resolve(message);
  }

  #publish(snapshot: ResourceInstallerSnapshot): void {
    this.#snapshot = snapshot;
    for (const listener of this.#listeners) listener(snapshot);
  }

  #refreshActivity(snapshot: ResourceInstallerSnapshot): void {
    const paused = snapshot.phase === 'paused' || snapshot.readiness === 'paused';
    for (const [id, pending] of this.#pending) {
      if (!ACTIVITY_COMMANDS.has(pending.command)) continue;
      if (paused && PAUSABLE_COMMANDS.has(pending.command)) {
        if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
        pending.timeout = null;
        continue;
      }
      this.#armPendingTimeout(id, pending);
    }
  }

  #enqueueMutation<T>(task: () => Promise<T>): Promise<T> {
    const queued = this.#mutationQueue.catch(() => undefined).then(task);
    this.#mutationQueue = queued;
    return queued;
  }

  #nextId(): string {
    const raw = this.#randomUUID().replaceAll('-', '');
    return `rpc-${raw.slice(0, 120)}`;
  }

  #recoverConnection(error: ResourceInstallerFrameRpcError): void {
    if (this.#destroyed || this.#recovering) return;
    const identity = this.#resolveIdentity();
    const frame = this.#frame;
    if (!identity || !frame?.contentWindow) {
      this.#failConnection(error);
      return;
    }

    this.#recovering = true;
    if (this.#windowMessageListener) {
      this.#window.removeEventListener('message', this.#windowMessageListener);
      this.#windowMessageListener = null;
    }
    this.#port?.close();
    this.#port = null;
    for (const [id, pending] of this.#pending) {
      if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
      pending.timeout = null;
      if (pending.command === 'INIT') {
        this.#pending.delete(id);
        pending.reject(error);
      } else {
        pending.recovering = true;
      }
    }

    const connection = this.#openPort(frame, identity, true);
    this.#connectPromise = connection;
    void connection.then(
      (snapshot) => {
        if (this.#destroyed || this.#connectPromise !== connection) return;
        this.#recovering = false;
        this.#reconcileRecovered(snapshot);
      },
      (connectionError) => {
        if (this.#connectPromise !== connection) return;
        this.#recovering = false;
        this.#failConnection(connectionError instanceof ResourceInstallerFrameRpcError ? connectionError : error);
      },
    );
  }

  #reconcileRecovered(snapshot: ResourceInstallerSnapshot): void {
    for (const [id, pending] of this.#pending) {
      if (!pending.recovering || !pending.message) continue;
      const command = pending.command;
      const active =
        snapshot.phase === 'planning' ||
        snapshot.phase === 'downloading' ||
        snapshot.phase === 'verifying' ||
        snapshot.phase === 'activating' ||
        snapshot.phase === 'repairing' ||
        snapshot.phase === 'paused';
      const successful =
        snapshot.phase === 'idle' && (snapshot.readiness === 'ready' || snapshot.readiness === 'update-available');
      const failed =
        snapshot.phase === 'idle' &&
        (snapshot.readiness === 'error' || snapshot.readiness === 'repair-needed' || snapshot.errorCode !== null);

      if (command === 'APPLY') {
        const releaseId = pending.message.command === 'APPLY' ? pending.message.plan.releaseId : null;
        if (successful && snapshot.installedRelease === releaseId) {
          this.#resolveRecovered(id, pending, snapshot);
          continue;
        }
        if (active && snapshot.targetRelease === releaseId) {
          this.#armRecoveredActivity(id, pending, snapshot);
          continue;
        }
      } else if (command === 'REPAIR' || command === 'RESUME') {
        if (successful) {
          this.#resolveRecovered(id, pending, snapshot);
          continue;
        }
        if (active && !(command === 'RESUME' && snapshot.phase === 'paused')) {
          this.#armRecoveredActivity(id, pending, snapshot);
          continue;
        }
      } else if (
        (command === 'PAUSE' && snapshot.phase === 'paused') ||
        (command === 'CANCEL' && snapshot.phase === 'idle')
      ) {
        this.#resolveRecovered(id, pending, snapshot);
        continue;
      }

      if (failed && (command === 'APPLY' || command === 'REPAIR' || command === 'RESUME')) {
        this.#rejectRecovered(
          id,
          pending,
          new ResourceInstallerFrameRpcError({
            code: snapshot.errorCode ?? 'unavailable',
            retryable:
              snapshot.errorCode !== 'integrity' &&
              snapshot.errorCode !== 'manifest' &&
              snapshot.errorCode !== 'incompatible',
            path: snapshot.failedResources[0]?.path,
          }),
        );
        continue;
      }

      pending.recovering = false;
      this.#armPendingTimeout(id, pending);
      try {
        this.#port?.postMessage(pending.message);
      } catch {
        this.#recoverConnection(new ResourceInstallerFrameRpcError({ code: 'unavailable', retryable: true }));
      }
    }
  }

  #armRecoveredActivity(id: string, pending: PendingRequest, snapshot: ResourceInstallerSnapshot): void {
    if ((snapshot.phase === 'paused' || snapshot.readiness === 'paused') && PAUSABLE_COMMANDS.has(pending.command)) {
      if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
      pending.timeout = null;
      return;
    }
    this.#armPendingTimeout(id, pending);
  }

  #resolveRecovered(id: string, pending: PendingRequest, snapshot: ResourceInstallerSnapshot): void {
    this.#pending.delete(id);
    if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
    pending.resolve({
      protocol: RESOURCE_INSTALLER_FRAME_PROTOCOL,
      type: 'RESULT',
      id,
      command: pending.command as Exclude<ResourceInstallerFrameCommand, 'INIT' | 'PLAN'>,
      snapshot,
    });
  }

  #rejectRecovered(id: string, pending: PendingRequest, error: ResourceInstallerFrameRpcError): void {
    this.#pending.delete(id);
    if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
    pending.reject(error);
  }

  #failConnection(error: ResourceInstallerFrameRpcError): void {
    this.#rejectAll(error);
    this.#connectPromise = null;
    this.#discardTransport();
  }

  #rejectAll(error: unknown): void {
    for (const pending of this.#pending.values()) {
      if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
      pending.reject(error);
    }
    this.#pending.clear();
  }

  #armPendingTimeout(id: string, pending: PendingRequest, timeoutMs = this.#requestTimeoutMs): void {
    if (pending.timeout !== null) this.#window.clearTimeout(pending.timeout);
    pending.timeout = this.#window.setTimeout(() => {
      if (this.#pending.get(id) !== pending) return;
      this.#pending.delete(id);
      pending.timeout = null;
      this.#rememberExpired(id);
      pending.reject(new ResourceInstallerFrameRpcError({ code: 'timeout', retryable: true }));
    }, timeoutMs);
  }

  #rememberExpired(id: string): void {
    this.#expiredIds.add(id);
    while (this.#expiredIds.size > 128) {
      const oldest = this.#expiredIds.values().next().value;
      if (typeof oldest !== 'string') break;
      this.#expiredIds.delete(oldest);
    }
  }

  #resolveIdentity(): ResourceInstallerFrameIdentity | null {
    return resolveResourceInstallerClientIdentity({
      parentOrigin: this.#window.location.origin,
      canonicalOrigin: this.#options.canonicalOrigin,
      allowLocalTestMode: this.#options.allowLocalTestMode,
    });
  }

  #discardTransport(): void {
    if (this.#windowMessageListener) {
      this.#window.removeEventListener('message', this.#windowMessageListener);
      this.#windowMessageListener = null;
    }
    this.#port?.close();
    this.#port = null;
    this.#frame?.remove();
    this.#frame = null;
    this.#recovering = false;
  }
}

export function createResourceInstallerFrameClient(
  options: ResourceInstallerFrameClientOptions = {},
): ResourceInstallerFrameClient {
  return new ResourceInstallerFrameClient(options);
}
