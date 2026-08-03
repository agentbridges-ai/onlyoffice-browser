import {
  EditorOriginPrimer,
  EditorOriginPrimeError,
  editorOriginForSlot,
  type EditorOriginPrimeProgress,
  type EditorOriginPrimeResult,
} from './editor-origin-primer';
import {
  EDITOR_SHELL_HOST_PATH,
  EDITOR_SHELL_MAX_ASSETS,
  EDITOR_SHELL_MAX_TOTAL_BYTES,
  EDITOR_SHELL_PRIME_PATH,
  isEditorShellAssetPath,
} from './editor-shell-cache';
import { OFFICE_EDITOR_ORIGIN_SLOTS } from './office-origin-pool';
import {
  ResourceInstallerError,
  type OfficeRuntimeResourceInstaller,
  type ResourceErrorCode,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
  type ResourcePlanRequest,
} from './release-resources';

const READINESS_PROOF_CHANNEL = 'onlyoffice-canonical-readiness-v1';
const READINESS_PROOF_PROTOCOL = 'onlyoffice-canonical-readiness-proof-v1';
// A successful readiness probe is an immutable release proof. Re-running all
// 12 editor-origin shell checks on every settings-panel refresh only adds
// latency and can re-trigger third-party storage failures. Revalidate on a
// bounded heartbeat instead; explicit updates and mutations always bypass it.
const READINESS_PROOF_MAX_AGE_MS = 10 * 60 * 1000;

export type CanonicalReleaseReadinessProbe = (
  releaseId: string,
  onProgress?: EditorOriginPrimeProgress,
) => Promise<EditorOriginPrimeResult[]>;

export type CanonicalResourceReadinessGateOptions = {
  installer: OfficeRuntimeResourceInstaller & {
    initialize?: () => Promise<void>;
    rollbackActivation?: (releaseId: string, failure: { code: ResourceErrorCode; path: string }) => Promise<void>;
  };
  probeRelease?: CanonicalReleaseReadinessProbe;
  primeRelease?: CanonicalReleaseReadinessProbe;
  verifyRelease?: CanonicalReleaseReadinessProbe;
  canonicalOrigin?: string;
  localTestMode?: boolean;
  onProgress?: () => void;
};

type GateFailure = {
  releaseId: string;
  code: ResourceErrorCode;
  path: string;
};

function cloneSnapshot(snapshot: ResourceInstallerSnapshot): ResourceInstallerSnapshot {
  return {
    ...snapshot,
    installedProfiles: [...snapshot.installedProfiles],
    failedResources: snapshot.failedResources.map((failure) => ({ ...failure })),
  };
}

function gateError(error: unknown): ResourceInstallerError {
  if (error instanceof ResourceInstallerError) return error;
  if (
    error instanceof EditorOriginPrimeError ||
    (error instanceof Error &&
      error.name === 'EditorOriginPrimeError' &&
      typeof (error as Partial<EditorOriginPrimeError>).origin === 'string' &&
      typeof (error as Partial<EditorOriginPrimeError>).code === 'string' &&
      typeof (error as Partial<EditorOriginPrimeError>).stage === 'string')
  ) {
    const primeError = error as EditorOriginPrimeError;
    const code: ResourceErrorCode =
      primeError.code === 'timeout'
        ? 'timeout'
        : primeError.code === 'cancelled'
          ? 'aborted'
          : primeError.code === 'identity'
            ? 'incompatible'
            : 'storage';
    const origin = new URL(primeError.origin).hostname;
    return new ResourceInstallerError(code, `editor-origins/${origin}/${primeError.stage}/${primeError.code}`);
  }
  const offline = typeof navigator !== 'undefined' && navigator.onLine === false;
  return new ResourceInstallerError(offline ? 'offline' : 'storage', 'editor-origins/broker-probe');
}

function validateProbeResults(
  releaseId: string,
  results: EditorOriginPrimeResult[],
  expectedOrigins: ReadonlySet<string>,
): void {
  if (
    results.length !== expectedOrigins.size ||
    new Set(results.map((result) => result.origin)).size !== results.length ||
    results.some((result) => !expectedOrigins.has(result.origin))
  ) {
    throw new ResourceInstallerError('storage', 'editor-origins/incomplete');
  }
  let brokerReadCount = 0;
  for (const result of results) {
    if (
      result.releaseId !== releaseId ||
      result.brokerReady === result.occupied ||
      !result.serviceWorkerVersion ||
      result.cachedPaths.length < 3 ||
      result.cachedPaths.length > EDITOR_SHELL_MAX_ASSETS ||
      new Set(result.cachedPaths).size !== result.cachedPaths.length ||
      !result.cachedPaths.every(isEditorShellAssetPath) ||
      !result.cachedPaths.includes(EDITOR_SHELL_HOST_PATH) ||
      !result.cachedPaths.includes(EDITOR_SHELL_PRIME_PATH) ||
      !Number.isSafeInteger(result.cachedBytes) ||
      result.cachedBytes <= 0 ||
      result.cachedBytes > EDITOR_SHELL_MAX_TOTAL_BYTES ||
      (result.storageMode !== 'cache' && result.storageMode !== 'network')
    ) {
      throw new ResourceInstallerError('storage', `editor-origins/${result.origin}`);
    }
    if (result.brokerReady) brokerReadCount += 1;
  }
  if (brokerReadCount === 0) {
    throw new ResourceInstallerError('storage', 'editor-origins/no-live-broker-probe');
  }
}

/**
 * Keeps the public installer snapshot fail-closed until the canonical objects,
 * canonical Broker, Editor Service Worker protocol, and every fixed editor
 * origin shell have been exercised for the active release. Origins occupied
 * by an older pinned editor may defer only the broker read; at least one free
 * origin must complete a real routed read before activation.
 */
export class CanonicalResourceReadinessGate implements OfficeRuntimeResourceInstaller {
  readonly #installer: CanonicalResourceReadinessGateOptions['installer'];
  readonly #primeRelease: CanonicalReleaseReadinessProbe;
  readonly #verifyRelease: CanonicalReleaseReadinessProbe;
  readonly #expectedOrigins: ReadonlySet<string>;
  readonly #onProgress: (() => void) | null;
  readonly #listeners = new Set<(snapshot: ResourceInstallerSnapshot) => void>();
  readonly #proofChannel: BroadcastChannel | null;
  #rawSnapshot: ResourceInstallerSnapshot;
  #validatedRelease: string | null = null;
  #peerValidatedRelease: string | null = null;
  #validatedAt = 0;
  #gateFailure: GateFailure | null = null;
  #gateRelease: string | null = null;
  #gateMode: 'prime' | 'verify' | null = null;
  #gatePromise: Promise<void> | null = null;

  constructor(options: CanonicalResourceReadinessGateOptions) {
    this.#installer = options.installer;
    this.#rawSnapshot = cloneSnapshot(options.installer.getInstallerSnapshot());
    let primer: EditorOriginPrimer | null = null;
    const editorSlots = options.localTestMode ? OFFICE_EDITOR_ORIGIN_SLOTS.slice(0, 3) : OFFICE_EDITOR_ORIGIN_SLOTS;
    const getPrimer = () =>
      (primer ??= new EditorOriginPrimer({
        canonicalOrigin: options.canonicalOrigin,
        localMatrix: options.localTestMode,
      }));
    this.#primeRelease =
      options.primeRelease ??
      options.probeRelease ??
      ((releaseId, onProgress) => getPrimer().primeRelease(releaseId, editorSlots, onProgress));
    this.#verifyRelease =
      options.verifyRelease ??
      options.probeRelease ??
      ((releaseId, onProgress) => getPrimer().verifyRelease(releaseId, editorSlots, onProgress));
    const canonicalOrigin = options.canonicalOrigin ?? 'https://onlyoffice.getpi.work';
    this.#expectedOrigins = new Set(
      editorSlots.map((slot) => editorOriginForSlot(slot, canonicalOrigin, options.localTestMode)),
    );
    this.#onProgress = options.onProgress ?? null;
    this.#proofChannel = typeof BroadcastChannel === 'function' ? new BroadcastChannel(READINESS_PROOF_CHANNEL) : null;
    if (this.#proofChannel) {
      this.#proofChannel.onmessage = (event) => {
        const value = event.data;
        if (
          !value ||
          typeof value !== 'object' ||
          Array.isArray(value) ||
          Object.keys(value).length !== 3 ||
          value.protocol !== READINESS_PROOF_PROTOCOL ||
          value.type !== 'READY' ||
          typeof value.releaseId !== 'string'
        ) {
          return;
        }
        this.#peerValidatedRelease = value.releaseId;
        if (this.#rawSnapshot.installedRelease === value.releaseId) {
          this.#validatedRelease = value.releaseId;
          this.#validatedAt = Date.now();
          this.#gateFailure = null;
          this.#publish();
        }
      };
    }
    options.installer.subscribeInstaller((snapshot) => {
      this.#rawSnapshot = cloneSnapshot(snapshot);
      if (snapshot.installedRelease && snapshot.installedRelease === this.#peerValidatedRelease) {
        this.#validatedRelease = snapshot.installedRelease;
        this.#gateFailure = null;
      }
      this.#publish();
    });
  }

  async initialize(): Promise<void> {
    await this.#installer.initialize?.();
    this.#rawSnapshot = cloneSnapshot(this.#installer.getInstallerSnapshot());
    await this.#ensureReady('verify');
  }

  getInstallerSnapshot(): ResourceInstallerSnapshot {
    return this.#publicSnapshot();
  }

  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  getInstalledPaths(): string[] {
    return this.#installer.getInstalledPaths();
  }

  plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    return this.#installer.plan(request);
  }

  async apply(plan: ResourcePlan): Promise<void> {
    await this.#mutateAndGate(() => this.#installer.apply(plan));
  }

  async checkForUpdates(): Promise<void> {
    await this.#installer.checkForUpdates();
    this.#rawSnapshot = cloneSnapshot(this.#installer.getInstallerSnapshot());
    this.#invalidateReadinessProof();
    await this.#ensureReady('verify');
  }

  async checkHealth(options?: { deep?: false }): Promise<void> {
    await this.#installer.checkHealth(options);
    this.#rawSnapshot = cloneSnapshot(this.#installer.getInstallerSnapshot());
    const releaseId = this.#rawSnapshot.installedRelease;
    if (
      releaseId &&
      this.#validatedRelease === releaseId &&
      Date.now() - this.#validatedAt < READINESS_PROOF_MAX_AGE_MS
    ) {
      this.#publish();
      return;
    }
    this.#invalidateReadinessProof();
    await this.#ensureReady('verify');
  }

  async repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void> {
    await this.#mutateAndGate(() => this.#installer.repair(options));
  }

  pause(): void {
    this.#installer.pause();
  }

  async resume(): Promise<void> {
    await this.#mutateAndGate(() => this.#installer.resume());
  }

  cancel(): void {
    this.#installer.cancel();
  }

  async #mutateAndGate(operation: () => Promise<void>): Promise<void> {
    const previousRelease = this.#rawSnapshot.installedRelease;
    this.#gateFailure = null;
    this.#publish();
    await operation();
    this.#rawSnapshot = cloneSnapshot(this.#installer.getInstallerSnapshot());
    const activatedRelease = this.#rawSnapshot.installedRelease;
    try {
      await this.#ensureReady('prime');
    } catch (error) {
      const failure = gateError(error);
      if (activatedRelease && activatedRelease !== previousRelease && this.#installer.rollbackActivation) {
        try {
          await this.#installer.rollbackActivation(activatedRelease, {
            code: failure.code,
            path: failure.path || 'editor-origins/broker-probe',
          });
          this.#rawSnapshot = cloneSnapshot(this.#installer.getInstallerSnapshot());
        } catch {
          this.#validatedRelease = null;
          this.#gateFailure = {
            releaseId: activatedRelease,
            code: 'storage',
            path: 'release/activation-rollback',
          };
          this.#rawSnapshot = cloneSnapshot(this.#installer.getInstallerSnapshot());
          this.#publish();
          throw new ResourceInstallerError('storage', 'release/activation-rollback');
        }
        this.#publish();
      }
      throw failure;
    }
  }

  async #ensureReady(mode: 'prime' | 'verify'): Promise<void> {
    const releaseId = this.#rawSnapshot.installedRelease;
    if (
      !releaseId ||
      (this.#rawSnapshot.readiness !== 'ready' && this.#rawSnapshot.readiness !== 'update-available') ||
      this.#validatedRelease === releaseId
    ) {
      return;
    }
    if (this.#gatePromise && this.#gateRelease === releaseId && this.#gateMode === mode) {
      return this.#gatePromise;
    }

    this.#gateRelease = releaseId;
    this.#gateMode = mode;
    this.#gateFailure = null;
    let operation!: Promise<void>;
    operation = (async () => {
      this.#publish();
      try {
        const reportProgress: EditorOriginPrimeProgress = () => this.#publish();
        const results = await (mode === 'prime'
          ? this.#primeRelease(releaseId, reportProgress)
          : this.#verifyRelease(releaseId, reportProgress));
        validateProbeResults(releaseId, results, this.#expectedOrigins);
        if (this.#installer.getInstallerSnapshot().installedRelease !== releaseId) {
          throw new ResourceInstallerError('incompatible', 'release/changed-during-readiness-probe');
        }
        this.#validatedRelease = releaseId;
        this.#validatedAt = Date.now();
        this.#gateFailure = null;
        this.#peerValidatedRelease = releaseId;
        this.#proofChannel?.postMessage({
          protocol: READINESS_PROOF_PROTOCOL,
          type: 'READY',
          releaseId,
        });
      } catch (error) {
        if (this.#validatedRelease === releaseId || this.#peerValidatedRelease === releaseId) return;
        const failure = gateError(error);
        this.#validatedRelease = null;
        this.#gateFailure = {
          releaseId,
          code: failure.code,
          path: failure.path || 'editor-origins/broker-probe',
        };
        throw failure;
      } finally {
        if (this.#gatePromise === operation) {
          this.#gatePromise = null;
          this.#gateRelease = null;
          this.#gateMode = null;
        }
        this.#publish();
      }
    })();
    this.#gatePromise = operation;
    return operation;
  }

  #invalidateReadinessProof(): void {
    this.#validatedRelease = null;
    this.#peerValidatedRelease = null;
    this.#validatedAt = 0;
    this.#gateFailure = null;
    this.#publish();
  }

  #publicSnapshot(): ResourceInstallerSnapshot {
    const snapshot = cloneSnapshot(this.#rawSnapshot);
    const releaseId = snapshot.installedRelease;
    const gateFailure = this.#gateFailure;
    if (gateFailure?.releaseId === releaseId) {
      return {
        ...snapshot,
        readiness: 'repair-needed',
        phase: 'idle',
        errorCode: gateFailure.code,
        failedResources: [
          {
            path: gateFailure.path,
            code: gateFailure.code,
            attempts: 1,
          },
        ],
        canPause: false,
        canResume: false,
        canRetry: true,
      };
    }
    if (
      gateFailure &&
      snapshot.phase === 'idle' &&
      (snapshot.targetRelease === gateFailure.releaseId || snapshot.availableRelease === gateFailure.releaseId)
    ) {
      return {
        ...snapshot,
        readiness: releaseId ? 'update-available' : 'error',
        errorCode: gateFailure.code,
        failedResources: [
          {
            path: gateFailure.path,
            code: gateFailure.code,
            attempts: 1,
          },
        ],
        canPause: false,
        canResume: false,
        canRetry: true,
      };
    }
    if (!releaseId || this.#validatedRelease === releaseId) return snapshot;
    if (snapshot.readiness === 'ready' || snapshot.readiness === 'update-available') {
      return {
        ...snapshot,
        readiness: 'updating',
        phase: 'activating',
        errorCode: null,
        failedResources: [],
        canPause: false,
        canResume: false,
        canRetry: false,
      };
    }
    return snapshot;
  }

  #publish(): void {
    const snapshot = this.#publicSnapshot();
    this.#onProgress?.();
    for (const listener of this.#listeners) listener(cloneSnapshot(snapshot));
  }
}
