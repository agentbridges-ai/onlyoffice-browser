import {
  ResourceInstallerFrameClient,
  ResourceInstallerFrameRpcError,
  type ResourceInstallerFrameClientOptions,
} from './resource-installer-frame-client';
import {
  ResourceInstallerError,
  type OfficeRuntimeResourceInstaller,
  type ResourceErrorCode,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
  type ResourcePlanRequest,
} from './release-resources';

const RESOURCE_ERROR_CODES = new Set<ResourceErrorCode>([
  'offline',
  'network',
  'timeout',
  'integrity',
  'quota',
  'manifest',
  'incompatible',
  'storage',
  'aborted',
]);

function installerError(error: unknown): ResourceInstallerError {
  if (error instanceof ResourceInstallerError) return error;
  if (error instanceof ResourceInstallerFrameRpcError) {
    const code: ResourceErrorCode = RESOURCE_ERROR_CODES.has(error.code as ResourceErrorCode)
      ? (error.code as ResourceErrorCode)
      : error.code === 'identity' || error.code === 'capability' || error.code === 'protocol'
        ? 'incompatible'
        : 'storage';
    return new ResourceInstallerError(code, error.path);
  }
  return new ResourceInstallerError('storage');
}

/**
 * Metadata-only adapter for Piwork. Resource bytes never cross this RPC: the
 * remote frame executes every mutation against the canonical origin and sends
 * only plans, structured errors, and snapshots back to the parent.
 */
export class ResourceInstallerFrameAdapter implements OfficeRuntimeResourceInstaller {
  readonly #client: ResourceInstallerFrameClient;
  #snapshot: ResourceInstallerSnapshot;

  private constructor(client: ResourceInstallerFrameClient, snapshot: ResourceInstallerSnapshot) {
    this.#client = client;
    this.#snapshot = snapshot;
    client.subscribe((next) => {
      this.#snapshot = next;
    });
  }

  static async create(options: ResourceInstallerFrameClientOptions = {}): Promise<ResourceInstallerFrameAdapter> {
    const client = new ResourceInstallerFrameClient(options);
    try {
      const snapshot = await client.connect();
      return new ResourceInstallerFrameAdapter(client, snapshot);
    } catch (error) {
      client.destroy();
      throw installerError(error);
    }
  }

  getInstallerSnapshot(): ResourceInstallerSnapshot {
    return {
      ...this.#snapshot,
      installedProfiles: [...this.#snapshot.installedProfiles],
      failedResources: this.#snapshot.failedResources.map((failure) => ({ ...failure })),
    };
  }

  subscribeInstaller(listener: (snapshot: ResourceInstallerSnapshot) => void): () => void {
    return this.#client.subscribe((snapshot) => {
      this.#snapshot = snapshot;
      listener(this.getInstallerSnapshot());
    });
  }

  getInstalledPaths(): string[] {
    // Paths and resource bytes stay on the canonical origin. Profile readiness
    // in the typed snapshot is the cross-origin authority.
    return [];
  }

  async plan(request: ResourcePlanRequest): Promise<ResourcePlan> {
    try {
      return await this.#client.plan(request);
    } catch (error) {
      throw installerError(error);
    }
  }

  async apply(plan: ResourcePlan): Promise<void> {
    await this.#run(() => this.#client.apply(plan));
  }

  async checkForUpdates(): Promise<void> {
    await this.#run(() => this.#client.checkForUpdates());
  }

  async checkHealth(): Promise<void> {
    await this.#run(() => this.#client.checkHealth());
  }

  async repair(options: { scope: 'required' | 'installed' | 'all' }): Promise<void> {
    await this.#run(() => this.#client.repair(options));
  }

  pause(): void {
    void this.#client.pause().catch(() => undefined);
  }

  async resume(): Promise<void> {
    await this.#run(() => this.#client.resume());
  }

  cancel(): void {
    void this.#client.cancel().catch(() => undefined);
  }

  async #run(operation: () => Promise<void>): Promise<void> {
    try {
      await operation();
      const snapshot = this.#client.getSnapshot();
      if (snapshot) this.#snapshot = snapshot;
    } catch (error) {
      throw installerError(error);
    }
  }
}
