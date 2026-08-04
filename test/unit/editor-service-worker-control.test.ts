import { describe, expect, it, vi } from 'vitest';
import { ensureControlledEditorServiceWorker } from '../../src/lib/editor-service-worker-control';

class FakeWorker extends EventTarget {
  readonly messages: unknown[] = [];
  state: ServiceWorkerState;
  readonly scriptURL: string;
  onPostMessage?: (message: unknown) => void;

  constructor(scriptURL: string, state: ServiceWorkerState) {
    super();
    this.scriptURL = scriptURL;
    this.state = state;
  }

  postMessage(message: unknown): void {
    this.messages.push(message);
    this.onPostMessage?.(message);
  }

  setState(state: ServiceWorkerState): void {
    this.state = state;
    this.dispatchEvent(new Event('statechange'));
  }
}

class FakeRegistration extends EventTarget {
  active: ServiceWorker | null = null;
  waiting: ServiceWorker | null = null;
  installing: ServiceWorker | null = null;
  update = vi.fn(async () => undefined);
}

class FakeServiceWorkerContainer extends EventTarget {
  controller: ServiceWorker | null = null;
  readonly registration: FakeRegistration;
  readonly register = vi.fn(async () => this.registration as unknown as ServiceWorkerRegistration);
  readonly ready: Promise<ServiceWorkerRegistration>;

  constructor(registration: FakeRegistration) {
    super();
    this.registration = registration;
    this.ready = Promise.resolve(registration as unknown as ServiceWorkerRegistration);
  }

  replaceController(worker: FakeWorker): void {
    this.controller = worker as unknown as ServiceWorker;
    this.dispatchEvent(new Event('controllerchange'));
  }
}

describe('editor Service Worker control', () => {
  it('updates and activates a waiting worker before the first prime request', async () => {
    const oldWorker = new FakeWorker('/document_editor_service_worker.js', 'activated');
    const newWorker = new FakeWorker('/document_editor_service_worker.js', 'installed');
    const registration = new FakeRegistration();
    const container = new FakeServiceWorkerContainer(registration);
    registration.active = oldWorker as unknown as ServiceWorker;
    container.controller = oldWorker as unknown as ServiceWorker;
    registration.update.mockImplementation(async () => {
      registration.waiting = newWorker as unknown as ServiceWorker;
    });
    newWorker.onPostMessage = (message) => {
      if ((message as { type?: string }).type !== 'SKIP_WAITING') return;
      registration.waiting = null;
      registration.active = newWorker as unknown as ServiceWorker;
      container.replaceController(newWorker);
    };

    await expect(ensureControlledEditorServiceWorker(container, '/document_editor_service_worker.js')).resolves.toBe(
      newWorker,
    );
    expect(registration.update).toHaveBeenCalledOnce();
    expect(container.register).toHaveBeenCalledWith('/document_editor_service_worker.js', {
      scope: '/',
      updateViaCache: 'none',
    });
    expect(newWorker.messages).toEqual([{ type: 'SKIP_WAITING' }]);
  });

  it('waits for an installing worker to finish before asking it to take over', async () => {
    const oldWorker = new FakeWorker('/document_editor_service_worker.js', 'activated');
    const installingWorker = new FakeWorker('/document_editor_service_worker.js', 'installing');
    const registration = new FakeRegistration();
    const container = new FakeServiceWorkerContainer(registration);
    registration.active = oldWorker as unknown as ServiceWorker;
    container.controller = oldWorker as unknown as ServiceWorker;
    registration.update.mockImplementation(async () => {
      registration.installing = installingWorker as unknown as ServiceWorker;
      window.setTimeout(() => installingWorker.setState('installed'), 0);
    });
    installingWorker.onPostMessage = (message) => {
      if ((message as { type?: string }).type !== 'SKIP_WAITING') return;
      registration.installing = null;
      registration.active = installingWorker as unknown as ServiceWorker;
      container.replaceController(installingWorker);
    };

    await expect(ensureControlledEditorServiceWorker(container, '/document_editor_service_worker.js')).resolves.toBe(
      installingWorker,
    );
    expect(installingWorker.messages).toEqual([{ type: 'SKIP_WAITING' }]);
  });

  it('keeps the active controller usable when an update check fails offline', async () => {
    const activeWorker = new FakeWorker('/document_editor_service_worker.js', 'activated');
    const registration = new FakeRegistration();
    const container = new FakeServiceWorkerContainer(registration);
    registration.active = activeWorker as unknown as ServiceWorker;
    container.controller = activeWorker as unknown as ServiceWorker;
    registration.update.mockRejectedValue(new Error('offline'));

    await expect(ensureControlledEditorServiceWorker(container, '/document_editor_service_worker.js')).resolves.toBe(
      activeWorker,
    );
  });

  it('fails with a control timeout when a waiting worker never takes over', async () => {
    const oldWorker = new FakeWorker('/document_editor_service_worker.js', 'activated');
    const newWorker = new FakeWorker('/document_editor_service_worker.js', 'installed');
    const registration = new FakeRegistration();
    const container = new FakeServiceWorkerContainer(registration);
    registration.active = oldWorker as unknown as ServiceWorker;
    container.controller = oldWorker as unknown as ServiceWorker;
    registration.update.mockImplementation(async () => {
      registration.waiting = newWorker as unknown as ServiceWorker;
    });

    await expect(
      ensureControlledEditorServiceWorker(container, '/document_editor_service_worker.js', { timeoutMs: 10 }),
    ).rejects.toThrow('Editor Service Worker control timed out');
  });
});
