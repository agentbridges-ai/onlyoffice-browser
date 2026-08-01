import { mountOfficeEditor, type OfficeEditorInstance, type OfficeEditorMount } from './lib/office-editor';
import { resolveDemoHostUrl } from './lib/demo-host-url';

type BurstOptions = {
  count?: number;
  intervalMs?: number;
  activationBudget?: number;
};

type BurstStatus = {
  mounted: number;
  ready: number;
  errors: string[];
  inFlight: number;
  maxInFlight: number;
  retries: number;
  uniqueOrigins: number;
  outerFrames: number;
  done: boolean;
};

type BurstRecord = {
  mount: OfficeEditorMount;
  instance: OfficeEditorInstance | null;
  container: HTMLElement;
  file: File;
  attempts: number;
};

const grid = document.querySelector<HTMLElement>('#burst-grid')!;
const statusElement = document.querySelector<HTMLElement>('#burst-status')!;
const records: BurstRecord[] = [];
const pageUrl = new URL(window.location.href);
const configuredReleaseId = pageUrl.searchParams.get('releaseId');
const hostUrlResolver = resolveDemoHostUrl(pageUrl);
let status: BurstStatus = emptyStatus();
let completion: Promise<void> = Promise.resolve();

function emptyStatus(): BurstStatus {
  return {
    mounted: 0,
    ready: 0,
    errors: [],
    inFlight: 0,
    maxInFlight: 0,
    retries: 0,
    uniqueOrigins: 0,
    outerFrames: 0,
    done: false,
  };
}

function renderStatus(): void {
  const origins = new Set(
    records
      .map((record) => record.container.querySelector<HTMLIFrameElement>('iframe')?.src)
      .filter((src): src is string => Boolean(src))
      .map((src) => new URL(src).origin),
  );
  status.uniqueOrigins = origins.size;
  status.outerFrames = records.filter((record) => record.container.querySelector('iframe')).length;
  statusElement.textContent = JSON.stringify(status);
}

function delay(ms: number): Promise<void> {
  return ms > 0 ? new Promise((resolve) => window.setTimeout(resolve, ms)) : Promise.resolve();
}

async function loadFixtures(): Promise<Array<{ name: string; type: string; bytes: ArrayBuffer }>> {
  const fixtures = [
    ['local.docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'],
    ['local.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'],
    ['local.pptx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'],
  ] as const;
  return Promise.all(
    fixtures.map(async ([name, type]) => {
      const response = await fetch(`/fixtures/office/${name}`);
      if (!response.ok) throw new Error(`Unable to load ${name}: ${response.status}`);
      return { name, type, bytes: await response.arrayBuffer() };
    }),
  );
}

function createRecordMount(record: Pick<BurstRecord, 'container' | 'file'>): OfficeEditorMount {
  return mountOfficeEditor(record.container, {
    hostUrl: (context) => {
      const base = typeof hostUrlResolver === 'function' ? hostUrlResolver(context) : hostUrlResolver;
      const hostUrl = new URL(base, window.location.href);
      if (configuredReleaseId) {
        hostUrl.pathname = `/r/${encodeURIComponent(configuredReleaseId)}/office-host.html`;
      }
      return hostUrl.href;
    },
    file: record.file,
    fileName: record.file.name,
    destroyTimeoutMs: 5_000,
  });
}

async function run(options: BurstOptions): Promise<void> {
  await closeAll();
  status = emptyStatus();
  renderStatus();
  const count = Math.max(1, Math.floor(options.count ?? 20));
  const intervalMs = Math.max(0, Math.min(20, Math.floor(options.intervalMs ?? 10)));
  const activationBudget = 1;
  const fixtures = await loadFixtures();

  for (let index = 0; index < count; index += 1) {
    const fixture = fixtures[index % fixtures.length]!;
    const extension = fixture.name.split('.').pop()!;
    const fileName = `burst-${String(index + 1).padStart(2, '0')}.${extension}`;
    const container = document.createElement('section');
    container.className = 'burst-editor';
    container.dataset.burstIndex = String(index);
    if (index > 0) {
      container.inert = true;
      container.style.opacity = '0';
      container.style.pointerEvents = 'none';
    }
    grid.appendChild(container);
    const record: BurstRecord = {
      mount: null as unknown as OfficeEditorMount,
      instance: null,
      container,
      file: new File([fixture.bytes.slice(0)], fileName, { type: fixture.type }),
      attempts: 0,
    };
    record.mount = createRecordMount(record);
    records.push(record);
    status.mounted += 1;
    renderStatus();
    await delay(intervalMs);
  }

  const activationQueue = [...records];
  let nextIndex = 0;
  const activateNext = async () => {
    while (nextIndex < activationQueue.length) {
      const record = activationQueue[nextIndex++]!;
      status.inFlight += 1;
      status.maxInFlight = Math.max(status.maxInFlight, status.inFlight);
      renderStatus();
      try {
        record.instance = await record.mount.activate();
        status.ready += 1;
      } catch (error) {
        await record.mount.destroy().catch(() => undefined);
        if (record.attempts < 1) {
          record.attempts += 1;
          status.retries += 1;
          record.mount = createRecordMount(record);
          activationQueue.push(record);
        } else {
          status.errors.push(error instanceof Error ? error.message : String(error));
        }
      } finally {
        status.inFlight -= 1;
        renderStatus();
      }
    }
  };
  await Promise.all(Array.from({ length: activationBudget }, () => activateNext()));
  status.done = true;
  renderStatus();
}

async function closeAll(): Promise<void> {
  const current = records.splice(0);
  await Promise.allSettled(current.map((record) => record.mount.destroy()));
  grid.replaceChildren();
  status.outerFrames = 0;
  renderStatus();
}

const controller = {
  start(options: BurstOptions = {}) {
    completion = run(options);
  },
  wait: () => completion,
  closeAll,
  getStatus: (): BurstStatus => ({ ...status, errors: [...status.errors] }),
};

(window as typeof window & { __ONLYOFFICE_BURST_E2E__?: typeof controller }).__ONLYOFFICE_BURST_E2E__ = controller;
