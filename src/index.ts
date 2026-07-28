import { createOfficeEditor, type OfficeEditorInstance, type OfficeEditorMode } from './lib/office-editor';
import { Workbox } from 'workbox-window';
import { clearLegacyDemoHostState, resolveDemoHostUrl } from './lib/demo-host-url';
import { RuntimeCacheController, type RuntimeCacheProgress } from './lib/runtime-cache';
import './styles/base.css';

type DemoRecord = {
  id: number;
  instance: OfficeEditorInstance;
  panel: HTMLElement;
  status: HTMLElement;
  readonlyButton: HTMLButtonElement;
  closing: boolean;
};
type DemoEditorOptions = Omit<Parameters<typeof createOfficeEditor>[1], 'hostUrl'>;

const records: DemoRecord[] = [];
let nextPanelId = 1;
let selectedMode: OfficeEditorMode = 'edit';
const bootId = `office-demo-${Date.now()}-${Math.random().toString(36).slice(2)}`;

const app = document.querySelector<HTMLDivElement>('#app');
if (!app) {
  throw new Error('Missing #app root');
}

app.innerHTML = `
  <section class="demo-toolbar" aria-label="Office editor demo controls">
    <div>
      <h1>Browser Office Editor</h1>
      <p>Local DOCX, XLSX, PPTX, and CSV preview/edit component demo.</p>
      <button id="runtime-cache-status" class="runtime-cache-status" type="button" aria-live="polite">
        <span id="runtime-cache-label">Shared static assets: checking…</span>
        <progress id="runtime-cache-progress" max="1" value="0"></progress>
      </button>
    </div>
    <div class="demo-actions">
      <fieldset class="mode-selector" aria-label="Open mode">
        <label>
          <input type="radio" name="open-mode" value="edit" checked />
          Edit
        </label>
        <label>
          <input type="radio" name="open-mode" value="readonly" />
          Readonly
        </label>
        <label>
          <input type="radio" name="open-mode" value="preview" />
          Preview
        </label>
      </fieldset>
      <button id="new-word-button" type="button">New Word</button>
      <button id="new-excel-button" type="button">New Excel</button>
      <button id="new-pptx-button" type="button">New PowerPoint</button>
      <button id="new-csv-button" type="button">New CSV</button>
      <button id="upload-button" type="button">Open Files</button>
      <button id="close-all-button" type="button">Close All</button>
    </div>
  </section>
  <input id="file-input" type="file" multiple accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.csv" />
  <section id="editor-grid" class="editor-grid" aria-live="polite"></section>
  <dialog id="runtime-cache-dialog" class="runtime-cache-dialog">
    <form method="dialog">
      <h2>Load all shared static assets?</h2>
      <p id="runtime-cache-detail">
        The complete Office runtime is not cached yet. Sizes below are uncompressed browser-cache sizes;
        network transfer is smaller because Cloudflare uses Brotli compression.
      </p>
      <progress id="runtime-cache-dialog-progress" max="1" value="0"></progress>
      <p id="runtime-cache-dialog-status" class="runtime-cache-dialog-status"></p>
      <div id="runtime-cache-categories" class="runtime-cache-categories"></div>
      <section class="font-downloads" aria-labelledby="font-downloads-title">
        <h3 id="font-downloads-title">Font family</h3>
        <p>
          Microsoft YaHei is the default fallback. The first font package also installs the shared
          thumbnails, indexes, and compatibility fonts once; later families reuse them.
        </p>
        <div id="font-download-list" class="font-download-list"></div>
      </section>
      <div class="runtime-cache-dialog-actions">
        <button id="runtime-cache-later" value="cancel" type="submit">Later</button>
        <button id="runtime-cache-load" value="default" type="button">Load all</button>
      </div>
    </form>
  </dialog>
`;

const grid = document.querySelector<HTMLElement>('#editor-grid')!;
const fileInput = document.querySelector<HTMLInputElement>('#file-input')!;
const hardResetOnLastDestroy = new URLSearchParams(window.location.search).get('hardResetOnLastDestroy') === 'true';
clearLegacyDemoHostState(window.location, window.localStorage);
const defaultOfficeHostUrl = resolveDemoHostUrl(new URL(window.location.href));
const cacheStatusButton = document.querySelector<HTMLButtonElement>('#runtime-cache-status')!;
const cacheLabel = document.querySelector<HTMLElement>('#runtime-cache-label')!;
const cacheProgress = document.querySelector<HTMLProgressElement>('#runtime-cache-progress')!;
const cacheDialog = document.querySelector<HTMLDialogElement>('#runtime-cache-dialog')!;
const cacheDialogDetail = document.querySelector<HTMLElement>('#runtime-cache-detail')!;
const cacheDialogProgress = document.querySelector<HTMLProgressElement>('#runtime-cache-dialog-progress')!;
const cacheDialogStatus = document.querySelector<HTMLElement>('#runtime-cache-dialog-status')!;
const cacheCategories = document.querySelector<HTMLElement>('#runtime-cache-categories')!;
const fontDownloadList = document.querySelector<HTMLElement>('#font-download-list')!;
const cacheLoadButton = document.querySelector<HTMLButtonElement>('#runtime-cache-load')!;
const cacheLaterButton = document.querySelector<HTMLButtonElement>('#runtime-cache-later')!;
let runtimeCacheController: RuntimeCacheController | null = null;
let runtimeCacheLoading = false;

if ('serviceWorker' in navigator) {
  const workbox = new Workbox('/sw.js', { scope: '/' });
  let reloadingForServiceWorkerUpdate = false;
  workbox.addEventListener('waiting', () => {
    // Activate a new shell only when no editor can lose unsaved work.
    if (records.length === 0) void workbox.messageSkipWaiting();
  });
  workbox.addEventListener('controlling', () => {
    if (records.length > 0 || reloadingForServiceWorkerUpdate) return;
    reloadingForServiceWorkerUpdate = true;
    window.location.reload();
  });
  void workbox.register().then(() => {
    const updateInterval = window.setInterval(() => void workbox.update(), 60 * 60 * 1000);
    window.addEventListener('pagehide', () => window.clearInterval(updateInterval), { once: true });
  });
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
}

function cacheProgressText(progress: RuntimeCacheProgress): string {
  const byteText =
    progress.totalBytes > 0 ? ` · ${formatBytes(progress.completedBytes)} / ${formatBytes(progress.totalBytes)}` : '';
  const failureText = progress.failedFiles > 0 ? ` · ${progress.failedFiles} failed` : '';
  return `${progress.completedFiles} / ${progress.totalFiles} files${byteText} cached${failureText}`;
}

function renderCacheProgress(progress: RuntimeCacheProgress): void {
  const ratio =
    progress.totalBytes > 0
      ? progress.completedBytes / progress.totalBytes
      : progress.totalFiles > 0
        ? progress.completedFiles / progress.totalFiles
        : 0;
  cacheProgress.value = Math.min(1, ratio);
  cacheDialogProgress.value = Math.min(1, ratio);
  const detail = cacheProgressText(progress);
  cacheDialogStatus.textContent = detail;
  const categoryLabels = {
    fonts: 'Fonts',
    core: 'Common runtime & x2t',
    word: 'Word',
    cell: 'Spreadsheet',
    slide: 'Presentation',
  } as const;
  cacheCategories.replaceChildren(
    ...progress.categories.map((category) => {
      const ratio =
        category.totalBytes > 0
          ? category.completedBytes / category.totalBytes
          : category.totalFiles > 0
            ? category.completedFiles / category.totalFiles
            : 0;
      const row = document.createElement('div');
      row.className = 'runtime-cache-category';
      const label = document.createElement('div');
      label.innerHTML = `<span>${categoryLabels[category.category]}</span><span>${category.completedFiles} / ${category.totalFiles} · ${formatBytes(category.completedBytes)} / ${formatBytes(category.totalBytes)} cached</span>`;
      const bar = document.createElement('progress');
      bar.max = 1;
      bar.value = Math.min(1, ratio);
      row.append(label, bar);
      return row;
    }),
  );
  if (progress.phase === 'complete') {
    cacheLabel.textContent = `Shared static assets: ready · ${detail}`;
  } else if (progress.phase === 'error') {
    cacheLabel.textContent = `Shared static assets: incomplete · ${detail}`;
  } else if (progress.phase === 'loading') {
    cacheLabel.textContent = `Shared static assets: loading · ${detail}`;
  } else if (progress.phase === 'checking') {
    cacheLabel.textContent = `Shared static assets: checking · ${detail}`;
  } else {
    cacheLabel.textContent = `Shared static assets: not fully loaded · ${detail}`;
  }
}

function renderFontDownloads(): void {
  if (!runtimeCacheController) return;
  fontDownloadList.replaceChildren(
    ...runtimeCacheController.listFonts().map((font) => {
      const row = document.createElement('div');
      row.className = 'font-download-row';
      const label = document.createElement('span');
      label.textContent = `${font.name} · ${formatBytes(font.bytes)}`;
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = font.downloaded ? (font.removable ? 'Remove' : 'Downloaded') : 'Download';
      button.disabled = font.downloaded && !font.removable;
      button.addEventListener('click', async () => {
        if (!runtimeCacheController) return;
        button.disabled = true;
        if (font.downloaded && font.removable) {
          button.textContent = 'Removing…';
          renderCacheProgress(await runtimeCacheController.uninstallFontFamily(font.id));
        } else {
          button.textContent = 'Downloading…';
          if (navigator.storage?.persist) await navigator.storage.persist();
          await runtimeCacheController.downloadFontFamily(font.id, renderCacheProgress);
        }
        renderFontDownloads();
      });
      row.append(label, button);
      return row;
    }),
  );
}

function verifiedFontPaths(): string[] {
  return (
    runtimeCacheController
      ?.listFonts()
      .filter((font) => font.downloaded)
      .flatMap((font) => font.paths) || []
  );
}

function showCacheDialog(): void {
  if (!cacheDialog.open) cacheDialog.showModal();
}

async function loadAllRuntimeAssets(): Promise<void> {
  if (!runtimeCacheController || runtimeCacheLoading) return;
  runtimeCacheLoading = true;
  cacheLoadButton.disabled = true;
  cacheLaterButton.disabled = true;
  cacheDialogDetail.textContent =
    'Keep this page open while the shared Office runtime is loaded. Displayed bytes are uncompressed cache size; network transfer is Brotli-compressed.';
  try {
    if (navigator.storage?.persist) await navigator.storage.persist();
    if (navigator.storage?.estimate) {
      const estimate = await navigator.storage.estimate();
      const available =
        typeof estimate.quota === 'number' && typeof estimate.usage === 'number'
          ? Math.max(0, estimate.quota - estimate.usage)
          : Number.POSITIVE_INFINITY;
      if (available < runtimeCacheController.remainingBytes()) {
        throw new Error(
          `Not enough browser storage is available (${formatBytes(available)} free). Free space or leave private browsing before loading the Office runtime.`,
        );
      }
    }
    const result = await runtimeCacheController.loadAll(renderCacheProgress);
    if (result.phase === 'complete') {
      cacheDialogDetail.textContent = 'All shared static assets are loaded for the current release.';
      cacheLaterButton.textContent = 'Done';
      cacheLaterButton.disabled = false;
    } else {
      cacheDialogDetail.textContent = 'Some assets failed to load. Check the connection and retry.';
      cacheLoadButton.textContent = 'Retry';
      cacheLoadButton.disabled = false;
      cacheLaterButton.disabled = false;
    }
  } catch (error) {
    cacheDialogDetail.textContent = error instanceof Error ? error.message : String(error);
    cacheLoadButton.textContent = 'Retry';
    cacheLoadButton.disabled = false;
    cacheLaterButton.disabled = false;
  } finally {
    runtimeCacheLoading = false;
  }
}

async function initializeRuntimeCache(): Promise<void> {
  try {
    runtimeCacheController = await RuntimeCacheController.create();
    const progress = runtimeCacheController.getProgress(runtimeCacheController.isComplete() ? 'complete' : 'ready');
    renderCacheProgress(progress);
    renderFontDownloads();
    if (!runtimeCacheController.isComplete()) {
      showCacheDialog();
    } else if (runtimeCacheController.shouldCheckHealth()) {
      void runtimeCacheController.checkHealth(renderCacheProgress).then((health) => {
        renderFontDownloads();
        if (health.phase === 'error') showCacheDialog();
      });
    }
  } catch (error) {
    cacheLabel.textContent = `Shared static assets: status unavailable`;
    cacheDialogDetail.textContent = error instanceof Error ? error.message : String(error);
    cacheDialogStatus.textContent = '';
  }
}

cacheLoadButton.addEventListener('click', () => void loadAllRuntimeAssets());
cacheStatusButton.addEventListener('click', showCacheDialog);
void initializeRuntimeCache();

function isOfficeEditorMode(value: string): value is OfficeEditorMode {
  return value === 'edit' || value === 'readonly' || value === 'preview';
}

function getSelectedMode(): OfficeEditorMode {
  return selectedMode;
}

function setSelectedMode(mode: OfficeEditorMode): void {
  selectedMode = mode;
  for (const input of document.querySelectorAll<HTMLInputElement>('input[name="open-mode"]')) {
    input.checked = input.value === mode;
  }
}

function setStatus(record: DemoRecord, text: string): void {
  record.status.textContent = text;
}

function getModeLabel(mode: OfficeEditorMode): string {
  if (mode === 'readonly') return 'readonly';
  if (mode === 'preview') return 'preview';
  return 'editable';
}

function refreshPanelActions(record: DemoRecord): void {
  const state = record.instance.getState();
  const isPreview = state.mode === 'preview';
  record.readonlyButton.disabled = isPreview;
  record.readonlyButton.textContent = state.readonly ? 'Edit' : 'Readonly';
}

async function removeRecord(record: DemoRecord): Promise<void> {
  if (record.closing) return;
  record.closing = true;
  record.readonlyButton.disabled = true;
  const closeButton = record.panel.querySelector<HTMLButtonElement>('[data-action="close"]');
  if (closeButton) closeButton.disabled = true;
  setStatus(record, 'closing');

  const index = records.indexOf(record);
  if (index >= 0) {
    records.splice(index, 1);
  }
  await record.instance.destroy();
  record.panel.remove();
}

async function openEditor(options: DemoEditorOptions): Promise<DemoRecord> {
  const id = nextPanelId++;
  const title =
    options.fileName ||
    (options.file instanceof File ? options.file.name : undefined) ||
    options.emptyType ||
    'document';
  const panel = document.createElement('article');
  panel.className = 'editor-panel';
  panel.innerHTML = `
    <header class="editor-panel-header">
      <div>
        <strong>${title}</strong>
        <span data-role="status">opening</span>
      </div>
      <div class="panel-actions">
        <button type="button" data-action="readonly">Readonly</button>
        <button type="button" data-action="close">Close</button>
      </div>
    </header>
    <div class="editor-slot"></div>
  `;

  const slot = panel.querySelector<HTMLElement>('.editor-slot')!;
  const status = panel.querySelector<HTMLElement>('[data-role="status"]')!;
  const readonlyButton = panel.querySelector<HTMLButtonElement>('[data-action="readonly"]')!;
  grid.appendChild(panel);

  let instance: OfficeEditorInstance;
  try {
    instance = await createOfficeEditor(slot, {
      ...options,
      hostUrl: defaultOfficeHostUrl,
      downloadedFonts: verifiedFontPaths(),
      saveBehavior: options.saveBehavior || 'download',
      hardResetOnLastDestroy,
      onReady: (readyInstance) => {
        const state = readyInstance.getState();
        status.textContent = `${state.fileType.toUpperCase()} ${getModeLabel(state.mode)}`;
      },
      onSave: (file) => {
        status.textContent = `saved ${file.name} (${file.size} bytes)`;
        refreshPanelActions(record);
      },
      onDirtyChange: (dirty) => {
        status.textContent = dirty ? 'modified' : 'clean';
        refreshPanelActions(record);
      },
      onError: (error) => {
        status.textContent = `error: ${error.message}`;
      },
    });
  } catch (error) {
    panel.remove();
    throw error;
  }

  const record: DemoRecord = { id, instance, panel, status, readonlyButton, closing: false };
  records.push(record);
  const initialState = instance.getState();
  setStatus(record, `${initialState.fileType.toUpperCase()} ${getModeLabel(initialState.mode)}`);
  refreshPanelActions(record);

  readonlyButton.addEventListener('click', () => {
    const nextReadonly = !instance.getState().readonly;
    instance.setReadonly(nextReadonly);
    refreshPanelActions(record);
    const state = instance.getState();
    setStatus(record, `${state.fileType.toUpperCase()} ${getModeLabel(state.mode)}`);
  });

  panel.querySelector<HTMLButtonElement>('[data-action="close"]')?.addEventListener('click', () => {
    void removeRecord(record);
  });

  return record;
}

function openEmpty(emptyType: 'docx' | 'xlsx' | 'pptx' | 'csv'): void {
  void openEditor({
    emptyType,
    fileName: `New_Document.${emptyType}`,
    mode: getSelectedMode(),
  }).catch(() => undefined);
}

for (const input of document.querySelectorAll<HTMLInputElement>('input[name="open-mode"]')) {
  input.addEventListener('change', () => {
    if (input.checked && isOfficeEditorMode(input.value)) {
      setSelectedMode(input.value);
    }
  });
}

document.querySelector('#new-word-button')?.addEventListener('click', () => openEmpty('docx'));
document.querySelector('#new-excel-button')?.addEventListener('click', () => openEmpty('xlsx'));
document.querySelector('#new-pptx-button')?.addEventListener('click', () => openEmpty('pptx'));
document.querySelector('#new-csv-button')?.addEventListener('click', () => openEmpty('csv'));

document.querySelector('#upload-button')?.addEventListener('click', () => {
  fileInput.value = '';
  fileInput.click();
});

fileInput.addEventListener('change', () => {
  const files = fileInput.files;
  if (!files) {
    return;
  }

  for (const file of files) {
    void openEditor({ file, fileName: file.name, mode: getSelectedMode() }).catch(() => undefined);
  }
  fileInput.value = '';
});

async function closeAllEditors(): Promise<void> {
  await Promise.all([...records].map((record) => removeRecord(record)));
}

document.querySelector('#close-all-button')?.addEventListener('click', () => {
  void closeAllEditors();
});

(window as typeof window & { __officeDemo?: unknown }).__officeDemo = {
  bootId,
  records,
  openEmpty,
  setMode: setSelectedMode,
  closeAll: closeAllEditors,
};
