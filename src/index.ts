import { createOfficeEditor, type OfficeEditorInstance } from './lib/office-editor';
import {
  createOfficeRuntimeResourceManager,
  type OfficeRuntimeResourceManager,
  type OfficeRuntimeResourceSnapshot,
} from './lib/runtime-resources';
import { clearLegacyDemoHostState, resolveDemoHostUrl } from './lib/demo-host-url';
import {
  createDocumentTabId,
  DocumentTabStore,
  documentTypeForName,
  hasReadPermission,
  requestReadWritePermission,
} from './pwa/document-tabs';
import {
  OFFICE_LOCALE_STORAGE_KEY,
  officeCopy,
  resolveOfficeLocale,
  type OfficeCopy,
  type OfficeLocale,
} from './pwa/i18n';
import { ONLYOFFICE_BROWSER_VERSION } from './version';
import { Workbox } from 'workbox-window';
import './styles/base.css';

declare global {
  interface Window {
    showOpenFilePicker?: (options?: {
      multiple?: boolean;
      types?: Array<{ description: string; accept: Record<string, string[]> }>;
    }) => Promise<FileSystemFileHandle[]>;
    showSaveFilePicker?: (options?: {
      suggestedName?: string;
      types?: Array<{ description: string; accept: Record<string, string[]> }>;
    }) => Promise<FileSystemFileHandle>;
    __officeDemo?: unknown;
  }
}

type DocumentTab = {
  id: string;
  name: string;
  handle?: FileSystemFileHandle;
  file?: File;
  emptyType?: 'docx' | 'xlsx' | 'pptx' | 'csv';
  dirty: boolean;
  inaccessible?: boolean;
};

type DirtyDecision = 'save' | 'discard' | 'cancel';

let locale: OfficeLocale = resolveOfficeLocale(localStorage.getItem(OFFICE_LOCALE_STORAGE_KEY), navigator.languages);
let copy: OfficeCopy = officeCopy[locale];
document.documentElement.lang = locale;

const app = document.querySelector<HTMLDivElement>('#app');
if (!app) throw new Error('Missing #app root');

app.innerHTML = `
  <header class="app-header">
    <div class="brand">
      <img src="/onlyoffice-icon.svg" alt="" />
      <strong data-i18n="product">${copy.product}</strong>
    </div>
    <div class="header-actions">
      <label class="language-control">
        <span class="visually-hidden" data-i18n="language">${copy.language}</span>
        <select id="language-select" aria-label="${copy.language}">
          <option value="zh-CN"${locale === 'zh-CN' ? ' selected' : ''}>中文</option>
          <option value="en-US"${locale === 'en-US' ? ' selected' : ''}>English</option>
        </select>
      </label>
      <button id="resource-button" class="status-button" type="button"><span class="status-dot"></span><span>${copy.resourcesNeeded}</span></button>
      <button id="open-button" class="primary-button" data-i18n="open" type="button">${copy.open}</button>
      <details class="new-menu">
        <summary data-i18n="new">${copy.new}</summary>
        <button type="button" data-new="docx" data-i18n="newWord">${copy.newWord}</button>
        <button type="button" data-new="xlsx" data-i18n="newSheet">${copy.newSheet}</button>
        <button type="button" data-new="pptx" data-i18n="newSlides">${copy.newSlides}</button>
      </details>
    </div>
  </header>
  <aside id="update-banner" class="update-banner" hidden>
    <span data-i18n="updateReady">${copy.updateReady}</span><button id="update-button" data-i18n="updateNow" type="button">${copy.updateNow}</button>
  </aside>
  <nav id="document-tabs" class="document-tabs" aria-label="${copy.openDocuments}"></nav>
  <main class="editor-workspace">
    <section id="empty-state" class="empty-state">
      <img src="/onlyoffice-icon.svg" alt="" />
      <h1 data-i18n="noDocument">${copy.noDocument}</h1>
      <p data-i18n="noDocumentHint">${copy.noDocumentHint}</p>
      <button id="empty-open-button" class="primary-button" data-i18n="open" type="button">${copy.open}</button>
    </section>
    <section id="permission-state" class="empty-state" hidden>
      <h1 data-i18n="permission">${copy.permission}</h1>
      <button id="authorize-button" class="primary-button" data-i18n="authorize" type="button">${copy.authorize}</button>
    </section>
    <section id="editor-panel" class="editor-panel" hidden>
      <div class="document-bar">
        <div><strong id="document-title"></strong><span id="document-status"></span></div>
        <div><button id="save-button" data-i18n="save" type="button">${copy.save}</button><button id="close-button" data-i18n="close" type="button">${copy.close}</button></div>
      </div>
      <div id="editor-slot" class="editor-slot"></div>
    </section>
  </main>
  <footer class="app-footer"><span data-i18n="product">${copy.product}</span><span><span data-i18n="version">${copy.version}</span> ${ONLYOFFICE_BROWSER_VERSION}</span></footer>
  <input id="file-input" type="file" multiple accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.csv,.rtf,.odt,.ods,.odp" />
  <dialog id="resource-dialog" class="settings-dialog">
    <form method="dialog">
      <header><div><h2 data-i18n="resources">${copy.resources}</h2><p data-i18n="resourceIntro">${copy.resourceIntro}</p></div><button value="cancel" aria-label="${copy.closeDialog}" type="submit">×</button></header>
      <div id="resource-summary" class="resource-summary"></div>
      <div class="preset-actions">
        <button id="basic-preset" data-i18n="basicPreset" type="button">${copy.basicPreset}</button>
        <button id="compat-preset" class="primary-button" data-i18n="compatPreset" type="button">${copy.compatPreset}</button>
      </div>
      <details>
        <summary data-i18n="advancedFonts">${copy.advancedFonts}</summary>
        <div id="font-list" class="font-list"></div>
        <button id="load-all-button" data-i18n="allResources" type="button">${copy.allResources}</button>
      </details>
      <footer><button id="repair-button" data-i18n="repair" type="button">${copy.repair}</button></footer>
    </form>
  </dialog>
  <dialog id="dirty-dialog" class="confirm-dialog">
    <form method="dialog">
      <h2 data-i18n="dirtyTitle">${copy.dirtyTitle}</h2><p data-i18n="dirtyBody">${copy.dirtyBody}</p>
      <div><button value="cancel" data-i18n="cancel">${copy.cancel}</button><button value="discard" data-i18n="discard">${copy.discard}</button><button class="primary-button" value="save" data-i18n="save">${copy.save}</button></div>
    </form>
  </dialog>
`;

const elements = {
  tabs: document.querySelector<HTMLElement>('#document-tabs')!,
  empty: document.querySelector<HTMLElement>('#empty-state')!,
  permission: document.querySelector<HTMLElement>('#permission-state')!,
  editor: document.querySelector<HTMLElement>('#editor-panel')!,
  slot: document.querySelector<HTMLElement>('#editor-slot')!,
  title: document.querySelector<HTMLElement>('#document-title')!,
  status: document.querySelector<HTMLElement>('#document-status')!,
  fileInput: document.querySelector<HTMLInputElement>('#file-input')!,
  resourceButton: document.querySelector<HTMLButtonElement>('#resource-button')!,
  resourceDialog: document.querySelector<HTMLDialogElement>('#resource-dialog')!,
  resourceSummary: document.querySelector<HTMLElement>('#resource-summary')!,
  fontList: document.querySelector<HTMLElement>('#font-list')!,
  basicPreset: document.querySelector<HTMLButtonElement>('#basic-preset')!,
  compatibilityPreset: document.querySelector<HTMLButtonElement>('#compat-preset')!,
  loadAllButton: document.querySelector<HTMLButtonElement>('#load-all-button')!,
  repairButton: document.querySelector<HTMLButtonElement>('#repair-button')!,
  updateBanner: document.querySelector<HTMLElement>('#update-banner')!,
  updateButton: document.querySelector<HTMLButtonElement>('#update-button')!,
  dirtyDialog: document.querySelector<HTMLDialogElement>('#dirty-dialog')!,
  languageSelect: document.querySelector<HTMLSelectElement>('#language-select')!,
};

const tabs: DocumentTab[] = [];
const tabStore = new DocumentTabStore();
let activeTab: DocumentTab | null = null;
let editor: OfficeEditorInstance | null = null;
let resourceManager: OfficeRuntimeResourceManager | null = null;
let latestResourceSnapshot: OfficeRuntimeResourceSnapshot | null = null;
let waitingWorkbox: Workbox | null = null;
let editorGeneration = 0;
const hardResetOnLastDestroy = new URLSearchParams(location.search).get('hardResetOnLastDestroy') === 'true';
clearLegacyDemoHostState(location, localStorage);
const officeHostUrl = resolveDemoHostUrl(new URL(location.href));

function applyLocale(nextLocale: OfficeLocale): void {
  locale = nextLocale;
  copy = officeCopy[locale];
  document.documentElement.lang = locale;
  document.title = copy.product;
  elements.languageSelect.value = locale;
  elements.languageSelect.setAttribute('aria-label', copy.language);
  elements.tabs.setAttribute('aria-label', copy.openDocuments);
  document
    .querySelector<HTMLButtonElement>('.settings-dialog header > button')
    ?.setAttribute('aria-label', copy.closeDialog);
  document.querySelectorAll<HTMLElement>('[data-i18n]').forEach((element) => {
    const key = element.dataset.i18n as keyof OfficeCopy | undefined;
    if (key) element.textContent = copy[key];
  });
  try {
    localStorage.setItem(OFFICE_LOCALE_STORAGE_KEY, locale);
  } catch {
    // Language switching remains live when storage is unavailable.
  }
  renderTabs();
  if (latestResourceSnapshot) {
    renderResources(latestResourceSnapshot);
  } else if (elements.resourceButton.dataset.state === 'error') {
    elements.resourceButton.lastElementChild!.textContent = copy.resourcesError;
  }
}

function formatBytes(bytes: number): string {
  if (bytes < 1024 ** 2) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / 1024 ** 2).toFixed(bytes >= 100 * 1024 ** 2 ? 0 : 1)} MB`;
}

function hasUnsafeWork(): boolean {
  return tabs.some((tab) => tab.dirty || !tab.handle);
}

function renderTabs(): void {
  elements.tabs.replaceChildren(
    ...tabs.map((tab) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = `document-tab${tab === activeTab ? ' active' : ''}`;
      button.title = tab.name;
      const label = document.createElement('span');
      label.textContent = tab.name;
      const dirty = document.createElement('i');
      dirty.textContent = tab.dirty ? '●' : '';
      const close = document.createElement('b');
      close.textContent = '×';
      close.title = copy.close;
      close.addEventListener('click', (event) => {
        event.stopPropagation();
        void closeTab(tab);
      });
      button.append(label, dirty, close);
      button.addEventListener('click', () => void activateTab(tab));
      return button;
    }),
  );
}

function setView(view: 'empty' | 'permission' | 'editor'): void {
  elements.empty.hidden = view !== 'empty';
  elements.permission.hidden = view !== 'permission';
  elements.editor.hidden = view !== 'editor';
}

function setDocumentStatus(text: string): void {
  elements.status.textContent = text;
}

async function persistTab(tab: DocumentTab): Promise<void> {
  if (!tab.handle) return;
  await tabStore.put({ id: tab.id, name: tab.name, handle: tab.handle, lastOpenedAt: Date.now() });
}

function downloadFallback(file: File): void {
  const url = URL.createObjectURL(file);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = file.name;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

async function saveFile(tab: DocumentTab, file: File): Promise<boolean> {
  let handle = tab.handle;
  if (!handle && window.showSaveFilePicker) {
    handle = await window.showSaveFilePicker({
      suggestedName: file.name || tab.name,
      types: [
        {
          description: copy.officeDocument,
          accept: { 'application/octet-stream': [`.${file.name.split('.').pop() || 'docx'}`] },
        },
      ],
    });
  }
  if (!handle) {
    downloadFallback(file);
    tab.dirty = false;
    renderTabs();
    return true;
  }
  if (!(await requestReadWritePermission(handle))) throw new Error(copy.writePermissionError);
  const writable = await handle.createWritable();
  await writable.write(file);
  await writable.close();
  tab.handle = handle;
  tab.name = handle.name;
  tab.dirty = false;
  await persistTab(tab);
  renderTabs();
  tryActivateUpdate();
  return true;
}

function scheduleIdle(task: () => Promise<unknown>): void {
  const run = () => void task().catch(() => undefined);
  const requestIdle = (
    window as Window & {
      requestIdleCallback?: (callback: () => void, options: { timeout: number }) => number;
    }
  ).requestIdleCallback;
  if (requestIdle) {
    requestIdle(run, { timeout: 5_000 });
  } else {
    globalThis.setTimeout(run, 1_500);
  }
}

async function destroyEditor(): Promise<void> {
  editorGeneration += 1;
  const current = editor;
  editor = null;
  if (current) await current.destroy();
  elements.slot.replaceChildren();
}

async function activateTab(tab: DocumentTab): Promise<void> {
  if (tab === activeTab && editor) return;
  if (activeTab?.dirty) {
    const decision = await askDirtyDecision();
    if (decision === 'cancel') return;
    if (decision === 'save' && editor) {
      await editor.save();
      if (activeTab.dirty) return;
    }
  }
  await destroyEditor();
  activeTab = tab;
  renderTabs();
  elements.title.textContent = tab.name;
  if (tab.inaccessible || (tab.handle && !(await hasReadPermission(tab.handle)))) {
    tab.inaccessible = true;
    setView('permission');
    return;
  }
  setView('editor');
  setDocumentStatus('…');
  const generation = ++editorGeneration;
  try {
    const file = tab.file || (tab.handle ? await tab.handle.getFile() : undefined);
    tab.file = undefined;
    const instance = await createOfficeEditor(elements.slot, {
      hostUrl: officeHostUrl,
      file,
      emptyType: tab.emptyType,
      fileName: tab.name,
      mode: 'edit',
      saveBehavior: 'callback',
      downloadedFonts: resourceManager?.getVerifiedFontPaths() || [],
      hardResetOnLastDestroy,
      onReady: (ready) => {
        if (generation !== editorGeneration) return;
        setDocumentStatus(ready.getState().fileType.toUpperCase());
        scheduleIdle(() => resourceManager?.prepareForDocumentType(documentTypeForName(tab.name)) || Promise.resolve());
      },
      onSave: (saved) => saveFile(tab, saved),
      onDirtyChange: (dirty) => {
        if (generation !== editorGeneration) return;
        tab.dirty = dirty;
        setDocumentStatus(dirty ? '●' : editor?.getState().fileType.toUpperCase() || '');
        renderTabs();
        if (!dirty) tryActivateUpdate();
      },
      onError: (error) => {
        if (generation === editorGeneration) setDocumentStatus(`${copy.error}: ${error.message}`);
      },
    });
    if (generation !== editorGeneration) {
      await instance.destroy();
      return;
    }
    editor = instance;
    await persistTab(tab);
  } catch (error) {
    if (generation === editorGeneration) {
      setDocumentStatus(`${copy.error}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}

function askDirtyDecision(): Promise<DirtyDecision> {
  return new Promise((resolve) => {
    const onClose = () => {
      elements.dirtyDialog.removeEventListener('close', onClose);
      const value = elements.dirtyDialog.returnValue;
      resolve(value === 'save' || value === 'discard' ? value : 'cancel');
    };
    elements.dirtyDialog.addEventListener('close', onClose);
    elements.dirtyDialog.showModal();
  });
}

async function closeTab(tab: DocumentTab): Promise<void> {
  if (tab === activeTab && tab.dirty) {
    const decision = await askDirtyDecision();
    if (decision === 'cancel') return;
    if (decision === 'save' && editor) {
      await editor.save();
      if (tab.dirty) return;
    }
  }
  const index = tabs.indexOf(tab);
  if (index < 0) return;
  if (tab === activeTab) {
    await destroyEditor();
    activeTab = null;
  }
  tabs.splice(index, 1);
  await tabStore.remove(tab.id).catch(() => undefined);
  renderTabs();
  const next = tabs[Math.min(index, tabs.length - 1)];
  if (next) await activateTab(next);
  else setView('empty');
  tryActivateUpdate();
}

async function addFile(handle: FileSystemFileHandle | undefined, file: File): Promise<void> {
  if (handle) {
    for (const tab of tabs) {
      if (tab.handle && (await tab.handle.isSameEntry(handle))) {
        await activateTab(tab);
        return;
      }
    }
  }
  const tab: DocumentTab = { id: createDocumentTabId(), name: file.name, handle, file, dirty: false };
  tabs.push(tab);
  await persistTab(tab);
  renderTabs();
  await activateTab(tab);
}

async function openFiles(): Promise<void> {
  if (window.showOpenFilePicker) {
    try {
      const handles = await window.showOpenFilePicker({
        multiple: true,
        types: [
          {
            description: copy.officeDocuments,
            accept: {
              'application/octet-stream': [
                '.docx',
                '.xlsx',
                '.pptx',
                '.doc',
                '.xls',
                '.ppt',
                '.csv',
                '.rtf',
                '.odt',
                '.ods',
                '.odp',
              ],
            },
          },
        ],
      });
      for (const handle of handles) await addFile(handle, await handle.getFile());
      return;
    } catch (error) {
      if (error instanceof DOMException && error.name === 'AbortError') return;
    }
  }
  elements.fileInput.value = '';
  elements.fileInput.click();
}

function createEmpty(emptyType: 'docx' | 'xlsx' | 'pptx' | 'csv'): void {
  const tab: DocumentTab = {
    id: createDocumentTabId(),
    name: `${copy.newDocumentName}.${emptyType}`,
    emptyType,
    dirty: false,
  };
  tabs.push(tab);
  renderTabs();
  void activateTab(tab);
}

function packLabel(id: string): string {
  return copy[id as 'word' | 'cell' | 'slide' | 'core' | 'fonts'] || id;
}

function renderResources(snapshot: OfficeRuntimeResourceSnapshot): void {
  latestResourceSnapshot = snapshot;
  const label =
    snapshot.readiness === 'ready'
      ? copy.resourcesReady
      : snapshot.readiness === 'updating'
        ? copy.resourcesUpdating
        : snapshot.readiness === 'error'
          ? copy.resourcesError
          : copy.resourcesNeeded;
  elements.resourceButton.lastElementChild!.textContent = label;
  elements.resourceButton.dataset.state = snapshot.readiness;
  const summaryRows = snapshot.packs.map((pack) => {
    const row = document.createElement('div');
    row.className = 'resource-pack';
    const text = document.createElement('span');
    text.textContent = packLabel(pack.id);
    const state = document.createElement('span');
    state.textContent = pack.ready
      ? copy.installed
      : `${formatBytes(pack.completedBytes)} / ${formatBytes(pack.totalBytes)}`;
    row.append(text, state);
    return row;
  });
  const failures = snapshot.progress.failures || [];
  if (failures.length > 0) {
    const failure = document.createElement('div');
    failure.className = 'resource-failure';
    const paths = failures.map(({ path }) => path).join(', ');
    const title = document.createElement('strong');
    title.textContent = `${copy.failedResources}: ${paths}`;
    const hint = document.createElement('span');
    hint.textContent = copy.repairFailedHint;
    failure.append(title, hint);
    summaryRows.push(failure);
  }
  elements.resourceSummary.replaceChildren(...summaryRows);
  const busy = Boolean(snapshot.operation);
  elements.basicPreset.disabled = busy;
  elements.compatibilityPreset.disabled = busy;
  elements.loadAllButton.disabled = busy;
  elements.repairButton.disabled = busy;
  elements.fontList.replaceChildren(
    ...snapshot.fonts.map((font) => {
      const row = document.createElement('div');
      row.className = 'font-row';
      const label = document.createElement('span');
      label.textContent = `${font.name} · ${formatBytes(font.bytes)}`;
      const button = document.createElement('button');
      button.type = 'button';
      button.textContent = font.removable ? (font.downloaded ? copy.remove : copy.download) : copy.required;
      button.disabled = !font.removable || Boolean(snapshot.operation);
      button.addEventListener('click', () => {
        if (!resourceManager) return;
        void (font.downloaded
          ? resourceManager.uninstallFontFamily(font.id)
          : resourceManager.downloadFontFamily(font.id));
      });
      row.append(label, button);
      return row;
    }),
  );
}

async function initializeResources(): Promise<void> {
  try {
    resourceManager = await createOfficeRuntimeResourceManager();
    resourceManager.subscribe(renderResources);
    renderResources(resourceManager.getSnapshot());
    scheduleIdle(() => resourceManager?.prefetchRecommended() || Promise.resolve());
  } catch {
    elements.resourceButton.lastElementChild!.textContent = copy.resourcesError;
    elements.resourceButton.dataset.state = 'error';
  }
}

function tryActivateUpdate(): void {
  if (!waitingWorkbox || hasUnsafeWork()) return;
  elements.updateButton.disabled = true;
  void waitingWorkbox.messageSkipWaiting();
}

function initializeServiceWorker(): void {
  if (!('serviceWorker' in navigator)) return;
  const workbox = new Workbox('/sw.js', { scope: '/' });
  let reloading = false;
  workbox.addEventListener('waiting', () => {
    waitingWorkbox = workbox;
    elements.updateBanner.hidden = false;
    tryActivateUpdate();
  });
  workbox.addEventListener('controlling', () => {
    if (reloading || hasUnsafeWork()) return;
    reloading = true;
    location.reload();
  });
  void workbox.register().then(() => {
    const interval = window.setInterval(() => void workbox.update(), 60 * 60 * 1000);
    addEventListener('pagehide', () => clearInterval(interval), { once: true });
  });
}

async function restoreTabs(): Promise<void> {
  if (!('indexedDB' in window)) return;
  try {
    const persisted = await tabStore.list();
    for (const record of persisted) {
      tabs.push({
        id: record.id,
        name: record.name,
        handle: record.handle,
        dirty: false,
        inaccessible: !(await hasReadPermission(record.handle)),
      });
    }
    renderTabs();
    if (tabs.length > 0) await activateTab(tabs[tabs.length - 1]);
  } catch {
    // Private browsing and policy-restricted contexts may not offer IndexedDB.
  }
}

document.querySelector('#open-button')?.addEventListener('click', () => void openFiles());
elements.languageSelect.addEventListener('change', () => {
  applyLocale(elements.languageSelect.value as OfficeLocale);
});
document.querySelector('#empty-open-button')?.addEventListener('click', () => void openFiles());
document.querySelector('#save-button')?.addEventListener('click', () => void editor?.save());
document.querySelector('#close-button')?.addEventListener('click', () => activeTab && void closeTab(activeTab));
document.querySelectorAll<HTMLElement>('[data-new]').forEach((button) => {
  button.addEventListener('click', () => createEmpty(button.dataset.new as 'docx' | 'xlsx' | 'pptx' | 'csv'));
});
elements.fileInput.addEventListener('change', () => {
  for (const file of Array.from(elements.fileInput.files || [])) void addFile(undefined, file);
});
elements.resourceButton.addEventListener('click', () => elements.resourceDialog.showModal());
elements.updateButton.addEventListener('click', tryActivateUpdate);
document
  .querySelector('#basic-preset')
  ?.addEventListener('click', () => void resourceManager?.installFontPreset('basic'));
document
  .querySelector('#compat-preset')
  ?.addEventListener('click', () => void resourceManager?.installFontPreset('office-compatibility'));
document.querySelector('#repair-button')?.addEventListener('click', () => void resourceManager?.repair());
document.querySelector('#load-all-button')?.addEventListener('click', () => void resourceManager?.loadAll());
document.querySelector('#authorize-button')?.addEventListener('click', () => {
  if (!activeTab?.handle) return;
  void requestReadWritePermission(activeTab.handle).then((granted) => {
    if (!granted || !activeTab) return;
    activeTab.inaccessible = false;
    void activateTab(activeTab);
  });
});
addEventListener('beforeunload', (event) => {
  if (!tabs.some((tab) => tab.dirty)) return;
  event.preventDefault();
});

window.__officeDemo = {
  get tabs() {
    return tabs;
  },
  openEmpty: createEmpty,
  closeAll: async () => {
    for (const tab of tabs.slice()) await closeTab(tab);
  },
};

applyLocale(locale);
initializeServiceWorker();
void initializeResources();
void restoreTabs();
