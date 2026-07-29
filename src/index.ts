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

const isChinese = navigator.language.toLocaleLowerCase().startsWith('zh');
const copy = isChinese
  ? {
      product: 'OnlyOffice 浏览器版',
      private: '文件只在此设备中处理',
      open: '打开文件',
      new: '新建',
      newWord: 'Word 文档',
      newSheet: '电子表格',
      newSlides: '演示文稿',
      noDocument: '打开或新建文档，开始在浏览器中编辑',
      noDocumentHint: '支持 DOCX、XLSX、PPTX、CSV 及常见旧版 Office 格式',
      save: '保存',
      close: '关闭',
      resources: 'Office 资源',
      resourcesReady: '资源已就绪',
      resourcesNeeded: '按需下载',
      resourcesUpdating: '正在更新资源',
      resourcesError: '资源需要修复',
      resourceIntro: '编辑器会按需下载当前文档所需资源，并在空闲时准备基础资源。',
      basicPreset: '准备基础资源',
      compatPreset: '安装 Office 兼容字体',
      repair: '检查并修复',
      allResources: '下载全部（高级）',
      advancedFonts: '高级字体管理',
      installed: '已安装',
      download: '下载',
      remove: '移除',
      required: '基础字体',
      updateReady: '新版本已准备好。保存或关闭未保存的文档后将自动更新。',
      updateNow: '立即更新',
      dirtyTitle: '保存更改？',
      dirtyBody: '切换或关闭前，是否保存当前文档的更改？',
      discard: '不保存',
      cancel: '取消',
      permission: '需要重新授权此文件后才能恢复标签页。',
      authorize: '重新打开',
      error: '文档打开失败',
      version: '版本',
      word: '文字',
      cell: '表格',
      slide: '演示',
      core: '基础组件',
      fonts: '字体',
    }
  : {
      product: 'OnlyOffice Browser',
      private: 'Files stay on this device',
      open: 'Open files',
      new: 'New',
      newWord: 'Word document',
      newSheet: 'Spreadsheet',
      newSlides: 'Presentation',
      noDocument: 'Open or create a document to start editing in your browser',
      noDocumentHint: 'Supports DOCX, XLSX, PPTX, CSV, and common legacy Office formats',
      save: 'Save',
      close: 'Close',
      resources: 'Office resources',
      resourcesReady: 'Resources ready',
      resourcesNeeded: 'Downloads on demand',
      resourcesUpdating: 'Updating resources',
      resourcesError: 'Resources need repair',
      resourceIntro: 'The editor downloads resources for the current document and prepares essentials when idle.',
      basicPreset: 'Prepare essentials',
      compatPreset: 'Install Office-compatible fonts',
      repair: 'Check and repair',
      allResources: 'Download everything (advanced)',
      advancedFonts: 'Advanced font management',
      installed: 'Installed',
      download: 'Download',
      remove: 'Remove',
      required: 'Essential',
      updateReady: 'A new version is ready. It will update automatically after unsaved documents are saved or closed.',
      updateNow: 'Update now',
      dirtyTitle: 'Save changes?',
      dirtyBody: 'Would you like to save changes before switching or closing this document?',
      discard: 'Discard',
      cancel: 'Cancel',
      permission: 'Authorize this file again to restore the tab.',
      authorize: 'Open again',
      error: 'Unable to open document',
      version: 'Version',
      word: 'Word',
      cell: 'Spreadsheet',
      slide: 'Presentation',
      core: 'Essentials',
      fonts: 'Fonts',
    };

const app = document.querySelector<HTMLDivElement>('#app');
if (!app) throw new Error('Missing #app root');

app.innerHTML = `
  <header class="app-header">
    <div class="brand">
      <img src="/onlyoffice-icon.svg" alt="" />
      <div><strong>${copy.product}</strong><span>${copy.private}</span></div>
    </div>
    <div class="header-actions">
      <button id="resource-button" class="status-button" type="button"><span class="status-dot"></span><span>${copy.resourcesNeeded}</span></button>
      <button id="open-button" class="primary-button" type="button">${copy.open}</button>
      <details class="new-menu">
        <summary>${copy.new}</summary>
        <button type="button" data-new="docx">${copy.newWord}</button>
        <button type="button" data-new="xlsx">${copy.newSheet}</button>
        <button type="button" data-new="pptx">${copy.newSlides}</button>
      </details>
    </div>
  </header>
  <aside id="update-banner" class="update-banner" hidden>
    <span>${copy.updateReady}</span><button id="update-button" type="button">${copy.updateNow}</button>
  </aside>
  <nav id="document-tabs" class="document-tabs" aria-label="Open documents"></nav>
  <main class="editor-workspace">
    <section id="empty-state" class="empty-state">
      <img src="/onlyoffice-icon.svg" alt="" />
      <h1>${copy.noDocument}</h1>
      <p>${copy.noDocumentHint}</p>
      <button id="empty-open-button" class="primary-button" type="button">${copy.open}</button>
    </section>
    <section id="permission-state" class="empty-state" hidden>
      <h1>${copy.permission}</h1>
      <button id="authorize-button" class="primary-button" type="button">${copy.authorize}</button>
    </section>
    <section id="editor-panel" class="editor-panel" hidden>
      <div class="document-bar">
        <div><strong id="document-title"></strong><span id="document-status"></span></div>
        <div><button id="save-button" type="button">${copy.save}</button><button id="close-button" type="button">${copy.close}</button></div>
      </div>
      <div id="editor-slot" class="editor-slot"></div>
    </section>
  </main>
  <footer class="app-footer"><span>${copy.product}</span><span>${copy.version} ${ONLYOFFICE_BROWSER_VERSION}</span></footer>
  <input id="file-input" type="file" multiple accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.csv,.rtf,.odt,.ods,.odp" />
  <dialog id="resource-dialog" class="settings-dialog">
    <form method="dialog">
      <header><div><h2>${copy.resources}</h2><p>${copy.resourceIntro}</p></div><button value="cancel" aria-label="Close" type="submit">×</button></header>
      <div id="resource-summary" class="resource-summary"></div>
      <div class="preset-actions">
        <button id="basic-preset" type="button">${copy.basicPreset}</button>
        <button id="compat-preset" class="primary-button" type="button">${copy.compatPreset}</button>
      </div>
      <details>
        <summary>${copy.advancedFonts}</summary>
        <div id="font-list" class="font-list"></div>
        <button id="load-all-button" type="button">${copy.allResources}</button>
      </details>
      <footer><button id="repair-button" type="button">${copy.repair}</button></footer>
    </form>
  </dialog>
  <dialog id="dirty-dialog" class="confirm-dialog">
    <form method="dialog">
      <h2>${copy.dirtyTitle}</h2><p>${copy.dirtyBody}</p>
      <div><button value="cancel">${copy.cancel}</button><button value="discard">${copy.discard}</button><button class="primary-button" value="save">${copy.save}</button></div>
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
  updateBanner: document.querySelector<HTMLElement>('#update-banner')!,
  updateButton: document.querySelector<HTMLButtonElement>('#update-button')!,
  dirtyDialog: document.querySelector<HTMLDialogElement>('#dirty-dialog')!,
};

const tabs: DocumentTab[] = [];
const tabStore = new DocumentTabStore();
let activeTab: DocumentTab | null = null;
let editor: OfficeEditorInstance | null = null;
let resourceManager: OfficeRuntimeResourceManager | null = null;
let waitingWorkbox: Workbox | null = null;
let editorGeneration = 0;
const hardResetOnLastDestroy = new URLSearchParams(location.search).get('hardResetOnLastDestroy') === 'true';
clearLegacyDemoHostState(location, localStorage);
const officeHostUrl = resolveDemoHostUrl(new URL(location.href));

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
          description: 'Office document',
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
  if (!(await requestReadWritePermission(handle))) throw new Error('Write permission was not granted.');
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
            description: 'Office documents',
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
    name: `New_Document.${emptyType}`,
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
  elements.resourceSummary.replaceChildren(
    ...snapshot.packs.map((pack) => {
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
    }),
  );
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

initializeServiceWorker();
void initializeResources();
void restoreTabs();
