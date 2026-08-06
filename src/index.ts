import {
  mountOfficeEditor,
  type OfficeEditorInstance,
  type OfficeEditorMount,
  type OfficeHostUrlContext,
} from './lib/office-editor';
import {
  isOfficeEditorOriginSlot,
  OFFICE_EDITOR_ORIGIN_SLOTS,
  type OfficeEditorOriginSlot,
} from './lib/office-origin-pool';
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
import { createElement } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { OfficeResourcePanel } from './pwa/resource-panel';
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
  originSlot?: OfficeEditorOriginSlot;
  dirty: boolean;
  inaccessible?: boolean;
  editor?: DocumentEditorRecord;
};

type DocumentEditorRecord = {
  container: HTMLDivElement;
  mount: OfficeEditorMount;
  instance: OfficeEditorInstance | null;
  origin: string;
  status: 'waiting' | 'activating' | 'ready' | 'error' | 'disposed';
  sequence: number;
  foreground: boolean;
  queued: boolean;
  disposed: boolean;
  settled: boolean;
  ready: Promise<OfficeEditorInstance>;
  resolveReady: (instance: OfficeEditorInstance) => void;
  rejectReady: (error: Error) => void;
  error?: Error;
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
  <main class="editor-workspace">
    <section class="preview-pane" data-testid="office-preview-pane">
      <div class="preview-tabbar" data-testid="office-preview-tabbar">
        <nav id="document-tabs" class="document-tabs piwork-scrollbar-hidden" aria-label="${copy.openDocuments}"></nav>
      </div>
      <div class="preview-body" data-testid="office-preview-body">
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
          <div class="document-bar" data-testid="office-preview-toolbar">
            <div class="document-heading">
              <span id="document-type-icon" class="document-type-icon" aria-hidden="true">W</span>
              <strong id="document-title"></strong>
              <span id="document-origin" class="document-origin"></span>
              <span id="document-status" class="document-status"></span>
            </div>
            <div class="document-actions">
              <button id="save-button" class="preview-toolbar-button" aria-label="${copy.save}" title="${copy.save}" type="button">
                <svg class="toolbar-glyph" viewBox="0 0 24 24" aria-hidden="true"><path d="M5 3h11l3 3v15H5zM8 3v6h8V3M8 21v-7h8v7" /></svg>
                <span class="toolbar-button-label" data-i18n="save">${copy.save}</span>
              </button>
              <button id="close-button" class="preview-toolbar-button" aria-label="${copy.close}" title="${copy.close}" type="button">
                <svg class="toolbar-glyph" viewBox="0 0 24 24" aria-hidden="true"><path d="m6 6 12 12M18 6 6 18" /></svg>
                <span class="toolbar-button-label" data-i18n="close">${copy.close}</span>
              </button>
            </div>
          </div>
          <div id="editor-slot" class="editor-slot" aria-live="polite"></div>
        </section>
      </div>
    </section>
  </main>
  <footer class="app-footer"><span data-i18n="product">${copy.product}</span><span id="version-label"><span data-i18n="version">${copy.version}</span> ${ONLYOFFICE_BROWSER_VERSION}</span></footer>
  <input id="file-input" type="file" multiple accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.csv,.rtf,.odt,.ods,.odp" />
  <dialog id="resource-dialog" class="settings-dialog">
    <form method="dialog">
      <header><div><h2 data-i18n="resources">${copy.resources}</h2><p data-i18n="resourceIntro">${copy.resourceIntro}</p></div><button value="cancel" aria-label="${copy.closeDialog}" type="submit">×</button></header>
      <div id="resource-react-root"></div>
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
  origin: document.querySelector<HTMLElement>('#document-origin')!,
  typeIcon: document.querySelector<HTMLElement>('#document-type-icon')!,
  status: document.querySelector<HTMLElement>('#document-status')!,
  fileInput: document.querySelector<HTMLInputElement>('#file-input')!,
  resourceButton: document.querySelector<HTMLButtonElement>('#resource-button')!,
  resourceDialog: document.querySelector<HTMLDialogElement>('#resource-dialog')!,
  updateBanner: document.querySelector<HTMLElement>('#update-banner')!,
  updateButton: document.querySelector<HTMLButtonElement>('#update-button')!,
  dirtyDialog: document.querySelector<HTMLDialogElement>('#dirty-dialog')!,
  languageSelect: document.querySelector<HTMLSelectElement>('#language-select')!,
  versionLabel: document.querySelector<HTMLElement>('#version-label')!,
};

const tabs: DocumentTab[] = [];
const tabStore = new DocumentTabStore();
let activeTab: DocumentTab | null = null;
let resourceManager: OfficeRuntimeResourceManager | null = null;
let resourcePanelRoot: Root | null = null;
let latestResourceSnapshot: OfficeRuntimeResourceSnapshot | null = null;
let resourceInitialization: Promise<void> | null = null;
let stopResourceMaintenance: (() => void) | null = null;
let pendingActivation: DocumentTab | null = null;
let waitingWorkbox: Workbox | null = null;
let updateActivationPending = false;
let tabSwitchGeneration = 0;
let nextEditorSequence = 1;
let activeEditorActivations = 0;
let activationDrainScheduled = false;
const pendingEditorActivations: DocumentEditorRecord[] = [];
const MAX_OPEN_DOCUMENTS = OFFICE_EDITOR_ORIGIN_SLOTS.length;
const hardResetOnLastDestroy = new URLSearchParams(location.search).get('hardResetOnLastDestroy') === 'true';
clearLegacyDemoHostState(location, localStorage);
const resolvedDemoHost = resolveDemoHostUrl(new URL(location.href));
const officeHostUrl = (context: OfficeHostUrlContext) => {
  const base = typeof resolvedDemoHost === 'function' ? resolvedDemoHost(context) : resolvedDemoHost;
  const resolved = new URL(base, location.href);
  const resourceSnapshot = resourceManager?.getSnapshot();
  const releaseId = resourceSnapshot?.installedRelease || resourceSnapshot?.targetRelease;
  if (
    releaseId &&
    (resolved.hostname.endsWith('.getpi.work') ||
      resolved.hostname.endsWith('.office.localhost') ||
      resolved.hostname === 'host.localhost')
  ) {
    resolved.pathname = `/r/${encodeURIComponent(releaseId)}/office-host.html`;
  }
  return resolved.href;
};

function applyLocale(nextLocale: OfficeLocale): void {
  locale = nextLocale;
  copy = officeCopy[locale];
  document.documentElement.lang = locale;
  document.title = copy.product;
  elements.languageSelect.value = locale;
  elements.languageSelect.setAttribute('aria-label', copy.language);
  elements.tabs.setAttribute('aria-label', copy.openDocuments);
  for (const id of ['save-button', 'close-button']) {
    const button = document.querySelector<HTMLButtonElement>(`#${id}`);
    const key = id === 'save-button' ? copy.save : copy.close;
    button?.setAttribute('aria-label', key);
    button?.setAttribute('title', key);
  }
  document
    .querySelector<HTMLButtonElement>('.settings-dialog header > button')
    ?.setAttribute('aria-label', copy.closeDialog);
  document.querySelectorAll<HTMLElement>('[data-i18n]').forEach((element) => {
    const key = element.dataset.i18n as keyof OfficeCopy | undefined;
    if (key && typeof copy[key] === 'string') element.textContent = copy[key];
  });
  try {
    localStorage.setItem(OFFICE_LOCALE_STORAGE_KEY, locale);
  } catch {
    // Language switching remains live when storage is unavailable.
  }
  renderTabs();
  if (activeTab) setActiveDocumentHeader(activeTab);
  if (latestResourceSnapshot) {
    renderResources(latestResourceSnapshot);
  } else if (elements.resourceButton.dataset.state === 'error') {
    elements.resourceButton.lastElementChild!.textContent = copy.resourcesError;
  }
}

function hasUnsafeWork(): boolean {
  const resourcePhase = resourceManager?.getSnapshot().phase;
  return (
    tabs.some((tab) => tab.dirty || !tab.handle) ||
    Boolean(resourcePhase && resourcePhase !== 'idle' && resourcePhase !== 'paused')
  );
}

const updateChannel = typeof BroadcastChannel === 'undefined' ? null : new BroadcastChannel('onlyoffice-pwa-update-v1');
const updateTabId = crypto.randomUUID();
const updatePeers = new Map<string, number>();
const updateResponses = new Map<string, Map<string, boolean>>();

function announceUpdatePresence(): void {
  updateChannel?.postMessage({ type: 'PRESENCE', tabId: updateTabId, protocol: 1 });
}

updateChannel?.addEventListener('message', (event) => {
  const message = event.data;
  if (!message || message.tabId === updateTabId) return;
  if (message.type === 'PRESENCE' && message.protocol === 1) {
    updatePeers.set(message.tabId, Date.now());
  } else if (message.type === 'PREPARE_UPDATE' && message.protocol === 1) {
    updateChannel.postMessage({
      type: 'UPDATE_STATUS',
      protocol: 1,
      requestId: message.requestId,
      tabId: updateTabId,
      unsafe: hasUnsafeWork(),
    });
  } else if (message.type === 'UPDATE_STATUS' && message.protocol === 1) {
    updateResponses.get(message.requestId)?.set(message.tabId, Boolean(message.unsafe));
  }
});
announceUpdatePresence();
const updatePresenceInterval = window.setInterval(announceUpdatePresence, 5_000);
addEventListener(
  'pagehide',
  () => {
    clearInterval(updatePresenceInterval);
    updateChannel?.close();
  },
  { once: true },
);

async function allTabsSafeForUpdate(): Promise<boolean> {
  if (hasUnsafeWork()) return false;
  if (!updateChannel) return true;
  const requestId = crypto.randomUUID();
  const responses = new Map<string, boolean>();
  updateResponses.set(requestId, responses);
  updateChannel.postMessage({ type: 'PREPARE_UPDATE', protocol: 1, requestId, tabId: updateTabId });
  await new Promise((resolve) => setTimeout(resolve, 400));
  updateResponses.delete(requestId);
  const livePeers = [...updatePeers.entries()]
    .filter(([, seenAt]) => Date.now() - seenAt < 15_000)
    .map(([tabId]) => tabId);
  return livePeers.every((tabId) => responses.get(tabId) === false);
}

const OFFICE_ORIGIN_SYMBOLS: Record<string, string> = {
  aries: '♈︎',
  taurus: '♉︎',
  gemini: '♊︎',
  cancer: '♋︎',
  leo: '♌︎',
  virgo: '♍︎',
  libra: '♎︎',
  scorpio: '♏︎',
  sagittarius: '♐︎',
  capricorn: '♑︎',
  aquarius: '♒︎',
  pisces: '♓︎',
};

function originSlot(origin: string | undefined): OfficeEditorOriginSlot | null {
  if (!origin) return null;
  try {
    const hostname = new URL(origin).hostname.toLowerCase();
    const slot = hostname.endsWith('.getpi.work')
      ? hostname.slice(0, -'.getpi.work'.length)
      : /^host-([^.]+)\.office\.localhost$/.exec(hostname)?.[1] || /^([^.]+)\.localhost$/.exec(hostname)?.[1];
    return slot && isOfficeEditorOriginSlot(slot) ? slot : null;
  } catch {
    return null;
  }
}

function editorOriginLabel(origin: string | undefined): string {
  const slot = originSlot(origin);
  if (!origin || !slot) return copy.originPending;
  return `${OFFICE_ORIGIN_SYMBOLS[slot]} ${slot}.getpi.work`;
}

function documentTypeSymbol(name: string): 'W' | 'X' | 'P' {
  const type = documentTypeForName(name);
  return type === 'cell' ? 'X' : type === 'slide' ? 'P' : 'W';
}

function isEditorPoolExhausted(error: unknown): boolean {
  return error instanceof Error && error.name === 'OfficeHostPoolExhaustedError';
}

function activeEditor(): OfficeEditorInstance | null {
  return activeTab?.editor?.instance || null;
}

function setEditorVisibility(tab: DocumentTab): void {
  for (const candidate of tabs) {
    const container = candidate.editor?.container;
    if (!container) continue;
    const active = candidate === tab;
    container.dataset.active = active ? 'true' : 'false';
    container.setAttribute('aria-hidden', String(!active));
    if (active) container.removeAttribute('inert');
    else container.setAttribute('inert', '');
  }
}

function renderTabs(): void {
  elements.tabs.replaceChildren(
    ...tabs.map((tab) => {
      const surface = document.createElement('div');
      surface.className = `document-tab${tab === activeTab ? ' active' : ''}`;
      surface.dataset.editorOrigin = tab.editor?.origin || '';
      surface.dataset.editorState = tab.editor?.status || (tab.inaccessible ? 'permission' : 'waiting');
      surface.dataset.documentKind = documentTypeForName(tab.name);
      const tabTitle = `${tab.name} · ${copy.editorOrigin}: ${editorOriginLabel(tab.editor?.origin)}`;
      const select = document.createElement('button');
      select.type = 'button';
      select.className = 'document-tab-select';
      select.title = tabTitle;
      select.setAttribute('aria-label', tabTitle);
      select.addEventListener('click', () => void activateTab(tab));
      const typeIcon = document.createElement('span');
      typeIcon.className = 'document-tab-icon';
      typeIcon.textContent = documentTypeSymbol(tab.name);
      typeIcon.setAttribute('aria-hidden', 'true');
      const label = document.createElement('span');
      label.className = 'document-tab-title';
      label.textContent = tab.name;
      const origin = document.createElement('span');
      origin.className = 'document-tab-origin';
      origin.textContent = originSlot(tab.editor?.origin)
        ? OFFICE_ORIGIN_SYMBOLS[originSlot(tab.editor?.origin)!]
        : '…';
      origin.setAttribute('aria-hidden', 'true');
      origin.title = editorOriginLabel(tab.editor?.origin);
      const dirty = document.createElement('span');
      dirty.className = 'document-tab-dirty';
      dirty.hidden = !tab.dirty;
      dirty.setAttribute('aria-hidden', 'true');
      const close = document.createElement('button');
      close.type = 'button';
      close.className = 'document-tab-close';
      close.setAttribute('aria-label', `${copy.close}: ${tab.name}`);
      close.title = copy.close;
      close.innerHTML = '<span aria-hidden="true">×</span>';
      close.addEventListener('click', (event) => {
        event.stopPropagation();
        void closeTab(tab);
      });
      surface.append(select, typeIcon, label, origin, close, dirty);
      return surface;
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
  const state =
    text === '…'
      ? 'waiting'
      : text === '●'
        ? 'dirty'
        : text === copy.openLimitReached || text.startsWith(`${copy.error}:`)
          ? 'error'
          : 'ready';
  elements.status.dataset.state = state;
}

function setActiveDocumentHeader(tab: DocumentTab): void {
  elements.typeIcon.textContent = documentTypeSymbol(tab.name);
  elements.typeIcon.dataset.kind = documentTypeForName(tab.name);
  elements.title.textContent = tab.name;
  elements.origin.textContent = editorOriginLabel(tab.editor?.origin);
  elements.origin.title = `${copy.editorOrigin}: ${editorOriginLabel(tab.editor?.origin)}`;
  elements.origin.dataset.state = tab.editor?.status || 'waiting';
}

async function persistTab(tab: DocumentTab): Promise<void> {
  if (!tab.handle) return;
  await tabStore.put({
    id: tab.id,
    name: tab.name,
    handle: tab.handle,
    lastOpenedAt: Date.now(),
    ...(tab.originSlot ? { originSlot: tab.originSlot } : {}),
  });
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
  void tryActivateUpdate();
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

function canMaintainResourcesInBackground(): boolean {
  if (navigator.onLine === false || document.visibilityState === 'hidden') return false;
  const connection = (navigator as Navigator & { connection?: { saveData?: boolean } }).connection;
  return connection?.saveData !== true;
}

function warmBackgroundEditors(): void {
  if (latestResourceSnapshot?.readiness !== 'ready' || !activeTab || !canMaintainResourcesInBackground()) return;
  for (const tab of tabs) {
    if (tab === activeTab || tab.inaccessible || tab.editor) continue;
    scheduleIdle(async () => {
      if (!tabs.includes(tab) || tab === activeTab || tab.inaccessible || tab.editor) return;
      await ensureEditor(tab, false);
    });
  }
}

function startResourceMaintenance(): void {
  stopResourceMaintenance?.();
  const maintain = () => {
    if (!canMaintainResourcesInBackground()) return;
    scheduleIdle(() => resourceManager?.maintain() || Promise.resolve());
  };
  const onVisible = () => {
    if (document.visibilityState === 'visible') maintain();
  };
  const interval = window.setInterval(maintain, 30 * 60 * 1000);
  const stop = () => {
    clearInterval(interval);
    removeEventListener('online', maintain);
    document.removeEventListener('visibilitychange', onVisible);
    removeEventListener('pagehide', stop);
    if (stopResourceMaintenance === stop) stopResourceMaintenance = null;
  };
  stopResourceMaintenance = stop;
  addEventListener('online', maintain);
  document.addEventListener('visibilitychange', onVisible);
  addEventListener('pagehide', stop, { once: true });
  maintain();
}

async function ensureResourcesReady(tab?: DocumentTab): Promise<boolean> {
  await resourceInitialization?.catch(() => undefined);
  if (latestResourceSnapshot?.readiness === 'ready') return true;
  if (tab) pendingActivation = tab;
  if (!elements.resourceDialog.open) elements.resourceDialog.showModal();
  return false;
}

function scheduleActivationDrain(): void {
  if (activationDrainScheduled) return;
  activationDrainScheduled = true;
  queueMicrotask(() => {
    activationDrainScheduled = false;
    pendingEditorActivations.sort(
      (left, right) => Number(right.foreground) - Number(left.foreground) || left.sequence - right.sequence,
    );
    while (activeEditorActivations < 1 && pendingEditorActivations.length > 0) {
      const record = pendingEditorActivations.shift()!;
      record.queued = false;
      if (record.disposed) continue;
      record.status = 'activating';
      activeEditorActivations += 1;
      renderTabs();
      void activateEditorRecord(record).finally(() => {
        activeEditorActivations -= 1;
        scheduleActivationDrain();
      });
    }
  });
}

function enqueueEditorActivation(record: DocumentEditorRecord, foreground: boolean): void {
  record.foreground ||= foreground;
  if (record.status !== 'waiting' || record.queued || record.disposed) return;
  record.queued = true;
  pendingEditorActivations.push(record);
  scheduleActivationDrain();
}

async function activateEditorRecord(record: DocumentEditorRecord): Promise<void> {
  try {
    const instance = await record.mount.activate();
    if (record.disposed) {
      await instance.destroy().catch(() => undefined);
      return;
    }
    record.instance = instance;
    record.status = 'ready';
    if (!record.settled) {
      record.settled = true;
      record.resolveReady(instance);
    }
    renderTabs();
    if (activeTab?.editor === record) {
      setActiveDocumentHeader(activeTab);
      setDocumentStatus(instance.getState().fileType.toUpperCase());
    }
  } catch (error) {
    const normalized = error instanceof Error ? error : new Error(String(error));
    record.error = normalized;
    record.status = 'error';
    await record.mount.destroy().catch(() => undefined);
    if (!record.settled) {
      record.settled = true;
      record.rejectReady(normalized);
    }
    if (activeTab?.editor === record) {
      setActiveDocumentHeader(activeTab);
      setDocumentStatus(
        isEditorPoolExhausted(normalized) ? copy.openLimitReached : `${copy.error}: ${normalized.message}`,
      );
    }
    renderTabs();
  }
}

async function disposeEditorRecord(tab: DocumentTab): Promise<void> {
  const record = tab.editor;
  if (!record) return;
  record.disposed = true;
  record.status = 'disposed';
  const queueIndex = pendingEditorActivations.indexOf(record);
  if (queueIndex >= 0) pendingEditorActivations.splice(queueIndex, 1);
  if (!record.settled) {
    record.settled = true;
    record.rejectReady(new DOMException('Office editor was closed before it became ready', 'AbortError'));
  }
  await record.mount.destroy().catch(() => undefined);
  record.container.remove();
  if (tab.editor === record) delete tab.editor;
}

async function createEditorRecord(tab: DocumentTab): Promise<DocumentEditorRecord> {
  const file = tab.file || (tab.handle ? await tab.handle.getFile() : undefined);
  if (!tabs.includes(tab)) throw new DOMException('Office document was closed before it could be opened', 'AbortError');
  const container = document.createElement('div');
  container.className = 'document-editor-container';
  container.dataset.documentTabId = tab.id;
  container.dataset.active = 'false';
  container.setAttribute('aria-hidden', 'true');
  container.setAttribute('inert', '');
  elements.slot.append(container);

  let resolveReady!: (instance: OfficeEditorInstance) => void;
  let rejectReady!: (error: Error) => void;
  const ready = new Promise<OfficeEditorInstance>((resolve, reject) => {
    resolveReady = resolve;
    rejectReady = reject;
  });
  void ready.catch(() => undefined);

  let record!: DocumentEditorRecord;
  const options = {
    hostUrl: officeHostUrl,
    file,
    emptyType: tab.emptyType,
    fileName: tab.name,
    mode: 'edit' as const,
    saveBehavior: 'callback' as const,
    downloadedFonts: resourceManager?.getVerifiedFontPaths() || [],
    preferredHostSlot: tab.originSlot,
    hardResetOnLastDestroy,
    onReady: (instance: OfficeEditorInstance) => {
      if (tab.editor !== record || record.disposed) return;
      if (activeTab === tab) setDocumentStatus(instance.getState().fileType.toUpperCase());
      scheduleIdle(() => resourceManager?.prepareForDocumentType(documentTypeForName(tab.name)) || Promise.resolve());
    },
    onSave: (saved: File) => saveFile(tab, saved),
    onDirtyChange: (dirty: boolean, instance: OfficeEditorInstance) => {
      if (tab.editor !== record || record.disposed || record.instance?.id !== instance.id) return;
      tab.dirty = dirty;
      if (activeTab === tab) {
        setDocumentStatus(dirty ? '●' : instance.getState().fileType.toUpperCase());
      }
      renderTabs();
      if (!dirty) void tryActivateUpdate();
    },
    onError: (error: Error) => {
      if (tab.editor !== record || record.disposed || activeTab !== tab) return;
      setDocumentStatus(isEditorPoolExhausted(error) ? copy.openLimitReached : `${copy.error}: ${error.message}`);
    },
  };

  try {
    record = {
      container,
      mount: mountOfficeEditor(container, options),
      instance: null,
      origin: '',
      status: 'waiting',
      sequence: nextEditorSequence++,
      foreground: false,
      queued: false,
      disposed: false,
      settled: false,
      ready,
      resolveReady,
      rejectReady,
    };
  } catch (error) {
    container.remove();
    throw error;
  }
  record.origin = record.mount.getState().origin;
  tab.originSlot = originSlot(record.origin) || undefined;
  tab.file = undefined;
  tab.editor = record;
  await persistTab(tab);
  renderTabs();
  return record;
}

async function ensureEditor(tab: DocumentTab, foreground: boolean): Promise<OfficeEditorInstance> {
  let record = tab.editor;
  if (record?.status === 'error' || record?.status === 'disposed') {
    await disposeEditorRecord(tab);
    record = undefined;
  }
  if (!record) record = await createEditorRecord(tab);
  record.foreground ||= foreground;
  if (record.status === 'waiting') enqueueEditorActivation(record, foreground);
  if (record.instance) return record.instance;
  return record.ready;
}

async function activateTab(tab: DocumentTab): Promise<void> {
  if (!tabs.includes(tab)) return;
  const requestGeneration = ++tabSwitchGeneration;
  if (tab === activeTab && tab.editor?.instance) {
    setEditorVisibility(tab);
    setActiveDocumentHeader(tab);
    return;
  }
  if (!(await ensureResourcesReady(tab))) return;
  if (requestGeneration !== tabSwitchGeneration) return;
  pendingActivation = null;
  const previousTab = activeTab;
  if (previousTab?.dirty && previousTab !== tab) {
    const decision = await askDirtyDecision();
    if (decision === 'cancel') return;
    if (decision === 'save' && previousTab.editor?.instance) {
      await previousTab.editor.instance.save();
      if (previousTab.dirty) return;
    }
  }
  if (requestGeneration !== tabSwitchGeneration) return;
  activeTab = tab;
  setEditorVisibility(tab);
  renderTabs();
  setActiveDocumentHeader(tab);
  if (tab.inaccessible || (tab.handle && !(await hasReadPermission(tab.handle)))) {
    tab.inaccessible = true;
    setEditorVisibility(tab);
    setView('permission');
    return;
  }
  setView('editor');
  setDocumentStatus('…');
  try {
    const instance = await ensureEditor(tab, true);
    if (requestGeneration !== tabSwitchGeneration || activeTab !== tab) return;
    setEditorVisibility(tab);
    setActiveDocumentHeader(tab);
    setDocumentStatus(instance.getState().dirty ? '●' : instance.getState().fileType.toUpperCase());
    await persistTab(tab);
    warmBackgroundEditors();
  } catch (error) {
    if (requestGeneration !== tabSwitchGeneration || activeTab !== tab) return;
    const normalized = error instanceof Error ? error : new Error(String(error));
    setDocumentStatus(
      isEditorPoolExhausted(normalized) ? copy.openLimitReached : `${copy.error}: ${normalized.message}`,
    );
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
  if (tab.dirty) {
    const decision = await askDirtyDecision();
    if (decision === 'cancel') return;
    if (decision === 'save' && tab.editor?.instance) {
      await tab.editor.instance.save();
      if (tab.dirty) return;
    }
  }
  const index = tabs.indexOf(tab);
  if (index < 0) return;
  const wasActive = tab === activeTab;
  if (wasActive) {
    tabSwitchGeneration += 1;
    activeTab = null;
  }
  await disposeEditorRecord(tab);
  tabs.splice(index, 1);
  await tabStore.remove(tab.id).catch(() => undefined);
  renderTabs();
  if (wasActive) {
    const next = tabs[Math.min(index, tabs.length - 1)];
    if (next) await activateTab(next);
    else setView('empty');
  }
  void tryActivateUpdate();
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
  if (tabs.length >= MAX_OPEN_DOCUMENTS) {
    setDocumentStatus(copy.openLimitReached);
    return;
  }
  const tab: DocumentTab = { id: createDocumentTabId(), name: file.name, handle, file, dirty: false };
  tabs.push(tab);
  await persistTab(tab);
  renderTabs();
  await activateTab(tab);
}

async function openFiles(): Promise<void> {
  if (!(await ensureResourcesReady())) return;
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

async function createEmpty(emptyType: 'docx' | 'xlsx' | 'pptx' | 'csv'): Promise<void> {
  if (!(await ensureResourcesReady())) return;
  if (tabs.length >= MAX_OPEN_DOCUMENTS) {
    setDocumentStatus(copy.openLimitReached);
    return;
  }
  const tab: DocumentTab = {
    id: createDocumentTabId(),
    name: `${copy.newDocumentName}.${emptyType}`,
    emptyType,
    dirty: false,
  };
  tabs.push(tab);
  renderTabs();
  await activateTab(tab);
}

function renderResources(snapshot: OfficeRuntimeResourceSnapshot): void {
  latestResourceSnapshot = snapshot;
  elements.versionLabel.replaceChildren(
    Object.assign(document.createElement('span'), { textContent: copy.version }),
    ` ${snapshot.availablePackageVersion || snapshot.packageVersion}`,
  );
  const state =
    snapshot.readiness === 'ready'
      ? 'ready'
      : snapshot.readiness === 'updating' || snapshot.readiness === 'paused'
        ? 'updating'
        : snapshot.readiness === 'error' || snapshot.readiness === 'repair-needed'
          ? 'error'
          : 'needed';
  elements.resourceButton.dataset.state = state;
  elements.resourceButton.lastElementChild!.textContent =
    state === 'ready'
      ? copy.resourcesReady
      : state === 'updating'
        ? copy.resourcesUpdating
        : state === 'error'
          ? copy.resourcesError
          : copy.resourcesNeeded;
  if (!resourcePanelRoot) {
    resourcePanelRoot = createRoot(document.querySelector<HTMLElement>('#resource-react-root')!);
  }
  if (resourceManager) {
    resourcePanelRoot.render(createElement(OfficeResourcePanel, { manager: resourceManager, copy }));
  }
  if (snapshot.readiness === 'ready' && pendingActivation) {
    const tab = pendingActivation;
    pendingActivation = null;
    if (elements.resourceDialog.open) elements.resourceDialog.close();
    queueMicrotask(() => void activateTab(tab));
  }
}

async function initializeResources(): Promise<void> {
  try {
    const localMatrix =
      location.hostname === 'onlyoffice.localhost' && (location.protocol === 'http:' || location.protocol === 'https:');
    resourceManager = await createOfficeRuntimeResourceManager(
      localMatrix
        ? {
            canonicalOrigin: location.origin,
            allowLocalTestMode: true,
          }
        : undefined,
    );
    resourceManager.subscribe(renderResources);
    const snapshot = resourceManager.getSnapshot();
    renderResources(snapshot);
    startResourceMaintenance();
  } catch {
    const snapshot = resourceManager?.getSnapshot();
    if (snapshot) {
      renderResources(snapshot);
    } else {
      elements.resourceButton.lastElementChild!.textContent = copy.resourcesError;
      elements.resourceButton.dataset.state = 'error';
    }
  }
}

async function tryActivateUpdate(): Promise<void> {
  if (!waitingWorkbox || updateActivationPending || !(await allTabsSafeForUpdate())) return;
  updateActivationPending = true;
  elements.updateButton.disabled = true;
  try {
    await waitingWorkbox.messageSkipWaiting();
  } finally {
    updateActivationPending = false;
  }
}

function initializeServiceWorker(): void {
  if (!('serviceWorker' in navigator)) return;
  const workbox = new Workbox('/sw.js', { scope: '/' });
  let reloading = false;
  workbox.addEventListener('waiting', () => {
    waitingWorkbox = workbox;
    elements.updateBanner.hidden = false;
    void tryActivateUpdate();
  });
  workbox.addEventListener('controlling', () => {
    if (reloading || hasUnsafeWork()) return;
    const reloadKey = `onlyoffice-browser:pwa-reload:${ONLYOFFICE_BROWSER_VERSION}`;
    if (sessionStorage.getItem(reloadKey)) {
      elements.updateBanner.hidden = false;
      elements.updateButton.disabled = false;
      return;
    }
    sessionStorage.setItem(reloadKey, '1');
    reloading = true;
    location.reload();
  });
  void workbox.register().then(() => {
    const check = () => void workbox.update();
    check();
    const interval = window.setInterval(check, 10 * 60 * 1000);
    const visible = () => {
      if (document.visibilityState === 'visible') check();
    };
    addEventListener('online', check);
    document.addEventListener('visibilitychange', visible);
    addEventListener(
      'pagehide',
      () => {
        clearInterval(interval);
        removeEventListener('online', check);
        document.removeEventListener('visibilitychange', visible);
      },
      { once: true },
    );
  });
}

async function restoreTabs(): Promise<void> {
  if (!('indexedDB' in window)) return;
  try {
    const persisted = await tabStore.list();
    const restored = persisted.slice(-MAX_OPEN_DOCUMENTS);
    for (const record of restored) {
      tabs.push({
        id: record.id,
        name: record.name,
        handle: record.handle,
        originSlot: record.originSlot && isOfficeEditorOriginSlot(record.originSlot) ? record.originSlot : undefined,
        dirty: false,
        inaccessible: !(await hasReadPermission(record.handle)),
      });
    }
    renderTabs();
    if (tabs.length > 0) {
      const last = tabs[tabs.length - 1]!;
      await activateTab(last);
      warmBackgroundEditors();
    }
  } catch {
    // Private browsing and policy-restricted contexts may not offer IndexedDB.
  }
}

document.querySelector('#open-button')?.addEventListener('click', () => void openFiles());
elements.languageSelect.addEventListener('change', () => {
  applyLocale(elements.languageSelect.value as OfficeLocale);
});
document.querySelector('#empty-open-button')?.addEventListener('click', () => void openFiles());
document.querySelector('#save-button')?.addEventListener('click', () => void activeEditor()?.save());
document.querySelector('#close-button')?.addEventListener('click', () => activeTab && void closeTab(activeTab));
const newDocumentMenu = document.querySelector<HTMLDetailsElement>('.new-menu');
document.querySelectorAll<HTMLElement>('[data-new]').forEach((button) => {
  button.addEventListener('click', () => {
    if (newDocumentMenu) newDocumentMenu.open = false;
    void createEmpty(button.dataset.new as 'docx' | 'xlsx' | 'pptx' | 'csv');
  });
});
elements.fileInput.addEventListener('change', () => {
  for (const file of Array.from(elements.fileInput.files || [])) void addFile(undefined, file);
});
elements.resourceButton.addEventListener('click', () => elements.resourceDialog.showModal());
elements.updateButton.addEventListener('click', () => void tryActivateUpdate());
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
  get editor() {
    return activeEditor();
  },
  get editors() {
    return tabs
      .filter((tab) => tab.editor?.instance)
      .map((tab) => ({
        tabId: tab.id,
        fileName: tab.name,
        origin: tab.editor!.origin,
        instance: tab.editor!.instance,
      }));
  },
  get resourceManager() {
    return resourceManager;
  },
  get resourceSnapshot() {
    return resourceManager?.getSnapshot() ?? latestResourceSnapshot;
  },
  get resourceReady() {
    return latestResourceSnapshot?.readiness === 'ready';
  },
  openEmpty: createEmpty,
  closeAll: async () => {
    for (const tab of tabs.slice()) await closeTab(tab);
  },
};

applyLocale(locale);
initializeServiceWorker();
resourceInitialization = initializeResources();
void resourceInitialization;
void restoreTabs();
