import { createOfficeEditor, type OfficeEditorInstance, type OfficeEditorState } from './lib/office-editor';
import { convertBinToDocument, initX2T } from './lib/converter';
import { g_sEmpty_bin } from './lib/empty_bin';
import './styles/base.css';

type SaveScenario = 'local-file' | 'new-document';
type OfficeTestType = 'xlsx' | 'docx' | 'pptx';

type SaveE2EStatus = {
  scenario: SaveScenario;
  type: OfficeTestType;
  ready: boolean;
  dirty: boolean;
  writeCount: number;
  lastHash: string;
  lastSize: number;
  initialHash: string;
  initialSize: number;
  savedFileName: string;
  error: string;
  state: OfficeEditorState | null;
};

type SaveE2EController = {
  ready: Promise<void>;
  getStatus: () => SaveE2EStatus;
  getState: () => OfficeEditorState | null;
  destroy: () => Promise<void>;
};

declare global {
  interface Window {
    __ONLYOFFICE_SAVE_E2E__?: SaveE2EController;
  }
}

const supportedTypes: OfficeTestType[] = ['xlsx', 'docx', 'pptx'];
const rootElement = document.querySelector<HTMLElement>('#save-e2e-root');
const statusElement = document.querySelector<HTMLElement>('#save-e2e-status');
const editorElement = document.querySelector<HTMLElement>('#save-e2e-editor');

if (!rootElement || !statusElement || !editorElement) {
  throw new Error('Missing save E2E page root');
}

const root = rootElement;
const statusEl = statusElement;
const editorEl = editorElement;

root.style.height = '100vh';
root.style.height = '100dvh';
root.style.display = 'flex';
root.style.flexDirection = 'column';
root.style.background = '#f7f8fa';
statusEl.style.minHeight = '34px';
statusEl.style.padding = '8px 12px';
statusEl.style.borderBottom = '1px solid #dfe4ec';
statusEl.style.background = '#fff';
statusEl.style.color = '#526073';
statusEl.style.fontSize = '13px';
editorEl.style.flex = '1';
editorEl.style.minHeight = '0';
editorEl.style.background = '#fff';

const params = new URLSearchParams(window.location.search);
const scenario = (params.get('scenario') === 'new-document' ? 'new-document' : 'local-file') satisfies SaveScenario;
const requestedType = params.get('type') || 'xlsx';
const type: OfficeTestType = supportedTypes.includes(requestedType as OfficeTestType)
  ? (requestedType as OfficeTestType)
  : 'xlsx';

let editor: OfficeEditorInstance | null = null;
let readyResolve!: () => void;
let readyReject!: (error: Error) => void;
let state: OfficeEditorState | null = null;
let dirty = false;
let writeCount = 0;
let lastHash = '';
let lastSize = 0;
let initialHash = '';
let initialSize = 0;
let savedFileName = '';
let errorMessage = '';

const ready = new Promise<void>((resolve, reject) => {
  readyResolve = resolve;
  readyReject = reject;
});

function getDefaultOfficeHostUrl(): string {
  const configured = new URLSearchParams(window.location.search).get('hostUrl');
  if (configured) return new URL(configured, window.location.href).href;

  const hostUrl = new URL('/office-host.html', window.location.href);
  if (hostUrl.hostname === 'localhost' || (hostUrl.hostname.endsWith('.localhost') && hostUrl.hostname !== 'host.localhost')) {
    hostUrl.hostname = 'host.localhost';
  } else if (hostUrl.hostname === 'host.localhost') {
    hostUrl.hostname = 'app.localhost';
  } else if (hostUrl.hostname === '127.0.0.1') {
    hostUrl.hostname = 'localhost';
  } else if (!hostUrl.hostname.endsWith('.localhost')) {
    hostUrl.hostname = `host.${hostUrl.hostname}`;
  }
  return hostUrl.href;
}

function mimeTypeForType(fileType: OfficeTestType): string {
  if (fileType === 'docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (fileType === 'pptx') return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
}

async function hashFile(file: File): Promise<string> {
  const buffer = await file.arrayBuffer();
  const digest = await crypto.subtle.digest('SHA-256', buffer);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function setStatus(message: string): void {
  statusEl.textContent = message;
}

function syncState(instance?: OfficeEditorInstance): void {
  state = instance?.getState() || editor?.getState() || state;
  dirty = Boolean(state?.dirty);
  const bits = [
    scenario,
    type,
    state?.status || 'opening',
    dirty ? 'dirty' : 'clean',
    `writes=${writeCount}`,
    lastSize ? `size=${lastSize}` : '',
  ].filter(Boolean);
  setStatus(bits.join(' | '));
}

class MemoryFileHandle {
  private file: File;

  constructor(file: File) {
    this.file = file;
  }

  async getFile(): Promise<File> {
    const copy = await this.file.arrayBuffer();
    return new File([copy], this.file.name, { type: this.file.type });
  }

  async write(file: File): Promise<void> {
    const copy = await file.arrayBuffer();
    this.file = new File([copy], this.file.name, { type: file.type || this.file.type });
    writeCount += 1;
    savedFileName = file.name;
    lastSize = file.size;
    lastHash = await hashFile(file);
    syncState();
  }

  resetWriteCount(): void {
    writeCount = 0;
  }
}

async function seedExistingOfficeFile(fileType: OfficeTestType): Promise<MemoryFileHandle> {
  await initX2T();
  const template = g_sEmpty_bin[`.${fileType}`];
  if (!template) {
    throw new Error(`Missing empty document template: ${fileType}`);
  }
  const bin = new Uint8Array(await new Blob([template], { type: 'application/octet-stream' }).arrayBuffer());
  const converted = await convertBinToDocument(bin, `local-save.${fileType}`, fileType.toUpperCase());
  const seededFile = new File([converted.data], converted.fileName || `local-save.${fileType}`, {
    type: mimeTypeForType(fileType),
  });

  const handle = new MemoryFileHandle(seededFile);
  initialSize = seededFile.size;
  initialHash = await hashFile(seededFile);
  lastSize = initialSize;
  lastHash = initialHash;
  savedFileName = seededFile.name;
  handle.resetWriteCount();
  return handle;
}

async function openLocalFileScenario(): Promise<void> {
  setStatus('seeding local file');
  const handle = await seedExistingOfficeFile(type);
  const file = await handle.getFile();

  setStatus('opening local file');
  editor = await createOfficeEditor(editorEl, {
    hostUrl: getDefaultOfficeHostUrl(),
    file,
    fileName: file.name,
    mode: 'edit',
    saveBehavior: 'callback',
    onReady: (instance) => {
      state = instance.getState();
      syncState(instance);
    },
    onDirtyChange: (_dirty, instance) => {
      syncState(instance);
    },
    onSave: async (fileToSave) => {
      await handle.write(fileToSave);
      return true;
    },
    onError: (error) => {
      errorMessage = error.message;
      setStatus(`error: ${error.message}`);
    },
  });
  syncState(editor);
}

async function openNewDocumentScenario(): Promise<void> {
  setStatus('opening new document');
  editor = await createOfficeEditor(editorEl, {
    hostUrl: getDefaultOfficeHostUrl(),
    emptyType: type,
    fileName: `New_Document.${type}`,
    mode: 'edit',
    saveBehavior: 'download',
    onReady: (instance) => {
      state = instance.getState();
      syncState(instance);
    },
    onDirtyChange: (_dirty, instance) => {
      syncState(instance);
    },
    onError: (error) => {
      errorMessage = error.message;
      setStatus(`error: ${error.message}`);
    },
  });
  syncState(editor);
}

window.__ONLYOFFICE_SAVE_E2E__ = {
  ready,
  getStatus: () => ({
    scenario,
    type,
    ready: Boolean(editor && state?.status === 'ready'),
    dirty,
    writeCount,
    lastHash,
    lastSize,
    initialHash,
    initialSize,
    savedFileName,
    error: errorMessage,
    state,
  }),
  getState: () => state,
  destroy: async () => {
    await editor?.destroy();
    editor = null;
    syncState();
  },
};

void (async () => {
  try {
    if (scenario === 'new-document') {
      await openNewDocumentScenario();
    } else {
      await openLocalFileScenario();
    }
    readyResolve();
  } catch (error) {
    const normalized = error instanceof Error ? error : new Error(String(error));
    errorMessage = normalized.message;
    setStatus(`error: ${normalized.message}`);
    readyReject(normalized);
  }
})();
