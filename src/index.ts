import { createOfficeEditor, type OfficeEditorInstance, type OfficeEditorMode } from './lib/office-editor';
import './styles/base.css';

type DemoRecord = {
  id: number;
  instance: OfficeEditorInstance;
  panel: HTMLElement;
  status: HTMLElement;
  readonlyButton: HTMLButtonElement;
  saveButton: HTMLButtonElement;
};

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
      <label class="hard-reset-toggle">
        <input id="hard-reset-toggle" type="checkbox" />
        hard reset after last close
      </label>
    </div>
  </section>
  <input id="file-input" type="file" multiple accept=".docx,.xlsx,.pptx,.doc,.xls,.ppt,.csv" />
  <section id="editor-grid" class="editor-grid" aria-live="polite"></section>
`;

const grid = document.querySelector<HTMLElement>('#editor-grid')!;
const fileInput = document.querySelector<HTMLInputElement>('#file-input')!;
const hardResetToggle = document.querySelector<HTMLInputElement>('#hard-reset-toggle')!;

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
  record.saveButton.disabled = state.readonly;
}

function removeRecord(record: DemoRecord): void {
  const index = records.indexOf(record);
  if (index >= 0) {
    records.splice(index, 1);
  }
  record.instance.destroy();
  record.panel.remove();
}

async function openEditor(options: Parameters<typeof createOfficeEditor>[1]): Promise<DemoRecord> {
  const id = nextPanelId++;
  const title = options.fileName || (options.file instanceof File ? options.file.name : undefined) || options.emptyType || 'document';
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
        <button type="button" data-action="save">Save</button>
        <button type="button" data-action="close">Close</button>
      </div>
    </header>
    <div class="editor-slot"></div>
  `;

  const slot = panel.querySelector<HTMLElement>('.editor-slot')!;
  const status = panel.querySelector<HTMLElement>('[data-role="status"]')!;
  const readonlyButton = panel.querySelector<HTMLButtonElement>('[data-action="readonly"]')!;
  const saveButton = panel.querySelector<HTMLButtonElement>('[data-action="save"]')!;
  grid.appendChild(panel);

  const instance = await createOfficeEditor(slot, {
    ...options,
    hardResetOnLastDestroy: hardResetToggle.checked,
    onReady: (readyInstance) => {
      const state = readyInstance.getState();
      status.textContent = `${state.fileType.toUpperCase()} ${getModeLabel(state.mode)}`;
    },
    onSave: (file) => {
      status.textContent = `saved ${file.name} (${file.size} bytes)`;
    },
    onError: (error) => {
      status.textContent = `error: ${error.message}`;
    },
  });

  const record: DemoRecord = { id, instance, panel, status, readonlyButton, saveButton };
  records.push(record);
  const initialState = instance.getState();
  setStatus(record, `${initialState.fileType.toUpperCase()} ${getModeLabel(initialState.mode)}`);
  refreshPanelActions(record);

  saveButton.addEventListener('click', async () => {
    try {
      setStatus(record, 'saving');
      await instance.save();
    } catch (error) {
      setStatus(record, `save failed: ${error instanceof Error ? error.message : String(error)}`);
    }
  });

  readonlyButton.addEventListener('click', () => {
    const nextReadonly = !instance.getState().readonly;
    instance.setReadonly(nextReadonly);
    refreshPanelActions(record);
    const state = instance.getState();
    setStatus(record, `${state.fileType.toUpperCase()} ${getModeLabel(state.mode)}`);
  });

  panel.querySelector<HTMLButtonElement>('[data-action="close"]')?.addEventListener('click', () => {
    removeRecord(record);
  });

  return record;
}

function openEmpty(emptyType: 'docx' | 'xlsx' | 'pptx' | 'csv'): void {
  void openEditor({
    emptyType,
    fileName: `New_Document.${emptyType}`,
    mode: getSelectedMode(),
  });
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
    void openEditor({ file, fileName: file.name, mode: getSelectedMode() });
  }
  fileInput.value = '';
});

function closeAllEditors(): void {
  while (records.length > 0) {
    removeRecord(records[records.length - 1]);
  }
}

document.querySelector('#close-all-button')?.addEventListener('click', () => {
  closeAllEditors();
});

(window as typeof window & { __officeDemo?: unknown }).__officeDemo = {
  bootId,
  records,
  openEmpty,
  setMode: setSelectedMode,
  closeAll: closeAllEditors,
};
