import { g_sEmpty_bin } from './empty_bin';
import { c_oAscFileType2 } from './file-types';
import { convertBinToDocument, convertDocument, initX2T } from './converter';
import { getDocumentType, getMimeTypeFromExtension, BASE_PATH } from './document-utils';
import { getOnlyOfficeLang, t } from './i18n';
import { createOnlyOfficeMockServer, LOCAL_ONLYOFFICE_USER_ID, LOCAL_ONLYOFFICE_USER_NAME } from './onlyoffice-mock-server';
import type { BinConversionResult, SaveEvent } from './document-types';
import { assertGeneratedFontAssetsAvailable } from './font-assets';

const ONLYOFFICE_BROWSER_BUILD_VERSION = '9.3.0';
const ONLYOFFICE_BROWSER_BUILD_NUMBER = 140;
const ONLYOFFICE_ZOOM_FIT_TO_WIDTH = -2;
const SAVE_TIMEOUT_MS = 60_000;
const BLANK_NAVIGATION_TIMEOUT_MS = 250;
const SUPPORTED_EMPTY_TYPES = ['docx', 'xlsx', 'pptx', 'csv'] as const;

type OfficeEmptyType = (typeof SUPPORTED_EMPTY_TYPES)[number];
type OfficeEditorStatus = 'opening' | 'ready' | 'destroyed' | 'error';
export type OfficeEditorMode = 'edit' | 'readonly' | 'preview';

export type OfficeEditorInput = Blob | ArrayBuffer | Uint8Array;

export interface CreateOfficeEditorOptions {
  file?: File | Blob;
  buffer?: OfficeEditorInput;
  url?: string;
  emptyType?: OfficeEmptyType;
  fileName?: string;
  mode?: OfficeEditorMode;
  readonly?: boolean;
  spellcheck?: boolean;
  lang?: string;
  fetchOptions?: RequestInit;
  hardResetOnLastDestroy?: boolean;
  onReady?: (instance: OfficeEditorInstance) => void;
  onSave?: (file: File, instance: OfficeEditorInstance) => void | Promise<void>;
  onError?: (error: Error, instance?: OfficeEditorInstance) => void;
}

export interface OfficeEditorState {
  id: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
  readonly: boolean;
  status: OfficeEditorStatus;
  destroyed: boolean;
}

export interface OfficeEditorInstance {
  readonly id: string;
  save(targetExt?: string): Promise<File>;
  setReadonly(readonly: boolean): void;
  destroy(): Promise<void>;
  getState(): OfficeEditorState;
}

type PreparedDocument = {
  fileName: string;
  fileType: string;
  binData: BlobPart;
  media?: Record<string, string>;
  sourceObjectUrl?: string;
  originalFile?: File;
};

type SaveRequest = {
  targetExt: string;
  resolve: (file: File) => void;
  reject: (error: Error) => void;
  timeoutId: number;
  settled: boolean;
};

type RuntimeWindow = typeof window & {
  DocsAPI?: DocsAPI;
  __fonts_visible_names?: unknown;
  APP?: {
    getImageURL?: (name: string, callback: (url: string) => void) => void;
    [key: string]: unknown;
  };
};

type OnlyOfficeFrameWindow = Window & {
  AscCommon?: {
    baseEditorsApi?: {
      prototype?: {
        sync_InitEditorFonts?: (guiFonts: unknown[]) => unknown;
        __onlyOfficeBrowserFontPickerFilter?: boolean;
      };
    };
  };
  __fonts_visible_names?: unknown;
};

let editorApiPromise: Promise<void> | null = null;
let nextEditorId = 1;
const activeInstances = new Map<string, BrowserOfficeEditor>();
const mediaRegistry = new Map<string, Record<string, string>>();

function createObjectUrl(blob: Blob): string {
  return URL.createObjectURL(blob);
}

function toError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}

function resolveInitialMode(options: CreateOfficeEditorOptions): OfficeEditorMode {
  if (options.readonly && options.mode !== 'preview') return 'readonly';
  return options.mode || (options.readonly ? 'readonly' : 'edit');
}

function shouldFitEditorModePreviewToWidth(fileType: string, previewMode: boolean): boolean {
  if (previewMode) return false;
  const documentType = getDocumentType(fileType);
  return documentType === 'word' || documentType === 'slide';
}

function getDefaultEditorModePreviewZoom(fileType: string, previewMode: boolean): number | undefined {
  return shouldFitEditorModePreviewToWidth(fileType, previewMode) ? ONLYOFFICE_ZOOM_FIT_TO_WIDTH : undefined;
}

function persistDefaultEditorModePreviewZoom(fileType: string, previewMode: boolean): void {
  if (!shouldFitEditorModePreviewToWidth(fileType, previewMode)) return;

  const settingsKey = getDocumentType(fileType) === 'slide' ? 'pe-settings-zoom' : 'de-settings-zoom';
  try {
    window.localStorage.setItem(settingsKey, String(ONLYOFFICE_ZOOM_FIT_TO_WIDTH));
  } catch {
    // Storage can be disabled in private contexts; customization.zoom remains the primary signal.
  }
}

function normalizeExtension(value: string | undefined, fallback: string): string {
  return (value || fallback).replace(/^\./, '').toLowerCase();
}

function getFileExtension(fileName: string): string {
  return fileName.split('.').pop()?.toLowerCase() || '';
}

function getDefaultFileName(emptyType: OfficeEmptyType): string {
  return `New_Document.${emptyType}`;
}

function getSavedFileMimeType(fileName: string): string {
  const extension = getFileExtension(fileName);
  const mimeMap: Record<string, string> = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    csv: 'text/csv',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt: 'application/vnd.ms-powerpoint',
    pdf: 'application/pdf',
  };
  return mimeMap[extension] || 'application/octet-stream';
}

function toUint8Array(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) {
    return data;
  }
  if (data instanceof ArrayBuffer) {
    return new Uint8Array(data);
  }
  if (ArrayBuffer.isView(data)) {
    const arrayBuffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer;
    return new Uint8Array(arrayBuffer);
  }
  throw new Error('Unsupported binary data type');
}

function makeFileFromInput(input: OfficeEditorInput | Blob, fileName: string): File {
  if (input instanceof File) {
    return input;
  }
  if (input instanceof Blob) {
    return new File([input], fileName, { type: input.type || getSavedFileMimeType(fileName) });
  }
  if (input instanceof Uint8Array) {
    const arrayBuffer = input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength) as ArrayBuffer;
    return new File([arrayBuffer], fileName, { type: getSavedFileMimeType(fileName) });
  }
  return new File([input], fileName, { type: getSavedFileMimeType(fileName) });
}

function readFileNameFromResponse(url: string, response: Response, fallback = 'document.docx'): string {
  const contentDisposition = response.headers.get('Content-Disposition');
  const match = contentDisposition?.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
  if (match?.[1]) {
    return match[1].replace(/['"]/g, '');
  }

  try {
    return new URL(url, window.location.href).pathname.split('/').pop() || fallback;
  } catch {
    return fallback;
  }
}

function resolveMediaFromRegistry(name: string): string {
  const candidates = (() => {
    const trimmed = name.replace(/^(\.\/|\/)+/, '');
    const decoded = (() => {
      try {
        return decodeURIComponent(trimmed);
      } catch {
        return trimmed;
      }
    })();
    const values = [name, trimmed, decoded];
    for (const value of [trimmed, decoded]) {
      if (value && !value.startsWith('media/')) {
        values.push(`media/${value}`);
      }
    }
    return [...new Set(values.filter(Boolean))];
  })();

  for (const media of mediaRegistry.values()) {
    for (const candidate of candidates) {
      const url = media[candidate];
      if (url) {
        return url;
      }
    }
  }
  return '';
}

function installOnlyOfficeParentApp(): void {
  const runtimeWindow = window as RuntimeWindow;
  runtimeWindow.APP = runtimeWindow.APP || {};
  runtimeWindow.APP.getImageURL = (name, callback) => {
    callback(resolveMediaFromRegistry(name));
  };
}

function registerMedia(id: string, media: Record<string, string>): void {
  mediaRegistry.set(id, media);
  installOnlyOfficeParentApp();
}

function unregisterMedia(id: string): void {
  mediaRegistry.delete(id);
  installOnlyOfficeParentApp();
}

function maybeHardResetPage(): void {
  const url = `${window.location.pathname}${window.location.search}`;
  window.setTimeout(() => {
    window.location.replace(`/reset.html?to=${encodeURIComponent(url)}`);
  }, 0);
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function getFontNameForFilter(font: unknown): string {
  if (!font || typeof font !== 'object') return '';
  const value = font as { name?: unknown; Name?: unknown; asc_getFontName?: unknown };
  if (typeof value.name === 'string') return value.name;
  if (typeof value.Name === 'string') return value.Name;
  if (typeof value.asc_getFontName === 'function') {
    try {
      const name = value.asc_getFontName();
      return typeof name === 'string' ? name : '';
    } catch {
      return '';
    }
  }
  return '';
}

function installFontPickerFilter(frameWindow: OnlyOfficeFrameWindow): boolean {
  const visibleNames =
    Array.isArray(frameWindow.__fonts_visible_names) &&
    frameWindow.__fonts_visible_names.every((name) => typeof name === 'string')
      ? new Set(frameWindow.__fonts_visible_names)
      : null;
  if (!visibleNames || visibleNames.size === 0) return true;

  const prototype = frameWindow.AscCommon?.baseEditorsApi?.prototype;
  const original = prototype?.sync_InitEditorFonts;
  if (!prototype || typeof original !== 'function') return false;
  if (prototype.__onlyOfficeBrowserFontPickerFilter) return true;

  prototype.sync_InitEditorFonts = function syncInitEditorFontsWithFilter(this: unknown, guiFonts: unknown[]) {
    const filteredFonts = Array.isArray(guiFonts)
      ? guiFonts.filter((font) => {
          const name = getFontNameForFilter(font);
          return name ? visibleNames.has(name) : false;
        })
      : guiFonts;
    return original.call(this, filteredFonts);
  };
  prototype.__onlyOfficeBrowserFontPickerFilter = true;
  return true;
}

function installNestedFontPickerFilter(): void {
  const startedAt = Date.now();
  const timeoutMs = 10_000;
  const intervalMs = 50;

  const tryInstall = () => {
    const frame = document.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameWindow = frame?.contentWindow as OnlyOfficeFrameWindow | null | undefined;
    if (frameWindow) {
      try {
        if (installFontPickerFilter(frameWindow)) return;
      } catch {
        return;
      }
    }

    if (Date.now() - startedAt < timeoutMs) {
      window.setTimeout(tryInstall, intervalMs);
    }
  };

  tryInstall();
}

export function loadOfficeEditorApi(): Promise<void> {
  const runtimeWindow = window as RuntimeWindow;
  if (runtimeWindow.DocsAPI?.DocEditor) {
    return Promise.resolve();
  }
  if (editorApiPromise) {
    return editorApiPromise;
  }

  editorApiPromise = new Promise((resolve, reject) => {
    const apiUrl = new URL(`${BASE_PATH}web-apps/apps/api/documents/api.js`, window.location.href).href;
    const existing = Array.from(document.scripts).find((script) => script.src === apiUrl);
    if (existing) {
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new Error(t('failedToLoadEditor'))), { once: true });
      return;
    }

    const script = document.createElement('script');
    script.src = apiUrl;
    script.async = true;
    script.dataset.officeEditorApi = 'true';
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(t('failedToLoadEditor')));
    document.head.appendChild(script);
  });

  return editorApiPromise;
}

async function prepareDocument(options: CreateOfficeEditorOptions): Promise<PreparedDocument> {
  if (options.emptyType) {
    const emptyType = normalizeExtension(options.emptyType, 'docx') as OfficeEmptyType;
    if (!SUPPORTED_EMPTY_TYPES.includes(emptyType)) {
      throw new Error(`Unsupported empty document type: ${options.emptyType}`);
    }

    const fileName = options.fileName || getDefaultFileName(emptyType);
    if (emptyType === 'csv') {
      const file = new File(['\ufeff'], fileName, { type: 'text/csv' });
      const converted = await convertDocument(file);
      return {
        fileName,
        fileType: 'csv',
        binData: converted.bin,
        media: converted.media,
        originalFile: file,
      };
    }

    const template = g_sEmpty_bin[`.${emptyType}`];
    if (!template) {
      throw new Error(`Missing empty document template: ${emptyType}`);
    }

    return {
      fileName,
      fileType: emptyType,
      binData: template,
      media: {},
    };
  }

  if (options.url) {
    const response = await fetch(options.url, options.fetchOptions);
    if (!response.ok) {
      throw new Error(`Failed to fetch document: ${response.status} ${response.statusText}`);
    }
    const fileName = options.fileName || readFileNameFromResponse(options.url, response);
    const blob = await response.blob();
    const file = new File([blob], fileName, { type: blob.type || getSavedFileMimeType(fileName) });
    const sourceObjectUrl = createObjectUrl(file);
    const converted = await convertDocument(file);
    return {
      fileName: converted.fileName || fileName,
      fileType: getFileExtension(fileName),
      binData: converted.bin,
      media: converted.media,
      sourceObjectUrl,
      originalFile: file,
    };
  }

  const input = options.file || options.buffer;
  if (!input) {
    throw new Error('createOfficeEditor requires file, buffer, url, or emptyType');
  }

  const fileName = options.fileName || (input instanceof File ? input.name : 'document.docx');
  const file = makeFileFromInput(input, fileName);
  const sourceObjectUrl = createObjectUrl(file);
  const converted = await convertDocument(file);
  return {
    fileName: converted.fileName || fileName,
    fileType: getFileExtension(fileName),
    binData: converted.bin,
    media: converted.media,
    sourceObjectUrl,
    originalFile: file,
  };
}

class BrowserOfficeEditor implements OfficeEditorInstance {
  readonly id: string;
  private readonly container: HTMLElement;
  private readonly options: CreateOfficeEditorOptions;
  private readonly placeholder: HTMLDivElement;
  private editor: DocEditor | null = null;
  private editorBinUrl: string | null = null;
  private media: Record<string, string> = {};
  private saveRequest: SaveRequest | null = null;
  private sourceObjectUrl: string | null = null;
  private status: OfficeEditorStatus = 'opening';
  private destroyed = false;
  private readonly fileName: string;
  private readonly fileType: string;
  private readonly originalFile?: File;
  private readonly editorLang: string;
  private readonly readonlyMode: { value: boolean };
  private readonly previewMode: boolean;
  private destroyPromise: Promise<void> | null = null;

  private constructor(
    container: HTMLElement,
    options: CreateOfficeEditorOptions,
    prepared: PreparedDocument,
    placeholder: HTMLDivElement,
  ) {
    this.id = `office-editor-${nextEditorId++}`;
    this.container = container;
    this.options = options;
    this.placeholder = placeholder;
    this.fileName = prepared.fileName;
    this.fileType = prepared.fileType;
    this.originalFile = prepared.originalFile;
    this.sourceObjectUrl = prepared.sourceObjectUrl || null;
    this.media = prepared.media || {};
    this.editorLang = options.lang || getOnlyOfficeLang();
    const initialMode = resolveInitialMode(options);
    this.previewMode = initialMode === 'preview';
    this.readonlyMode = { value: initialMode !== 'edit' };
  }

  static async create(container: HTMLElement, options: CreateOfficeEditorOptions): Promise<BrowserOfficeEditor> {
    await assertGeneratedFontAssetsAvailable();
    await loadOfficeEditorApi();
    await initX2T();
    const prepared = await prepareDocument(options);
    const placeholder = document.createElement('div');
    const instance = new BrowserOfficeEditor(container, options, prepared, placeholder);
    await instance.mount(prepared);
    return instance;
  }

  private async mount(prepared: PreparedDocument): Promise<void> {
    const runtimeWindow = window as RuntimeWindow;
    if (!runtimeWindow.DocsAPI?.DocEditor) {
      throw new Error('OnlyOffice 9.3.0 browser wrapper is not available');
    }

    this.container.replaceChildren();
    this.container.classList.add('office-editor-host');
    this.placeholder.id = `${this.id}-frame`;
    this.placeholder.className = 'office-editor-frame';
    this.container.appendChild(this.placeholder);

    this.editorBinUrl = createObjectUrl(new Blob([prepared.binData], { type: 'application/octet-stream' }));
    registerMedia(this.id, this.media);
    activeInstances.set(this.id, this);

    try {
      installNestedFontPickerFilter();
      const defaultZoom = getDefaultEditorModePreviewZoom(this.fileType, this.previewMode);
      persistDefaultEditorModePreviewZoom(this.fileType, this.previewMode);
      const canEditInitially = !this.previewMode && !this.readonlyMode.value;
      const editor = new runtimeWindow.DocsAPI.DocEditor(this.placeholder.id, {
        type: this.previewMode ? 'embedded' : 'desktop',
        width: '100%',
        height: '100%',
        documentType: getDocumentType(this.fileType) || undefined,
        document: {
          title: this.fileName,
          url: this.editorBinUrl,
          fileType: this.fileType,
          key: `local-${this.id}-${Date.now()}`,
          permissions: {
            edit: canEditInitially,
            download: !this.previewMode,
            chat: false,
            protect: false,
          },
        },
        editorConfig: {
          lang: this.editorLang,
          mode: this.previewMode ? 'view' : 'edit',
          user: {
            id: LOCAL_ONLYOFFICE_USER_ID,
            name: LOCAL_ONLYOFFICE_USER_NAME,
          },
          embedded: this.previewMode
            ? {
                autostart: 'document',
                toolbarDocked: 'top',
              }
            : undefined,
          customization: {
            help: false,
            about: false,
            hideRightMenu: true,
            compactToolbar: true,
            zoom: defaultZoom,
            spellcheck: this.options.spellcheck ?? false,
            plugins: false,
            features: {
              featuresTips: false,
              spellcheck: {
                change: false,
              },
            },
            anonymous: {
              request: false,
              label: LOCAL_ONLYOFFICE_USER_NAME,
            },
          },
        },
        events: {
          onAppReady: () => {
            installNestedFontPickerFilter();
          },
          onDocumentReady: () => {
            if (this.destroyed) return;
            installNestedFontPickerFilter();
            this.status = 'ready';
            this.applyDefaultEditorModePreviewZoom();
            this.options.onReady?.(this);
          },
          onSave: (event: SaveEvent) => {
            void this.handleSaveDocument(event);
          },
          onDownloadAs: (event: { data?: { url?: string; fileType?: string | number } }) => {
            void this.handleDownloadAs(event);
          },
          writeFile: (event: any) => {
            void this.handleWriteFile(event);
          },
        },
      });

      this.editor = editor;
      this.attachDestroyCleanup(editor);

      if (typeof editor.connectMockServer !== 'function') {
        throw new Error('OnlyOffice 9.3.0 browser wrapper is not available');
      }

      editor.connectMockServer(
        createOnlyOfficeMockServer({
          media: this.media,
          buildVersion: ONLYOFFICE_BROWSER_BUILD_VERSION,
          buildNumber: ONLYOFFICE_BROWSER_BUILD_NUMBER,
          onMessage: (msg) => {
            if (msg.type === 'auth' && msg.openCmd?.url && msg.openCmd.url !== this.editorBinUrl) {
              console.warn('OnlyOffice auth requested an unexpected document URL');
            }
          },
          onCorruptionWarning: (duplicateId) => {
            console.warn('OnlyOffice corruption warning:', duplicateId);
          },
        }),
      );
      installOnlyOfficeParentApp();
    } catch (error) {
      this.status = 'error';
      this.options.onError?.(toError(error), this);
      await this.destroy();
      throw error;
    }
  }

  private attachDestroyCleanup(editor: DocEditor): void {
    const destroyEditor = editor.destroyEditor.bind(editor);
    let wrapperDestroyed = false;
    editor.destroyEditor = () => {
      if (wrapperDestroyed) return;
      wrapperDestroyed = true;
      destroyEditor();
    };
  }

  private sendOnlyOfficeCommand(command: string, data: Record<string, unknown>): void {
    if (!this.editor) return;
    if (typeof this.editor.serviceCommand === 'function') {
      this.editor.serviceCommand(command, data);
      return;
    }
    this.editor.sendCommand?.({ command, data });
  }

  private applyReadonlyState(): void {
    if (!this.editor || this.previewMode) return;
    const canEdit = !this.readonlyMode.value;
    const message = canEdit ? '' : 'Readonly mode';
    if (typeof this.editor.processRightsChange === 'function') {
      this.editor.processRightsChange(canEdit, message);
      return;
    }
    this.sendOnlyOfficeCommand('processRightsChange', {
      enabled: canEdit,
      message,
    });
  }

  private applyDefaultEditorModePreviewZoom(): void {
    if (!this.editor || !shouldFitEditorModePreviewToWidth(this.fileType, this.previewMode)) return;
    this.editor.zoomFitToWidth?.();
  }

  private cleanupSaveRequest(error?: Error): void {
    const request = this.saveRequest;
    if (!request) return;
    window.clearTimeout(request.timeoutId);
    this.saveRequest = null;
    if (error && !request.settled) {
      request.settled = true;
      request.reject(error);
    }
  }

  private resolveSaveRequest(file: File): void {
    const request = this.saveRequest;
    if (!request || request.settled) return;
    request.settled = true;
    window.clearTimeout(request.timeoutId);
    this.saveRequest = null;
    request.resolve(file);
  }

  private rejectSaveRequest(error: Error): void {
    const request = this.saveRequest;
    if (!request || request.settled) return;
    request.settled = true;
    window.clearTimeout(request.timeoutId);
    this.saveRequest = null;
    request.reject(error);
  }

  private async convertSavedBin(data: Uint8Array | ArrayBuffer, targetFormat: string): Promise<File> {
    const result: BinConversionResult = await convertBinToDocument(toUint8Array(data), this.fileName, targetFormat);
    const bytes = toUint8Array(result.data);
    return new File([bytes as BlobPart], result.fileName, { type: getSavedFileMimeType(result.fileName) });
  }

  private async emitSavedFile(file: File): Promise<void> {
    this.resolveSaveRequest(file);
    await this.options.onSave?.(file, this);
  }

  private async handleSaveDocument(event: SaveEvent): Promise<void> {
    try {
      const payload = event.data?.data?.data;
      if (!payload) {
        throw new Error('Save event did not include document data');
      }

      const optionFormat =
        c_oAscFileType2[event.data.option.outputformat] || getFileExtension(this.fileName).toUpperCase();
      const targetFormat = this.fileName.toLowerCase().endsWith('.csv')
        ? 'CSV'
        : this.saveRequest?.targetExt || optionFormat;
      const file = await this.convertSavedBin(payload, targetFormat);
      await this.emitSavedFile(file);
      this.sendOnlyOfficeCommand('asc_onSaveCallback', { err_code: 0 });
    } catch (error) {
      const normalized = toError(error);
      this.rejectSaveRequest(normalized);
      this.options.onError?.(normalized, this);
      this.sendOnlyOfficeCommand('asc_onSaveCallback', { err_code: 1 });
    }
  }

  private async handleDownloadAs(event: { data?: { url?: string; fileType?: string | number } }): Promise<void> {
    try {
      const url = event.data?.url;
      if (!url) {
        throw new Error('Download URL is empty');
      }
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`Failed to fetch exported file: ${response.status} ${response.statusText}`);
      }

      const blob = await response.blob();
      const baseName = this.fileName.replace(/\.[^/.]+$/, '');
      const ext = String(
        this.saveRequest?.targetExt || event.data?.fileType || getFileExtension(this.fileName),
      ).toLowerCase();
      const fileName = `${baseName}.${ext}`;
      await this.emitSavedFile(new File([blob], fileName, { type: blob.type || getSavedFileMimeType(fileName) }));
    } catch (error) {
      const normalized = toError(error);
      this.rejectSaveRequest(normalized);
      this.options.onError?.(normalized, this);
    }
  }

  private async handleWriteFile(event: any): Promise<void> {
    try {
      const eventData = event.data;
      const imageData = eventData?.data;
      const fileName = eventData?.file;
      if (!(imageData instanceof Uint8Array) || typeof fileName !== 'string') {
        throw new Error('Invalid writeFile event data');
      }

      const extension = fileName.split('.').pop()?.toLowerCase() || 'png';
      const objectUrl = createObjectUrl(
        new Blob([imageData as BlobPart], { type: getMimeTypeFromExtension(extension) }),
      );
      this.media[`media/${fileName}`] = objectUrl;
      registerMedia(this.id, this.media);

      this.sendOnlyOfficeCommand('asc_setImageUrls', { urls: this.media });
      this.sendOnlyOfficeCommand('asc_writeFileCallback', {
        path: objectUrl,
        imgName: fileName,
      });
    } catch (error) {
      const normalized = toError(error);
      this.sendOnlyOfficeCommand('asc_writeFileCallback', {
        success: false,
        error: normalized.message,
      });
      event.callback?.({ success: false, error: normalized.message });
      this.options.onError?.(normalized, this);
    }
  }

  save(targetExt?: string): Promise<File> {
    if (this.destroyed || !this.editor) {
      return Promise.reject(new Error('Editor is not open'));
    }
    if (this.readonlyMode.value) {
      return Promise.reject(new Error('Current document is readonly'));
    }
    if (this.saveRequest) {
      return Promise.reject(new Error('A save request is already in progress for this editor'));
    }
    if (typeof this.editor.downloadAs !== 'function') {
      return Promise.reject(new Error('The current editor does not support downloadAs export'));
    }

    const normalizedTargetExt = (targetExt || getFileExtension(this.fileName) || 'docx').toUpperCase();
    return new Promise<File>((resolve, reject) => {
      const timeoutId = window.setTimeout(() => {
        this.rejectSaveRequest(new Error('Save request timed out before receiving edited file data'));
      }, SAVE_TIMEOUT_MS);

      this.saveRequest = {
        targetExt: normalizedTargetExt,
        resolve,
        reject,
        timeoutId,
        settled: false,
      };

      this.editor?.downloadAs?.(normalizedTargetExt);
    });
  }

  setReadonly(readonly: boolean): void {
    if (this.previewMode && !readonly) {
      throw new Error('Preview mode cannot be switched to editing; recreate the editor with mode: "edit".');
    }
    this.readonlyMode.value = readonly;
    this.applyReadonlyState();
  }

  getState(): OfficeEditorState {
    const mode: OfficeEditorMode = this.previewMode ? 'preview' : this.readonlyMode.value ? 'readonly' : 'edit';
    return {
      id: this.id,
      fileName: this.fileName,
      fileType: this.fileType,
      mode,
      readonly: this.readonlyMode.value,
      status: this.status,
      destroyed: this.destroyed,
    };
  }

  destroy(): Promise<void> {
    if (this.destroyPromise) return this.destroyPromise;
    this.destroyed = true;
    this.status = 'destroyed';
    this.cleanupSaveRequest(new Error('Editor was destroyed before save completed'));

    this.destroyPromise = this.destroyInternal();
    return this.destroyPromise;
  }

  private async destroyInternal(): Promise<void> {
    try {
      this.editor?.destroyEditor?.();
    } catch (error) {
      console.warn('Failed to destroy editor:', error);
    }

    this.editor = null;
    await this.blankAndRemoveIframes();
    this.placeholder.remove();
    this.container.replaceChildren();

    if (this.editorBinUrl) {
      URL.revokeObjectURL(this.editorBinUrl);
      this.editorBinUrl = null;
    }
    if (this.sourceObjectUrl) {
      URL.revokeObjectURL(this.sourceObjectUrl);
      this.sourceObjectUrl = null;
    }

    for (const url of Object.values(this.media)) {
      if (url.startsWith('blob:')) {
        URL.revokeObjectURL(url);
      }
    }
    this.media = {};
    unregisterMedia(this.id);
    activeInstances.delete(this.id);

    if (this.options.hardResetOnLastDestroy && activeInstances.size === 0) {
      maybeHardResetPage();
    }
  }

  private async blankAndRemoveIframes(): Promise<void> {
    const frames = new Set<HTMLIFrameElement>([
      ...Array.from(this.placeholder.querySelectorAll<HTMLIFrameElement>('iframe')),
      ...Array.from(this.container.querySelectorAll<HTMLIFrameElement>('iframe')),
    ]);

    await Promise.all(
      Array.from(frames).map(async (frame) => {
        this.cleanupFrameWindow(frame);
        if (frame.isConnected) {
          const loaded = new Promise<void>((resolve) => {
            frame.addEventListener('load', () => resolve(), { once: true });
          });
          frame.src = 'about:blank';
          await Promise.race([loaded, delay(BLANK_NAVIGATION_TIMEOUT_MS)]);
        }
        frame.remove();
      }),
    );
  }

  private cleanupFrameWindow(frame: HTMLIFrameElement): void {
    const frameWindow = frame.contentWindow;
    if (!frameWindow) return;

    try {
      for (const childFrame of Array.from(frameWindow.document.querySelectorAll<HTMLIFrameElement>('iframe'))) {
        this.cleanupFrameWindow(childFrame);
      }
    } catch {
      // Cross-origin or already-disposed frames are best-effort only.
    }

    try {
      for (let id = 0; id < 10_000; id += 1) {
        frameWindow.clearTimeout(id);
        frameWindow.clearInterval(id);
        frameWindow.cancelAnimationFrame?.(id);
        frameWindow.cancelIdleCallback?.(id);
      }
    } catch {
      // Ignore inaccessible frame windows.
    }

    try {
      for (const media of Array.from(frameWindow.document.querySelectorAll<HTMLMediaElement>('audio, video'))) {
        media.pause();
        media.removeAttribute('src');
        media.load();
      }
      frameWindow.document.body?.replaceChildren();
    } catch {
      // Ignore DOMs that are already gone.
    }
  }
}

export async function createOfficeEditor(
  container: HTMLElement,
  options: CreateOfficeEditorOptions,
): Promise<OfficeEditorInstance> {
  if (!(container instanceof HTMLElement)) {
    throw new Error('createOfficeEditor requires an HTMLElement container');
  }

  try {
    return await BrowserOfficeEditor.create(container, options);
  } catch (error) {
    const normalized = toError(error);
    options.onError?.(normalized);
    throw normalized;
  }
}

export function getActiveOfficeEditorCount(): number {
  return activeInstances.size;
}
