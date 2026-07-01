import { g_sEmpty_bin } from './empty_bin';
import { c_oAscFileType2, oAscFileType } from './file-types';
import { convertBinToDocument, convertDocument, convertPrintDataToPdf, initX2T } from './converter';
import { getDocumentType, getMimeTypeFromExtension, BASE_PATH } from './document-utils';
import { getOnlyOfficeLang, t } from './i18n';
import {
  createOnlyOfficeMockServer,
  LOCAL_ONLYOFFICE_USER_ID,
  LOCAL_ONLYOFFICE_USER_NAME,
} from './onlyoffice-mock-server';
import type { BinConversionResult, DocumentMediaMap, SaveEvent } from './document-types';
import { assertGeneratedFontAssetsAvailable } from './font-assets';

const ONLYOFFICE_BROWSER_BUILD_VERSION = '9.3.0';
const ONLYOFFICE_BROWSER_BUILD_NUMBER = 140;
const ONLYOFFICE_ZOOM_FIT_TO_WIDTH = -2;
const SAVE_TIMEOUT_MS = 60_000;
const BLANK_NAVIGATION_TIMEOUT_MS = 250;
const PRINT_RESOURCE_REVOKE_MS = 5 * 60_000;
const PRINT_PDF_CACHE_NAME = 'onlyoffice-browser-print-pdfs';
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const PRINT_SERVICE_WORKER_URL = '/document_editor_service_worker.js';
const PRINT_SERVICE_WORKER_READY_TIMEOUT_MS = 3_000;
const PRINT_TITLE_RESTORE_MS = 45_000;
const SUPPORTED_EMPTY_TYPES = ['docx', 'xlsx', 'pptx', 'csv'] as const;

type OfficeEmptyType = (typeof SUPPORTED_EMPTY_TYPES)[number];
type OfficeEditorStatus = 'opening' | 'ready' | 'destroyed' | 'error';
export type OfficeEditorMode = 'edit' | 'readonly' | 'preview';
export type OfficeEditorSourceKind = 'local-file' | 'new-document' | 'buffer' | 'url';
export type OfficeSaveBehavior = 'auto' | 'callback' | 'download';

export type OfficeEditorInput = Blob | ArrayBuffer | Uint8Array;
export type OfficeSaveCallbackResult = void | boolean;

export interface CreateOfficeEditorOptions {
  file?: File | Blob;
  buffer?: OfficeEditorInput;
  url?: string;
  emptyType?: OfficeEmptyType;
  fileName?: string;
  sourceKind?: OfficeEditorSourceKind;
  mode?: OfficeEditorMode;
  readonly?: boolean;
  spellcheck?: boolean;
  lang?: string;
  fetchOptions?: RequestInit;
  hardResetOnLastDestroy?: boolean;
  onReady?: (instance: OfficeEditorInstance) => void;
  saveBehavior?: OfficeSaveBehavior;
  onSave?: (file: File, instance: OfficeEditorInstance) => OfficeSaveCallbackResult | Promise<OfficeSaveCallbackResult>;
  onDirtyChange?: (dirty: boolean, instance: OfficeEditorInstance) => void | Promise<void>;
  onError?: (error: Error, instance?: OfficeEditorInstance) => void;
}

export interface OfficeEditorState {
  id: string;
  fileName: string;
  fileType: string;
  mode: OfficeEditorMode;
  readonly: boolean;
  dirty: boolean;
  sourceKind: OfficeEditorSourceKind;
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
  sourceKind: OfficeEditorSourceKind;
};

type SaveRequest = {
  targetExt: string;
  resolve: (file: File) => void;
  reject: (error: Error) => void;
  timeoutId: number;
  settled: boolean;
};

type NativeSaveData =
  | Uint8Array
  | ArrayBuffer
  | ArrayBufferView
  | {
      data?: unknown;
      header?: string;
    }
  | string;
type NativePrintPdfData =
  | Uint8Array
  | ArrayBuffer
  | ArrayBufferView
  | {
      data?: unknown;
    }
  | null
  | undefined;
type PrintPdfDataContainer = {
  data?: unknown;
  part?: unknown;
  [key: string]: unknown;
};
type PrintPdfCallback = (result: { type: 'save'; status: 'ok'; data: string; filetype: number } | null) => void;
type DocumentStateChangeEvent = boolean | { data?: boolean };
type DownloadAsEvent = { data?: { url?: string; fileType?: string | number; title?: string } };
type OnlyOfficeNativeBridge = {
  Save_End?: (header: string, length: number) => unknown;
  [key: string]: unknown;
};

type RuntimeWindow = typeof window & {
  DocsAPI?: DocsAPI;
  __fonts_visible_names?: unknown;
  __onlyOfficeBrowserSetPrintTitle?: (title: string, durationMs?: number) => void;
  native?: OnlyOfficeNativeBridge;
  APP?: {
    getImageURL?: (name: string, callback: (url: string) => void) => void;
    printPdf?: (dataContainer: PrintPdfDataContainer, callback: PrintPdfCallback) => void;
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
  __onlyOfficeBrowserPrintOpenGuard?: boolean;
};

let editorApiPromise: Promise<void> | null = null;
let nextEditorId = 1;
const activeInstances = new Map<string, BrowserOfficeEditor>();
const mediaRegistry = new Map<string, Record<string, string>>();
let temporaryDocumentTitleOriginal: string | null = null;
let temporaryDocumentTitleTimeout: number | null = null;

function createObjectUrl(blob: Blob): string {
  return URL.createObjectURL(blob);
}

function downloadFile(file: File): void {
  const objectUrl = createObjectUrl(file);
  const anchor = document.createElement('a');
  anchor.href = objectUrl;
  anchor.download = file.name || 'document';
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(objectUrl), 0);
}

function schedulePrintResourceCleanup(url: string): void {
  window.setTimeout(() => {
    if (url.startsWith('blob:')) {
      URL.revokeObjectURL(url);
      return;
    }
    void deleteCachedPrintPdf(url);
  }, PRINT_RESOURCE_REVOKE_MS);
}

function restoreTemporaryDocumentTitle(): void {
  if (temporaryDocumentTitleTimeout !== null) {
    window.clearTimeout(temporaryDocumentTitleTimeout);
    temporaryDocumentTitleTimeout = null;
  }
  if (temporaryDocumentTitleOriginal !== null) {
    document.title = temporaryDocumentTitleOriginal;
    temporaryDocumentTitleOriginal = null;
  }
  window.removeEventListener('afterprint', restoreTemporaryDocumentTitle);
}

function setTemporaryDocumentTitle(title: string, durationMs = PRINT_TITLE_RESTORE_MS): void {
  const normalizedTitle = title.trim();
  if (!normalizedTitle) return;

  if (temporaryDocumentTitleOriginal === null) {
    temporaryDocumentTitleOriginal = document.title;
  }
  document.title = normalizedTitle;

  if (temporaryDocumentTitleTimeout !== null) {
    window.clearTimeout(temporaryDocumentTitleTimeout);
  }
  window.removeEventListener('afterprint', restoreTemporaryDocumentTitle);
  window.addEventListener('afterprint', restoreTemporaryDocumentTitle, { once: true });
  temporaryDocumentTitleTimeout = window.setTimeout(restoreTemporaryDocumentTitle, Math.max(1_000, durationMs));
}

function preparePrintDocumentTitle(fileName: string): void {
  const title = fileName.replace(/\.pdf$/i, '').trim() || fileName;
  const setter = (window as RuntimeWindow).__onlyOfficeBrowserSetPrintTitle;
  if (setter) {
    setter(title, PRINT_TITLE_RESTORE_MS);
    return;
  }
  setTemporaryDocumentTitle(title, PRINT_TITLE_RESTORE_MS);
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

function replaceFileExtension(fileName: string, extension: string): string {
  const normalized = normalizeExtension(extension, getFileExtension(fileName) || 'docx').toLowerCase();
  return fileName.includes('.') ? fileName.replace(/\.[^/.]+$/, `.${normalized}`) : `${fileName}.${normalized}`;
}

function getExtensionFromUrl(url: string | undefined): string {
  if (!url) return '';
  try {
    const pathname = new URL(url, window.location.href).pathname;
    return getFileExtension(pathname);
  } catch {
    return '';
  }
}

function normalizeDownloadTargetExtension(value: string): string {
  const normalized = value.replace(/^\./, '').trim().toLowerCase();
  const aliases: Record<string, string> = {
    file_document_doc: 'doc',
    file_document_docx: 'docx',
    file_document_dotx: 'dotx',
    file_document_epub: 'epub',
    file_document_fb2: 'fb2',
    file_document_html: 'html',
    file_document_md: 'md',
    file_document_odt: 'odt',
    file_document_ott: 'ott',
    file_document_rtf: 'rtf',
    file_document_txt: 'txt',
    file_crossplatform_pdf: 'pdf',
    file_crossplatform_pdfa: 'pdfa',
    jpeg: 'jpg',
    pdf_a: 'pdfa',
    'pdf/a': 'pdfa',
  };
  return aliases[normalized] || normalized;
}

function getDownloadTargetExtension(
  fileType: string | number | undefined,
  url: string | undefined,
  fallback: string,
): string {
  if (typeof fileType === 'number' && Number.isFinite(fileType)) {
    const mapped = c_oAscFileType2[fileType];
    if (mapped) return normalizeDownloadTargetExtension(mapped);
  }
  if (typeof fileType === 'string' && fileType.trim()) {
    const numeric = Number(fileType);
    if (Number.isFinite(numeric)) {
      const mapped = c_oAscFileType2[numeric];
      if (mapped) return normalizeDownloadTargetExtension(mapped);
    }
    return normalizeDownloadTargetExtension(fileType);
  }
  return getExtensionFromUrl(url) || fallback;
}

function getDownloadFileName(sourceFileName: string, targetExt: string): string {
  const outputExtension = targetExt === 'pdfa' ? 'pdf' : targetExt === 'jpeg' ? 'jpg' : targetExt;
  return replaceFileExtension(sourceFileName, outputExtension);
}

function getDefaultFileName(emptyType: OfficeEmptyType): string {
  return `New_Document.${emptyType}`;
}

function getSavedFileMimeType(fileName: string): string {
  const extension = getFileExtension(fileName);
  const mimeMap: Record<string, string> = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
    epub: 'application/epub+zip',
    fb2: 'application/x-fictionbook+xml',
    html: 'text/html',
    md: 'text/markdown',
    odt: 'application/vnd.oasis.opendocument.text',
    ott: 'application/vnd.oasis.opendocument.text-template',
    rtf: 'application/rtf',
    txt: 'text/plain',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    csv: 'text/csv',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt: 'application/vnd.ms-powerpoint',
    pdf: 'application/pdf',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    png: 'image/png',
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
  if (Array.isArray(data)) {
    return Uint8Array.from(data.map((value) => Number(value) & 0xff));
  }
  if (typeof data === 'string') {
    const trimmed = data.trim();
    if (/^[A-Za-z0-9+/]+={0,2}$/.test(trimmed)) {
      try {
        const decoded = window.atob(trimmed);
        if (/^(DOCY|XLSY|PPTY|PK|%PDF)/.test(decoded)) {
          return Uint8Array.from(decoded, (char) => char.charCodeAt(0) & 0xff);
        }
      } catch {
        // Fall back to treating it as a binary string below.
      }
    }
    return Uint8Array.from(data, (char) => char.charCodeAt(0) & 0xff);
  }
  if (data && typeof data === 'object') {
    const record = data as Record<string, unknown>;
    if (record.data && record.data !== data) {
      return toUint8Array(record.data);
    }
    if (typeof record.length === 'number' && Number.isFinite(record.length)) {
      const length = Math.max(0, Math.floor(record.length));
      return Uint8Array.from({ length }, (_value, index) => Number(record[index]) & 0xff);
    }
    const numericKeys = Object.keys(record)
      .filter((key) => /^\d+$/.test(key))
      .sort((left, right) => Number(left) - Number(right));
    if (numericKeys.length > 0) {
      return Uint8Array.from(numericKeys.map((key) => Number(record[key]) & 0xff));
    }
  }
  throw new Error('Unsupported binary data type');
}

function toNativeSaveUint8Array(data: unknown): Uint8Array {
  if (typeof data === 'string') {
    return Uint8Array.from(data, (char) => char.charCodeAt(0) & 0xff);
  }
  if (data && typeof data === 'object') {
    const record = data as { data?: unknown };
    if (record.data && record.data !== data) {
      return toNativeSaveUint8Array(record.data);
    }
  }
  return toUint8Array(data);
}

function looksLikeZipPackage(data: Uint8Array): boolean {
  return data.length >= 4 && data[0] === 0x50 && data[1] === 0x4b;
}

function looksLikePdfDocument(data: Uint8Array): boolean {
  return data.length >= 4 && data[0] === 0x25 && data[1] === 0x50 && data[2] === 0x44 && data[3] === 0x46;
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

function getActivePrintEditor(): BrowserOfficeEditor | undefined {
  return Array.from(activeInstances.values()).find((instance) => !instance.getState().destroyed);
}

function getPrintResourceUrl(fileName: string): string {
  const randomId =
    typeof window.crypto?.randomUUID === 'function'
      ? window.crypto.randomUUID()
      : `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  const safeFileName = fileName.replace(/[\\/:*?"<>|]+/g, '_') || 'document.pdf';
  const pdfFileName = safeFileName.toLowerCase().endsWith('.pdf')
    ? safeFileName
    : replaceFileExtension(safeFileName, 'pdf');
  const url = new URL(`${PRINT_PDF_ROUTE_PREFIX}${randomId}/${encodeURIComponent(pdfFileName)}`, window.location.origin);
  url.searchParams.set('filename', pdfFileName);
  return url.href;
}

async function ensurePrintPdfServiceWorker(): Promise<void> {
  if (!('serviceWorker' in navigator)) {
    throw new Error('OnlyOffice print requires Service Worker support for same-origin PDF delivery');
  }

  await navigator.serviceWorker.register(PRINT_SERVICE_WORKER_URL, { scope: '/' });
  await navigator.serviceWorker.ready;

  if (navigator.serviceWorker.controller) return;

  await new Promise<void>((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      navigator.serviceWorker.removeEventListener('controllerchange', onControllerChange);
      reject(new Error('OnlyOffice print service worker did not take control of the page'));
    }, PRINT_SERVICE_WORKER_READY_TIMEOUT_MS);

    const onControllerChange = () => {
      window.clearTimeout(timeoutId);
      resolve();
    };

    navigator.serviceWorker.addEventListener('controllerchange', onControllerChange, { once: true });
  });
}

function escapeHeaderQuotedString(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/"/g, '\\"')
    .replace(/[\r\n]+/g, ' ');
}

function encodeRfc5987Value(value: string): string {
  return encodeURIComponent(value)
    .replace(/['()]/g, (char) => `%${char.charCodeAt(0).toString(16).toUpperCase()}`)
    .replace(/\*/g, '%2A');
}

function getAsciiHeaderFileNameFallback(fileName: string): string {
  const fallback = fileName
    .replace(/[^\x20-\x7e]+/g, '_')
    .replace(/[\\/:*?"<>|]+/g, '_')
    .trim();
  return fallback || 'document.pdf';
}

function getPdfTitleFromFileName(fileName: string): string {
  const safeFileName = fileName.toLowerCase().endsWith('.pdf') ? fileName : replaceFileExtension(fileName, 'pdf');
  return safeFileName.replace(/\.pdf$/i, '').trim() || 'document';
}

function bytesToBinaryString(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  let output = '';
  for (let index = 0; index < bytes.length; index += chunkSize) {
    output += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }
  return output;
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (char) => char.charCodeAt(0) & 0xff);
}

function encodePdfUtf16BeHex(value: string): string {
  const bytes = [0xfe, 0xff];
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes.push((code >> 8) & 0xff, code & 0xff);
  }
  return bytes.map((byte) => byte.toString(16).padStart(2, '0').toUpperCase()).join('');
}

function appendPrintPdfTitleMetadata(bytes: Uint8Array, fileName: string): Uint8Array {
  const source = bytesToBinaryString(bytes);
  const startxrefMatches = Array.from(source.matchAll(/startxref\s+(\d+)\s+%%EOF/g));
  const startxrefMatch = startxrefMatches.at(-1);
  const previousXrefOffset = startxrefMatch ? Number(startxrefMatch[1]) : NaN;
  const rootMatch = source.match(/\/Root\s+(\d+)\s+(\d+)\s+R/);
  const sizeMatches = Array.from(source.matchAll(/\/Size\s+(\d+)/g));
  const sizeMatch = sizeMatches.at(-1);

  if (!Number.isFinite(previousXrefOffset) || !rootMatch || !sizeMatch) {
    return bytes;
  }

  const objectNumber = Number(sizeMatch[1]);
  const titleHex = encodePdfUtf16BeHex(getPdfTitleFromFileName(fileName));
  const objectOffset = bytes.byteLength;
  const object = `\n${objectNumber} 0 obj\n<< /Title <${titleHex}> >>\nendobj\n`;
  const xrefOffset = objectOffset + asciiBytes(object).byteLength;
  const xref = `xref\n${objectNumber} 1\n${String(objectOffset).padStart(10, '0')} 00000 n \n`;
  const trailer =
    `trailer\n<< /Size ${objectNumber + 1} /Root ${rootMatch[1]} ${rootMatch[2]} R ` +
    `/Info ${objectNumber} 0 R /Prev ${previousXrefOffset} >>\n` +
    `startxref\n${xrefOffset}\n%%EOF\n`;
  const append = asciiBytes(`${object}${xref}${trailer}`);
  const output = new Uint8Array(bytes.byteLength + append.byteLength);
  output.set(bytes, 0);
  output.set(append, bytes.byteLength);
  return output;
}

async function putCachedPrintPdf(url: string, bytes: Uint8Array, fileName: string): Promise<void> {
  if (!('caches' in window)) {
    throw new Error('OnlyOffice print requires Cache API support for same-origin PDF delivery');
  }

  await ensurePrintPdfServiceWorker();
  const safeFileName = fileName.toLowerCase().endsWith('.pdf') ? fileName : replaceFileExtension(fileName, 'pdf');
  const asciiFileName = getAsciiHeaderFileNameFallback(safeFileName);
  const encodedFileName = encodeRfc5987Value(safeFileName);
  const titledBytes = appendPrintPdfTitleMetadata(bytes, safeFileName);
  const cache = await window.caches.open(PRINT_PDF_CACHE_NAME);
  await cache.put(
    url,
    new Response(titledBytes as BodyInit, {
      headers: {
        'accept-ranges': 'bytes',
        'cache-control': 'no-store',
        'content-disposition': `inline; filename="${escapeHeaderQuotedString(
          asciiFileName,
        )}"; filename*=UTF-8''${encodedFileName}`,
        'content-length': String(titledBytes.byteLength),
        'content-type': 'application/pdf',
        'x-content-type-options': 'nosniff',
      },
    }),
  );
}

async function deleteCachedPrintPdf(url: string): Promise<void> {
  if (!('caches' in window)) return;
  await window.caches
    .open(PRINT_PDF_CACHE_NAME)
    .then((cache) => cache.delete(url))
    .catch(() => undefined);
}

async function createPrintPdfUrl(dataContainer?: PrintPdfDataContainer): Promise<string> {
  const activeEditor = getActivePrintEditor();
  const data = dataContainer?.data || dataContainer?.part;
  if (!data) {
    throw new Error('Print request did not include document data');
  }

  const sourceFileName = activeEditor?.getState().fileName || 'document.docx';
  const media = activeEditor?.getMedia();
  const sourceBytes = toUint8Array(data);
  let bytes: Uint8Array;
  let printFileName: string;
  const nativePrintBytes = activeEditor?.getNativePrintData();
  if (nativePrintBytes && looksLikePdfDocument(nativePrintBytes)) {
    bytes = nativePrintBytes;
    printFileName = replaceFileExtension(sourceFileName, 'pdf');
  } else if (nativePrintBytes) {
    const result = await convertPrintDataToPdf(nativePrintBytes, sourceFileName, media);
    bytes = toUint8Array(result.data);
    printFileName = result.fileName || replaceFileExtension(sourceFileName, 'pdf');
  } else if (looksLikePdfDocument(sourceBytes)) {
    bytes = sourceBytes;
    printFileName = replaceFileExtension(sourceFileName, 'pdf');
  } else {
    const result = await convertBinToDocument(sourceBytes, sourceFileName, 'PDF', media);
    bytes = toUint8Array(result.data);
    printFileName = result.fileName || replaceFileExtension(sourceFileName, 'pdf');
  }
  const pdfUrl = getPrintResourceUrl(printFileName);
  preparePrintDocumentTitle(printFileName);
  await putCachedPrintPdf(pdfUrl, bytes, printFileName);
  activeEditor?.trackPrintResourceUrl(pdfUrl);
  schedulePrintResourceCleanup(pdfUrl);
  return pdfUrl;
}

function installOnlyOfficeParentApp(): void {
  const runtimeWindow = window as RuntimeWindow;
  runtimeWindow.APP = runtimeWindow.APP || {};
  runtimeWindow.APP.getImageURL = (name, callback) => {
    callback(resolveMediaFromRegistry(name));
  };
  runtimeWindow.APP.printPdf = (dataContainer, callback) => {
    void createPrintPdfUrl(dataContainer)
      .then((printUrl) => {
        callback({ type: 'save', status: 'ok', data: printUrl, filetype: oAscFileType.PDF });
      })
      .catch((error) => {
        console.error('OnlyOffice print PDF export failed:', error);
        callback(null);
      });
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

function applyFillContainerDefaults(element: HTMLElement): void {
  const { style } = element;
  style.width ||= '100%';
  style.height ||= '100%';
  style.minWidth ||= '0';
  style.minHeight ||= '0';
}

function applyOnlyOfficeFrameDefaults(iframe: HTMLIFrameElement): void {
  applyFillContainerDefaults(iframe);
  iframe.style.display ||= 'block';
  iframe.style.border ||= '0';
}

function isPrintPdfUrl(value: unknown, baseUrl: string): boolean {
  if (!value) return false;
  try {
    return new URL(String(value), baseUrl).pathname.startsWith(PRINT_PDF_ROUTE_PREFIX);
  } catch {
    return false;
  }
}

function installOnlyOfficePrintOpenGuard(iframe: HTMLIFrameElement): void {
  const patchWindowOpen = () => {
    try {
      const frameWindow = iframe.contentWindow as OnlyOfficeFrameWindow | null;
      if (!frameWindow || frameWindow.__onlyOfficeBrowserPrintOpenGuard) return;

      const originalOpen = frameWindow.open.bind(frameWindow);
      frameWindow.open = ((url?: string | URL, target?: string, features?: string) => {
        if (isPrintPdfUrl(url, frameWindow.location.href)) {
          console.warn('Suppressed OnlyOffice print fallback popup:', String(url));
          return null;
        }
        return originalOpen(url as string | undefined, target, features);
      }) as typeof frameWindow.open;
      frameWindow.__onlyOfficeBrowserPrintOpenGuard = true;
    } catch {
      // The editor frame is expected to be same-origin. If it is not ready yet,
      // the load listener below will retry after navigation completes.
    }
  };

  patchWindowOpen();
  iframe.addEventListener('load', patchWindowOpen);
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
        sourceKind: 'new-document',
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
      sourceKind: 'new-document',
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
      fileName,
      fileType: getFileExtension(fileName),
      binData: converted.bin,
      media: converted.media,
      sourceObjectUrl,
      originalFile: file,
      sourceKind: 'url',
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
    fileName,
    fileType: getFileExtension(fileName),
    binData: converted.bin,
    media: converted.media,
    sourceObjectUrl,
    originalFile: file,
    sourceKind: options.sourceKind || (options.file ? 'local-file' : 'buffer'),
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
  private printResourceUrls = new Set<string>();
  private sourceObjectUrl: string | null = null;
  private status: OfficeEditorStatus = 'opening';
  private dirty = false;
  private destroyed = false;
  private readonly fileName: string;
  private readonly fileType: string;
  private readonly originalFile?: File;
  private readonly sourceKind: OfficeEditorSourceKind;
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
    this.sourceKind = prepared.sourceKind;
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
    applyFillContainerDefaults(this.container);
    this.placeholder.id = `${this.id}-frame`;
    this.placeholder.className = 'office-editor-frame';
    applyFillContainerDefaults(this.placeholder);
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
          coEditing: {
            mode: 'strict',
            change: false,
          },
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
            autosave: false,
            forcesave: false,
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
            this.applyNestedEditorFrameDefaults();
            installNestedFontPickerFilter();
          },
          onDocumentReady: () => {
            if (this.destroyed) return;
            this.applyNestedEditorFrameDefaults();
            installNestedFontPickerFilter();
            this.status = 'ready';
            this.applyDefaultEditorModePreviewZoom();
            this.options.onReady?.(this);
          },
          onSave: (event: SaveEvent) => {
            void this.handleSaveDocument(event);
          },
          onDocumentStateChange: (event: DocumentStateChangeEvent) => {
            this.handleDocumentStateChange(event);
          },
          onDownloadAs: (event: DownloadAsEvent) => {
            void this.handleDownloadAs(event);
          },
          writeFile: (event: any) => {
            void this.handleWriteFile(event);
          },
        },
      });

      this.editor = editor;
      this.applyNestedEditorFrameDefaults();
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
          onSaveRequest: () =>
            this.saveNativeBinary(
              this.fileName.toLowerCase().endsWith('.csv') ? 'CSV' : getFileExtension(this.fileName).toUpperCase(),
            ),
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

  private applyNestedEditorFrameDefaults(): void {
    const frame = this.placeholder.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    if (frame) {
      applyOnlyOfficeFrameDefaults(frame);
      installOnlyOfficePrintOpenGuard(frame);
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
    const sourceBytes = toUint8Array(data);
    try {
      const result: BinConversionResult = await convertBinToDocument(
        sourceBytes,
        this.fileName,
        targetFormat,
        this.media,
      );
      const bytes = toUint8Array(result.data);
      return new File([bytes as BlobPart], result.fileName, { type: getSavedFileMimeType(result.fileName) });
    } catch (error) {
      if (!looksLikeZipPackage(sourceBytes)) {
        const header = Array.from(sourceBytes.slice(0, 16))
          .map((byte) => byte.toString(16).padStart(2, '0'))
          .join(' ');
        throw new Error(`${toError(error).message}; native header: ${header}`);
      }
      const fileName = replaceFileExtension(this.fileName, targetFormat);
      return new File([sourceBytes as BlobPart], fileName, { type: getSavedFileMimeType(fileName) });
    }
  }

  private getSaveBehavior(): OfficeSaveBehavior {
    return this.options.saveBehavior || 'auto';
  }

  private async invokeSaveCallback(file: File): Promise<boolean> {
    const behavior = this.getSaveBehavior();
    const shouldCallCallback =
      behavior === 'callback' ||
      (behavior === 'auto' && (this.sourceKind !== 'new-document' || Boolean(this.options.onSave)));

    if (!shouldCallCallback) return false;
    if (!this.options.onSave) {
      throw new Error(
        'A save callback is required for this document source. Provide onSave or use saveBehavior: "download".',
      );
    }

    const handled = await this.options.onSave(file, this);
    return handled === true;
  }

  private async persistSavedFile(file: File): Promise<void> {
    const behavior = this.getSaveBehavior();
    const handledByCallback = await this.invokeSaveCallback(file);

    if (behavior === 'download' || (behavior === 'auto' && this.sourceKind === 'new-document' && !handledByCallback)) {
      downloadFile(file);
    }
  }

  private async emitSavedFile(file: File): Promise<void> {
    await this.persistSavedFile(file);
    this.setDirty(false);
    this.resolveSaveRequest(file);
  }

  private async handleSaveDocument(event: SaveEvent): Promise<void> {
    try {
      const payload = event.data?.data?.data;
      const nativeData = payload ? toNativeSaveUint8Array(payload) : this.getNativeSaveData();
      const outputFormat = event.data?.option?.outputformat;
      const optionFormat =
        (typeof outputFormat === 'number' ? c_oAscFileType2[outputFormat] : undefined) ||
        getFileExtension(this.fileName).toUpperCase();
      const targetFormat = this.fileName.toLowerCase().endsWith('.csv')
        ? 'CSV'
        : this.saveRequest?.targetExt || optionFormat;
      const file = await this.convertSavedBin(nativeData, targetFormat);
      await this.emitSavedFile(file);
      this.sendOnlyOfficeCommand('asc_onSaveCallback', { err_code: 0 });
    } catch (error) {
      const normalized = toError(error);
      this.rejectSaveRequest(normalized);
      this.options.onError?.(normalized, this);
      this.sendOnlyOfficeCommand('asc_onSaveCallback', { err_code: 1 });
    }
  }

  private setDirty(dirty: boolean): void {
    if (this.destroyed || this.dirty === dirty) return;
    this.dirty = dirty;
    void Promise.resolve(this.options.onDirtyChange?.(dirty, this)).catch((error) => {
      this.options.onError?.(toError(error), this);
    });
  }

  private handleDocumentStateChange(event: DocumentStateChangeEvent): void {
    const dirty = typeof event === 'boolean' ? event : Boolean(event?.data);
    this.setDirty(dirty);
  }

  private async handleDownloadAs(event: DownloadAsEvent): Promise<void> {
    try {
      const url = event.data?.url;
      const targetExt = getDownloadTargetExtension(event.data?.fileType, url, getFileExtension(this.fileName));
      if (!url) {
        const file = await this.convertSavedBin(this.getNativeSaveData(), targetExt.toUpperCase());
        downloadFile(file);
        return;
      }
      const response = await fetch(url);
      if (!response.ok) {
        const file = await this.convertSavedBin(this.getNativeSaveData(), targetExt.toUpperCase());
        downloadFile(file);
        return;
      }

      const blob = await response.blob();
      const fileName = event.data?.title || getDownloadFileName(this.fileName, targetExt);
      downloadFile(new File([blob], fileName, { type: blob.type || getSavedFileMimeType(fileName) }));
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

  private getNativeSaveData(): Uint8Array {
    if (!this.editor || typeof this.editor.asc_nativeGetFile3 !== 'function') {
      throw new Error('The current editor does not support native binary export');
    }

    const payload: NativeSaveData | undefined = this.editor.asc_nativeGetFile3();
    const data = payload && typeof payload === 'object' && 'data' in payload ? payload.data : payload;
    if (!data) {
      throw new Error('Native binary export did not include document data');
    }
    return toNativeSaveUint8Array(data);
  }

  getNativePrintData(): Uint8Array | null {
    if (!this.editor || typeof this.editor.asc_nativeGetPDF !== 'function') return null;

    const runtimeWindow =
      (this.editor.getEditorWindow?.() as RuntimeWindow | null | undefined) || (window as RuntimeWindow);
    const hadNative = Object.prototype.hasOwnProperty.call(runtimeWindow, 'native');
    const nativeBridge = runtimeWindow.native || {};
    const hadSaveEnd = Object.prototype.hasOwnProperty.call(nativeBridge, 'Save_End');
    const originalSaveEnd = nativeBridge.Save_End;
    let nativeLength: number | undefined;

    runtimeWindow.native = nativeBridge;
    nativeBridge.Save_End = (header, length) => {
      if (Number.isFinite(length)) {
        nativeLength = Math.max(0, Math.floor(length));
      }
      return originalSaveEnd?.call(nativeBridge, header, length);
    };

    try {
      const payload: NativePrintPdfData = this.editor.asc_nativeGetPDF({ isPrint: true });
      const data = payload && typeof payload === 'object' && 'data' in payload ? payload.data : payload;
      if (!data) return null;

      let bytes = toUint8Array(data);
      if (nativeLength !== undefined && nativeLength > 0 && nativeLength <= bytes.byteLength) {
        bytes = bytes.slice(0, nativeLength);
      }
      return bytes.byteLength > 0 ? bytes : null;
    } catch (error) {
      console.warn('OnlyOffice native PDF print export failed:', error);
      return null;
    } finally {
      if (hadSaveEnd) {
        nativeBridge.Save_End = originalSaveEnd;
      } else {
        delete nativeBridge.Save_End;
      }
      if (!hadNative) {
        delete runtimeWindow.native;
      }
    }
  }

  getMedia(): DocumentMediaMap {
    return this.media;
  }

  trackPrintResourceUrl(url: string): void {
    this.printResourceUrls.add(url);
  }

  private async saveNativeBinary(targetFormat: string): Promise<void> {
    try {
      const file = await this.convertSavedBin(this.getNativeSaveData(), targetFormat);
      await this.emitSavedFile(file);
    } catch (error) {
      const normalized = toError(error);
      this.rejectSaveRequest(normalized);
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
    if (typeof this.editor.asc_nativeGetFile3 !== 'function') {
      return Promise.reject(new Error('The current editor does not support native binary export'));
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

      void this.saveNativeBinary(normalizedTargetExt);
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
      dirty: this.dirty,
      sourceKind: this.sourceKind,
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
    for (const url of this.printResourceUrls) {
      if (url.startsWith('blob:')) {
        URL.revokeObjectURL(url);
      } else {
        void deleteCachedPrintPdf(url);
      }
    }
    this.printResourceUrls.clear();

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
