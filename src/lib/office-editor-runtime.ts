import { g_sEmpty_bin } from './empty_bin';
import { c_oAscFileType2, oAscFileType } from './file-types';
import {
  convertBinToDocument,
  convertDocument,
  convertHtmlToDocument,
  convertPdfToImage,
  convertPrintDataToPdf,
  initX2T,
} from './converter';
import { getDocumentType, getMimeTypeFromExtension, BASE_PATH } from './document-utils';
import { getOnlyOfficeLang, t } from './i18n';
import {
  createOnlyOfficeMockServer,
  isOnlyOfficeDirectMediaUrl,
  LOCAL_ONLYOFFICE_USER_ID,
  LOCAL_ONLYOFFICE_USER_NAME,
} from './onlyoffice-mock-server';
import type { BinConversionResult, DocumentMediaMap, SaveEvent } from './document-types';
import { assertGeneratedFontAssetsAvailable, resolveAvailableFontFamilyNames } from './font-assets';

const ONLYOFFICE_BROWSER_BUILD_VERSION = '9.3.0';
const ONLYOFFICE_BROWSER_BUILD_NUMBER = 140;
const ONLYOFFICE_LATIN_DEFAULT_FONT_FAMILY = 'Aptos';
const ONLYOFFICE_EAST_ASIA_DEFAULT_FONT_FAMILY = 'DengXian';
const ONLYOFFICE_ZOOM_FIT_TO_WIDTH = -2;
const SAVE_TIMEOUT_MS = 60_000;
const BLANK_NAVIGATION_TIMEOUT_MS = 250;
const PRINT_RESOURCE_REVOKE_MS = 5 * 60_000;
const PRINT_PDF_CACHE_NAME = 'onlyoffice-browser-print-pdfs';
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const PRINT_SERVICE_WORKER_URL = '/document_editor_service_worker.js';
const PRINT_SERVICE_WORKER_READY_TIMEOUT_MS = 3_000;
const PRINT_TITLE_RESTORE_MS = 45_000;
const ONLYOFFICE_PRINT_DOWNLOAD_TYPE = 'asc_onPrintUrl';
const ONLYOFFICE_PRINT_ACTION_TYPE = 7;
const SUPPORTED_EMPTY_TYPES = ['docx', 'xlsx', 'pptx', 'csv'] as const;
const SPREADSHEET_PDF_EXPORT_FILE_TYPES = new Set<number>([oAscFileType.PDF, oAscFileType.PDFA]);
const ZIP_DOWNLOAD_EXTENSIONS = new Set([
  'docm',
  'docx',
  'docxf',
  'dotx',
  'epub',
  'oform',
  'odp',
  'ods',
  'odt',
  'otp',
  'ots',
  'ott',
  'potm',
  'potx',
  'ppsm',
  'ppsx',
  'pptm',
  'pptx',
  'xlsb',
  'xlsm',
  'xlsx',
  'xltm',
  'xltx',
  'zip',
]);
const NATIVE_HTML_DOWNLOAD_EXTENSIONS = new Set(['html', 'md', 'fb2', 'epub']);
const RETURN_PREVIEW_ICON_NAME = 'qlementine-icons:preview-16';
const RETURN_PREVIEW_ICON_BODY =
  '<path fill="currentColor" fill-rule="evenodd" d="M12 4.57a.5.5 0 0 0-.024-.235l-.013-.063a1.5 1.5 0 0 0-.18-.434c-.092-.15-.222-.28-.482-.54L8.711.707c-.259-.26-.389-.39-.54-.483a1.5 1.5 0 0 0-.496-.193a.5.5 0 0 0-.235-.024C7.329.004 7.194.004 7.015.004h-2.21c-1.68 0-2.52 0-3.16.327a3.02 3.02 0 0 0-1.31 1.31C.008 2.283.008 3.12.008 4.8v6.4c0 1.68 0 2.52.327 3.16a3.02 3.02 0 0 0 1.31 1.31c.642.327 1.48.327 3.16.327h2.423c.401 0 .602-.523.347-.832a.45.45 0 0 0-.345-.168H4.8c-.857 0-1.44-.001-1.89-.038c-.438-.036-.663-.1-.819-.18a2 2 0 0 1-.874-.874c-.08-.156-.145-.38-.18-.819c-.036-.45-.037-1.03-.037-1.89v-6.4c0-.857 0-1.44.037-1.89c.036-.438.101-.663.18-.819c.192-.376.498-.682.874-.874c.156-.08.381-.145.82-.18C3.36.997 3.94.997 4.8.997H7v3.5a.5.5 0 0 0 .5.5H11v.547c0 .25.207.45.456.473c.285.025.543-.188.543-.474V4.99c0-.178 0-.313-.005-.425zM8 1.41L10.59 4H8z" clip-rule="evenodd"/><path fill="currentColor" fill-rule="evenodd" d="M11 15c.834 0 1.61-.255 2.25-.691l1.47 1.47a.749.749 0 1 0 1.06-1.06l-1.47-1.47c.436-.641.691-1.41.691-2.25c0-2.21-1.79-4-4-4s-4 1.79-4 4s1.79 4 4 4zm0-1c1.66 0 3-1.34 3-3s-1.34-3-3-3s-3 1.34-3 3s1.34 3 3 3" clip-rule="evenodd"/>';
const RETURN_PREVIEW_AVATAR_ANCHOR_IDS = [
  'slot-btn-user',
  'slot-btn-users',
  'slot-btn-user-info',
  'slot-btn-user-menu',
  'slot-btn-coauth',
  'slot-btn-profile',
  'id-btn-user',
  'id-btn-users',
] as const;

type OfficeEmptyType = (typeof SUPPORTED_EMPTY_TYPES)[number];
type OfficeEditorStatus = 'opening' | 'ready' | 'destroyed' | 'error';
export type OfficeEditorMode = 'edit' | 'readonly' | 'preview';
export type OfficeEditorSourceKind = 'local-file' | 'new-document' | 'buffer' | 'url';
export type OfficeSaveBehavior = 'auto' | 'callback' | 'download';
export type OfficeInterfaceTheme = 'system' | 'light' | 'dark';
export type OfficePluginOptions = {
  configUrls: string[];
  autostart?: string[];
};
type OnlyOfficeUiTheme = 'theme-system' | 'theme-white' | 'theme-night';
export type OfficeSaveToNewFormatConfirmationOptions = {
  title?: string;
  message?: string;
  dontshow?: boolean;
};

export type OfficeEditorInput = Blob | ArrayBuffer | Uint8Array;
export type OfficeSaveCallbackResult = void | boolean;
export type OfficeSaveAsCallbackResult = void | boolean;
export type OfficeDownloadCallbackResult = void;

export interface CreateOfficeEditorOptions {
  file?: File | Blob;
  buffer?: OfficeEditorInput;
  url?: string;
  emptyType?: OfficeEmptyType;
  fileName?: string;
  sourceKind?: OfficeEditorSourceKind;
  mode?: OfficeEditorMode;
  readonly?: boolean;
  /** Preserve the viewer origin when restoring a document directly in edit mode. */
  canReturnToPreview?: boolean;
  spellcheck?: boolean;
  interfaceTheme?: OfficeInterfaceTheme;
  lang?: string;
  plugins?: OfficePluginOptions;
  /** Picker-visible families derived from font files verified by the host. */
  visibleFontNames?: string[];
  /**
   * @deprecated Retained for source compatibility. Missing-font substitution
   * is always delegated to ONLYOFFICE's generated native font dictionary.
   */
  fallbackFontNames?: string[];
  fetchOptions?: RequestInit;
  hardResetOnLastDestroy?: boolean;
  onReady?: (instance: OfficeEditorInstance) => void;
  saveBehavior?: OfficeSaveBehavior;
  onSave?: (file: File, instance: OfficeEditorInstance) => OfficeSaveCallbackResult | Promise<OfficeSaveCallbackResult>;
  onSaveAs?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeSaveAsCallbackResult | Promise<OfficeSaveAsCallbackResult>;
  onDownload?: (
    file: File,
    instance: OfficeEditorInstance,
  ) => OfficeDownloadCallbackResult | Promise<OfficeDownloadCallbackResult>;
  onDirtyChange?: (dirty: boolean, instance: OfficeEditorInstance) => void | Promise<void>;
  onStateChange?: (state: OfficeEditorState, instance: OfficeEditorInstance) => void | Promise<void>;
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
  confirmSaveToNewFormat(options?: OfficeSaveToNewFormatConfirmationOptions): Promise<boolean>;
  setInterfaceTheme(theme: OfficeInterfaceTheme): void;
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
type NativeHtmlData = string | null | undefined;
type PrintPdfDataContainer = {
  data?: unknown;
  part?: unknown;
  [key: string]: unknown;
};
type PrintPdfCallback = (result: { type: 'save'; status: 'ok'; data: string; filetype: number } | null) => void;
type DocumentStateChangeEvent = boolean | { data?: boolean };
type DownloadAsEvent = {
  data?: { url?: string; fileType?: string | number; title?: string };
  nativeOptions?: NativeDownloadAsOptions;
  additionalData?: NativeDownloadAsAdditionalData;
  dataContainer?: NativeDownloadAsDataContainer;
  actionType?: unknown;
  downloadType?: unknown;
};
type RequestSaveAsEvent = DownloadAsEvent;
type NativeDownloadAsOptions = {
  fileType?: unknown;
  format?: unknown;
  outputformat?: unknown;
  outputFormat?: unknown;
  type?: unknown;
  asc_getFileType?: () => unknown;
  getFileType?: () => unknown;
  get_FileType?: () => unknown;
  asc_getFormat?: () => unknown;
  getFormat?: () => unknown;
  get_Format?: () => unknown;
  advancedOptions?: unknown;
  textParams?: unknown;
  printOptions?: unknown;
  isSaveAs?: unknown;
  asc_getIsSaveAs?: () => unknown;
  getIsSaveAs?: () => unknown;
  get_IsSaveAs?: () => unknown;
  wopiSaveAsPath?: unknown;
  isDownloadEvent?: unknown;
  isPdfPrint?: unknown;
  isPrint?: unknown;
};
type NativeDownloadAsAdditionalData = {
  outputformat?: unknown;
  title?: unknown;
  inline?: unknown;
  withoutPassword?: unknown;
  jsonparams?: unknown;
  textParams?: unknown;
  thumbnail?: unknown;
  isSaveAs?: unknown;
  saveAsPath?: unknown;
  [key: string]: unknown;
};
type NativeDownloadAsDataContainer = {
  data?: unknown;
  part?: unknown;
  index?: unknown;
  count?: unknown;
  [key: string]: unknown;
};
type NativeDownloadAsHandler = (
  actionType?: unknown,
  options?: NativeDownloadAsOptions,
  additionalData?: NativeDownloadAsAdditionalData,
  dataContainer?: NativeDownloadAsDataContainer,
  downloadType?: unknown,
) => boolean | undefined;
type NativeAscDownloadAsHandler = (options?: NativeDownloadAsOptions | string | number) => unknown;
type SpreadsheetPdfPrintExportMode = {
  fileType: number;
  isSaveAs: boolean;
  returnPanel: 'saveas' | 'save-copy';
  wopiSaveAsPath?: unknown;
};
type SpreadsheetPdfPrintPanelState = {
  printPanel: HTMLElement | null;
  sourcePanel: HTMLElement | null;
  previousPrintDisplay: string;
  previousSourceDisplay: string;
};
type OnlyOfficeNotificationCenter = {
  trigger?: (...args: unknown[]) => unknown;
  __onlyOfficeBrowserSpreadsheetPdfPrintTriggerPatched?: boolean;
  __onlyOfficeBrowserSpreadsheetPdfPrintTriggerOriginal?: (...args: unknown[]) => unknown;
};
type OnlyOfficeDownloadOptionsLike = {
  asc_setAdvancedOptions?: (options: unknown) => unknown;
  asc_setIsSaveAs?: (isSaveAs: boolean) => unknown;
  asc_setWopiSaveAsPath?: (path: unknown) => unknown;
};
type OnlyOfficeButtonLike = {
  setCaption?: (caption: string) => unknown;
  updateCaption?: (caption: string) => unknown;
  el?: unknown;
  $el?: unknown;
  btnEl?: unknown;
  cmpEl?: unknown;
};
type OnlyOfficePrintSettingsLike = {
  show?: () => unknown;
  txtPrint?: unknown;
  btnsPrint?: unknown[];
  btnsSave?: unknown[];
  btnPrintSystemDialog?: unknown;
  applySettings?: () => unknown;
  getRange?: () => unknown;
  getIgnorePrintArea?: () => unknown;
  getPagesFrom?: () => unknown;
  getPagesTo?: () => unknown;
  cmbPaperOrientation?: { getSelectedRecord?: () => { value?: unknown } | null };
  cmbPrinter?: { getSelectedRecord?: () => { value?: unknown } | null };
  cmbColorPrinting?: { getValue?: () => unknown };
  spnCopies?: { getNumberValue?: () => unknown };
  cmbSides?: { getValue?: () => unknown };
  printScroller?: { update?: (options?: unknown) => unknown };
  __onlyOfficeBrowserSpreadsheetPdfPrintShowPatched?: boolean;
  __onlyOfficeBrowserSpreadsheetPdfPrintShowOriginal?: () => unknown;
};
type OnlyOfficePrintControllerLike = {
  printSettings?: OnlyOfficePrintSettingsLike;
  adjPrintParams?: {
    asc_setPrintType?: (value: unknown) => unknown;
    asc_setPageOptionsMap?: (value: unknown) => unknown;
    asc_setIgnorePrintArea?: (value: unknown) => unknown;
    asc_setActiveSheetsArray?: (value: unknown) => unknown;
    asc_setStartPageIndex?: (value: unknown) => unknown;
    asc_setEndPageIndex?: (value: unknown) => unknown;
    asc_setNativeOptions?: (value: unknown) => unknown;
  };
  _changedProps?: unknown[] | Record<string, unknown>;
  api?: {
    asc_getActiveWorksheetIndex?: () => number;
    asc_getPageOptions?: (index: number, keepOriginal?: boolean, init?: boolean) => unknown;
    asc_DownloadAs?: (options: unknown) => unknown;
  };
  savePageOptions?: (settings: OnlyOfficePrintSettingsLike) => unknown;
  findPagePreset?: (settings: OnlyOfficePrintSettingsLike, width: number, height: number) => unknown;
  querySavePrintSettings?: (action: unknown, useSystemDialog?: unknown) => unknown;
  onHidePrintMenu?: () => unknown;
  updatePrintRenderContainerSize?: (redraw?: boolean) => unknown;
  __onlyOfficeBrowserSpreadsheetPdfPrintQueryPatched?: boolean;
  __onlyOfficeBrowserSpreadsheetPdfPrintQueryOriginal?: (action: unknown, useSystemDialog?: unknown) => unknown;
};
type OnlyOfficeEditorApi = {
  asc_DownloadAs?: NativeAscDownloadAsHandler;
  _downloadAs?: NativeDownloadAsHandler;
  asc_nativeCalculateFile?: (options?: Record<string, unknown> | null) => unknown;
  asc_nativeGetHtml?: (options?: Record<string, unknown> | null) => NativeHtmlData;
  __onlyOfficeBrowserDownloadAsPatched?: boolean;
  __onlyOfficeBrowserAscDownloadAsOriginal?: NativeAscDownloadAsHandler;
  __onlyOfficeBrowserDownloadAsOriginal?: NativeDownloadAsHandler;
  [key: string]: unknown;
};
type OnlyOfficeNativeBridge = {
  Save_End?: (header: string, length: number) => unknown;
  [key: string]: unknown;
};
type OnlyOfficeThemeController = {
  map?: () => Record<string, unknown>;
  get?: (theme: string) => unknown;
  setTheme?: (theme: OnlyOfficeUiTheme, source?: string) => unknown;
  defaultThemeId?: () => string;
  defaultTheme?: () => unknown;
  __onlyOfficeBrowserModernThemeFilter?: boolean;
  __onlyOfficeBrowserOriginalMap?: () => Record<string, unknown>;
  __onlyOfficeBrowserOriginalGet?: (theme: string) => unknown;
  __onlyOfficeBrowserOriginalSetTheme?: (theme: OnlyOfficeUiTheme, source?: string) => unknown;
  __onlyOfficeBrowserOriginalDefaultThemeId?: () => string;
  __onlyOfficeBrowserOriginalDefaultTheme?: () => unknown;
};

type RuntimeWindow = typeof window & {
  AscDesktopEditor_PrintOptions?: { advancedOptions?: unknown } | undefined;
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
  __onlyOfficeBrowserFontInputFallbackUntil?: number;
  __onlyOfficeBrowserFontInputFallbackTimer?: number;
  AscBuilder?: {
    Word?: {
      Api?: {
        GetDocument?: () => {
          GetDefaultTextPr?: () => {
            SetFontFamily?: (fontFamily: string) => unknown;
            TextPr?: {
              RFonts?: {
                EastAsia?: { Name: string; Index: number };
                EastAsiaTheme?: string;
              };
            };
            private_OnChange?: () => unknown;
          } | null;
        } | null;
      };
    };
  };
  AscFonts?: {
    g_fontApplication?: {
      FontPickerMap?: Record<string, unknown>;
      GetFontFileWeb?: (name: string, style?: number) => unknown;
      __onlyOfficeBrowserFontFileFallback?: boolean;
      __onlyOfficeBrowserOriginalGetFontFileWeb?: (name: string, style?: number) => unknown;
      __onlyOfficeBrowserFallbackFontNames?: ReadonlySet<string>;
      __onlyOfficeBrowserFallbackFontFamilyName?: string;
    };
    CFont?: {
      prototype?: {
        name?: string;
        asc_getFontName?: () => string;
        __onlyOfficeBrowserFontNameFallback?: boolean;
        __onlyOfficeBrowserFallbackFontNames?: ReadonlySet<string>;
        __onlyOfficeBrowserFallbackFontFamilyName?: string;
      };
    };
  };
  AscCommon?: {
    asc_CTextFontFamily?: {
      prototype?: {
        Name?: string;
        get_Name?: () => string;
        asc_getName?: () => string;
        __onlyOfficeBrowserFontNameFallback?: boolean;
        __onlyOfficeBrowserOriginalTextFontName?: () => string;
        __onlyOfficeBrowserFallbackFontNames?: ReadonlySet<string>;
        __onlyOfficeBrowserFallbackFontFamilyName?: string;
      };
    };
    CKeyboardEvent?: new () => {
      CtrlKey?: boolean;
      KeyCode?: number;
    };
    baseEditorsApi?: {
      prototype?: {
        sync_InitEditorFonts?: (guiFonts: unknown[]) => unknown;
        __onlyOfficeBrowserFontPickerFilter?: boolean;
        __onlyOfficeBrowserVisibleFontNames?: ReadonlySet<string>;
      };
    };
  };
  Common?: {
    NotificationCenter?: OnlyOfficeNotificationCenter;
    UI?: {
      Themes?: OnlyOfficeThemeController;
      warning?: (options: {
        closable?: boolean;
        width?: number;
        title?: string;
        msg?: string;
        buttons?: string[];
        dontshow?: boolean;
        callback?: (result: string, dontShow?: boolean) => void;
      }) => unknown;
      Window?: {
        prototype?: {
          textWarning?: string;
        };
      };
    };
  };
  SSE?: {
    getController?: (name: string) => unknown;
  };
  Asc?: {
    editor?: {
      sync_TextPrFontFamilyCallBack?: (fontFamily?: {
        Name?: string;
        get_Name?: () => string;
        put_Name?: (name: string) => unknown;
      }) => unknown;
      put_TextPrFontName?: (fontFamily: string) => unknown;
      UpdateInterfaceState?: () => unknown;
      __onlyOfficeBrowserFontFamilyCallbackPatched?: boolean;
      __onlyOfficeBrowserOriginalFontFamilyCallback?: (fontFamily?: {
        Name?: string;
        get_Name?: () => string;
        put_Name?: (name: string) => unknown;
      }) => unknown;
      __onlyOfficeBrowserFallbackFontNames?: ReadonlySet<string>;
      __onlyOfficeBrowserFallbackFontFamilyName?: string;
    };
    c_oAscPrintType?: {
      Selection?: unknown;
      ActiveSheets?: unknown;
      EntireWorkbook?: unknown;
    };
    asc_CDownloadOptions?: new (fileType?: unknown, asUrl?: unknown) => OnlyOfficeDownloadOptionsLike;
  };
  DE?: {
    Controllers?: {
      LeftMenu?: {
        prototype?: {
          notcriticalErrorTitle?: string;
          txtCompatible?: string;
        };
        notcriticalErrorTitle?: string;
        txtCompatible?: string;
      };
    };
  };
  __fonts_visible_names?: unknown;
  __onlyOfficeBrowserPrintOpenGuard?: boolean;
  __onlyOfficeBrowserSpreadsheetPdfPrintExportMode?: SpreadsheetPdfPrintExportMode | null;
  __onlyOfficeBrowserOpeningSpreadsheetPdfPrintPanel?: boolean;
  __onlyOfficeBrowserShowingSpreadsheetPdfPrintPanel?: boolean;
  __onlyOfficeBrowserSpreadsheetPdfPrintPanel?: SpreadsheetPdfPrintPanelState | null;
};

let editorApiPromise: Promise<void> | null = null;
let nextEditorId = 1;
const activeInstances = new Map<string, BrowserOfficeEditor>();
const mediaRegistry = new Map<string, Record<string, string>>();
let temporaryDocumentTitleOriginal: string | null = null;
let temporaryDocumentTitleTimeout: number | null = null;
const MODERN_ONLYOFFICE_UI_THEME_IDS: OnlyOfficeUiTheme[] = ['theme-system', 'theme-white', 'theme-night'];

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

function normalizeOfficeInterfaceTheme(
  value: unknown,
  fallback: OnlyOfficeUiTheme = 'theme-system',
): OnlyOfficeUiTheme {
  if (typeof value !== 'string') return fallback;
  const normalized = value.trim().toLowerCase();
  if (normalized === 'system' || normalized === 'theme-system') return 'theme-system';
  if (normalized === 'light' || normalized === 'theme-white') return 'theme-white';
  if (normalized === 'dark' || normalized === 'theme-night') return 'theme-night';
  if (normalized === 'theme-dark' || normalized === 'theme-contrast-dark') return 'theme-night';
  if (normalized === 'theme-classic-light' || normalized === 'theme-light' || normalized === 'theme-gray') {
    return 'theme-white';
  }
  return fallback;
}

function persistOnlyOfficeInterfaceTheme(theme: OnlyOfficeUiTheme): void {
  try {
    window.localStorage.setItem('ui-theme-id', theme);
    const storedTheme = window.localStorage.getItem('ui-theme');
    if (!storedTheme) return;
    const storedThemeId = /"id":\s*"([\w-]+)"/.exec(storedTheme)?.[1];
    if (!storedThemeId || normalizeOfficeInterfaceTheme(storedThemeId) !== theme) {
      window.localStorage.removeItem('ui-theme');
    }
  } catch {
    // Storage can be disabled in private contexts; customization.uiTheme remains authoritative.
  }
}

function filterModernOnlyOfficeThemeMap(themeMap: Record<string, unknown> | undefined): Record<string, unknown> {
  const source = themeMap || {};
  return Object.fromEntries(
    MODERN_ONLYOFFICE_UI_THEME_IDS.filter((themeId) => Object.prototype.hasOwnProperty.call(source, themeId)).map(
      (themeId) => [themeId, source[themeId]],
    ),
  );
}

function installModernOnlyOfficeThemeFilter(frameWindow: OnlyOfficeFrameWindow | null): void {
  const themes = frameWindow?.Common?.UI?.Themes;
  if (!themes || themes.__onlyOfficeBrowserModernThemeFilter) return;

  const originalMap = typeof themes.map === 'function' ? themes.map.bind(themes) : undefined;
  const originalGet = typeof themes.get === 'function' ? themes.get.bind(themes) : undefined;
  const originalSetTheme = typeof themes.setTheme === 'function' ? themes.setTheme.bind(themes) : undefined;
  const originalDefaultThemeId =
    typeof themes.defaultThemeId === 'function' ? themes.defaultThemeId.bind(themes) : undefined;
  const originalDefaultTheme = typeof themes.defaultTheme === 'function' ? themes.defaultTheme.bind(themes) : undefined;

  themes.__onlyOfficeBrowserModernThemeFilter = true;
  themes.__onlyOfficeBrowserOriginalMap = originalMap;
  themes.__onlyOfficeBrowserOriginalGet = originalGet;
  themes.__onlyOfficeBrowserOriginalSetTheme = originalSetTheme;
  themes.__onlyOfficeBrowserOriginalDefaultThemeId = originalDefaultThemeId;
  themes.__onlyOfficeBrowserOriginalDefaultTheme = originalDefaultTheme;

  if (originalMap) {
    themes.map = () => filterModernOnlyOfficeThemeMap(originalMap());
  }
  if (originalGet) {
    themes.get = (theme: string) => originalGet(normalizeOfficeInterfaceTheme(theme));
  }
  if (originalSetTheme) {
    themes.setTheme = (theme: OnlyOfficeUiTheme, source?: string) =>
      originalSetTheme(normalizeOfficeInterfaceTheme(theme), source);
  }
  themes.defaultThemeId = () => 'theme-white';
  if (originalGet) {
    themes.defaultTheme = () => originalGet('theme-white');
  } else if (originalDefaultTheme) {
    themes.defaultTheme = originalDefaultTheme;
  }
}

function shouldFitEditorModePreviewToWidth(fileType: string, previewMode: boolean): boolean {
  const documentType = getDocumentType(fileType);
  if (documentType === 'word') return true;
  return !previewMode && documentType === 'slide';
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

function getSaveToNewFormatDialogTitle(frameWindow: OnlyOfficeFrameWindow, lang: string): string {
  const leftMenu = frameWindow.DE?.Controllers?.LeftMenu;
  return (
    leftMenu?.prototype?.notcriticalErrorTitle ||
    leftMenu?.notcriticalErrorTitle ||
    frameWindow.Common?.UI?.Window?.prototype?.textWarning ||
    (lang.toLowerCase().startsWith('zh') ? '警告' : 'Warning')
  );
}

function getSaveToNewFormatDialogMessage(frameWindow: OnlyOfficeFrameWindow, lang: string): string {
  const leftMenu = frameWindow.DE?.Controllers?.LeftMenu;
  const officialMessage = leftMenu?.prototype?.txtCompatible || leftMenu?.txtCompatible;
  if (officialMessage) return officialMessage;

  if (lang.toLowerCase().startsWith('zh')) {
    return '文档将保存为新格式。它将允许使用所有编辑器功能，但可能会影响文档布局<br>如果要使文件与旧的 MS Office 版本兼容，请使用高级设置的“兼容性”选项。';
  }
  return 'The document will be saved to the new format. It will allow you to use all editor features, but might affect the document layout.<br>Use the compatibility option in advanced settings if you want to keep files compatible with older MS Office versions.';
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

function normalizeDownloadFileType(value: unknown): string | number | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) return value;
  return undefined;
}

function getNativeDownloadAsFileType(
  options: NativeDownloadAsOptions | string | number | undefined,
): string | number | undefined {
  const direct = normalizeDownloadFileType(options);
  if (direct !== undefined) return direct;
  if (!options || typeof options !== 'object') return undefined;

  for (const key of ['fileType', 'format', 'outputformat', 'outputFormat', 'type'] as const) {
    const value = normalizeDownloadFileType(options[key]);
    if (value !== undefined) return value;
  }

  for (const key of [
    'asc_getFileType',
    'getFileType',
    'get_FileType',
    'asc_getFormat',
    'getFormat',
    'get_Format',
  ] as const) {
    const getter = options[key];
    if (typeof getter !== 'function') continue;
    try {
      const value = normalizeDownloadFileType(getter.call(options));
      if (value !== undefined) return value;
    } catch {
      // OnlyOffice option objects are not part of the public wrapper contract; ignore broken accessors.
    }
  }

  return undefined;
}

function isTruthyOnlyOfficeFlag(value: unknown): boolean {
  return value === true || value === 1 || value === '1' || value === 'true';
}

function getOnlyOfficeActionTypeNumber(value: unknown): number | undefined {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) {
    const numeric = Number(value);
    if (Number.isFinite(numeric)) return numeric;
  }
  return undefined;
}

function isOnlyOfficeNativePrintRequest({
  actionType,
  options,
  additionalData,
  downloadType,
}: {
  actionType?: unknown;
  options?: NativeDownloadAsOptions;
  additionalData?: NativeDownloadAsAdditionalData;
  downloadType?: unknown;
}): boolean {
  if (downloadType === ONLYOFFICE_PRINT_DOWNLOAD_TYPE) return true;
  if (String(downloadType || '') === ONLYOFFICE_PRINT_DOWNLOAD_TYPE) return true;
  if (getOnlyOfficeActionTypeNumber(actionType) === ONLYOFFICE_PRINT_ACTION_TYPE) return true;
  if (isTruthyOnlyOfficeFlag(options?.isPrint) || isTruthyOnlyOfficeFlag(options?.isPdfPrint)) return true;
  return isTruthyOnlyOfficeFlag(additionalData?.inline);
}

function isOnlyOfficeNativeSaveAsRequest({
  options,
  additionalData,
}: {
  options?: NativeDownloadAsOptions;
  additionalData?: NativeDownloadAsAdditionalData;
}): boolean {
  if (isTruthyOnlyOfficeFlag(options?.isSaveAs) || isTruthyOnlyOfficeFlag(additionalData?.isSaveAs)) return true;
  if (!options || typeof options !== 'object') return false;
  for (const key of ['asc_getIsSaveAs', 'getIsSaveAs', 'get_IsSaveAs'] as const) {
    const getter = options[key];
    if (typeof getter !== 'function') continue;
    try {
      if (isTruthyOnlyOfficeFlag(getter.call(options))) return true;
    } catch {
      // Ignore non-standard OnlyOffice option accessors.
    }
  }
  return false;
}

function getSpreadsheetPdfExportFileType(value: unknown): number | null {
  const normalized = normalizeDownloadFileType(value);
  if (normalized === undefined) return null;
  const targetExt = getDownloadTargetExtension(normalized, undefined, '');
  if (targetExt === 'pdf') return oAscFileType.PDF;
  if (targetExt === 'pdfa') return oAscFileType.PDFA;
  if (typeof normalized === 'number' && SPREADSHEET_PDF_EXPORT_FILE_TYPES.has(normalized)) return normalized;
  return null;
}

function getOnlyOfficeController<T>(frameWindow: OnlyOfficeFrameWindow, name: string): T | null {
  const controller = frameWindow.SSE?.getController?.(name);
  return controller && typeof controller === 'object' ? (controller as T) : null;
}

function getOnlyOfficeNumber(value: unknown, fallback = 0): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

function getOnlyOfficePageSetup(pageOptions: unknown): {
  asc_getWidth?: () => unknown;
  asc_getHeight?: () => unknown;
  asc_getOrientation?: () => unknown;
} | null {
  if (!pageOptions || typeof pageOptions !== 'object') return null;
  const getter = (pageOptions as { asc_getPageSetup?: () => unknown }).asc_getPageSetup;
  if (typeof getter !== 'function') return null;
  const pageSetup = getter.call(pageOptions);
  return pageSetup && typeof pageSetup === 'object' ? pageSetup : null;
}

function getOnlyOfficeButtonElement(value: unknown): HTMLElement | null {
  if (!value || typeof value !== 'object') return null;
  if (isHTMLElementLike(value)) return value as HTMLElement;
  const record = value as Record<string, unknown>;
  for (const key of ['el', '$el', 'btnEl', 'cmpEl']) {
    const candidate = record[key];
    if (isHTMLElementLike(candidate)) return candidate as HTMLElement;
    if (candidate && typeof candidate === 'object') {
      const indexed = (candidate as { 0?: unknown })[0];
      if (isHTMLElementLike(indexed)) return indexed as HTMLElement;
      const element = (candidate as { get?: (index: number) => unknown }).get?.(0);
      if (isHTMLElementLike(element)) return element as HTMLElement;
    }
  }
  return null;
}

function isHTMLElementLike(value: unknown): boolean {
  return Boolean(
    value &&
    typeof value === 'object' &&
    (value as { nodeType?: unknown }).nodeType === 1 &&
    typeof (value as { textContent?: unknown }).textContent === 'string',
  );
}

function setOnlyOfficeButtonCaption(button: unknown, caption: string): void {
  if (!button || typeof button !== 'object') return;
  const control = button as OnlyOfficeButtonLike;
  if (typeof control.setCaption === 'function') {
    control.setCaption(caption);
  } else if (typeof control.updateCaption === 'function') {
    control.updateCaption(caption);
  }
  const element = getOnlyOfficeButtonElement(button);
  if (!element) return;
  const labelElement = element.querySelector<HTMLElement>('.caption, .btn-text, span') || element;
  labelElement.textContent = caption;
  element.setAttribute('title', caption);
  element.setAttribute('aria-label', caption);
}

function getSpreadsheetPdfPrintActionLabel(lang: string, isSaveAs: boolean): string {
  if (lang.toLowerCase().startsWith('zh')) return isSaveAs ? '保存副本' : '下载';
  return isSaveAs ? 'Save Copy' : 'Download';
}

function getSpreadsheetPdfPrintPanelTitleLabel(lang: string, isSaveAs: boolean): string {
  if (lang.toLowerCase().startsWith('zh')) return isSaveAs ? '另存副本' : '下载';
  return isSaveAs ? 'Save Copy' : 'Download';
}

function getSpreadsheetPdfNativePrintPanelTitleLabel(
  lang: string,
  printSettings?: OnlyOfficePrintSettingsLike,
): string {
  if (typeof printSettings?.txtPrint === 'string' && printSettings.txtPrint.trim()) return printSettings.txtPrint;
  return lang.toLowerCase().startsWith('zh') ? '打印' : 'Print';
}

function getPrintSettingsSaveCaption(lang: string): string {
  return lang.toLowerCase().startsWith('zh') ? '保存设置' : 'Save settings';
}

function getSpreadsheetPdfPrintPanelElement(frameWindow: OnlyOfficeFrameWindow): HTMLElement | null {
  return frameWindow.document?.getElementById('panel-print');
}

function getSpreadsheetPdfPrintPanelHeaderElement(frameWindow: OnlyOfficeFrameWindow): HTMLElement | null {
  return frameWindow.document?.querySelector<HTMLElement>('#panel-print #id-print-settings .main-header') || null;
}

function setSpreadsheetPdfPrintPanelTitle(frameWindow: OnlyOfficeFrameWindow, title: string): void {
  const headerElement = getSpreadsheetPdfPrintPanelHeaderElement(frameWindow);
  if (headerElement) {
    headerElement.textContent = title;
  }
}

function getSpreadsheetPdfSourcePanelElement(
  frameWindow: OnlyOfficeFrameWindow,
  mode: SpreadsheetPdfPrintExportMode,
): HTMLElement | null {
  const panelId = mode.returnPanel === 'save-copy' ? 'panel-savecopy' : 'panel-saveas';
  return frameWindow.document?.getElementById(panelId);
}

function isSpreadsheetPdfPrintPanelReady(panel: HTMLElement | null): panel is HTMLElement {
  return Boolean(panel?.querySelector('#id-print-settings') && panel?.querySelector('#print-preview-box'));
}

function getDownloadFileName(sourceFileName: string, targetExt: string): string {
  const outputExtension = targetExt === 'pdfa' ? 'pdf' : targetExt === 'jpeg' ? 'jpg' : targetExt;
  return replaceFileExtension(sourceFileName, outputExtension);
}

function withFileName(file: File, fileName: string | undefined): File {
  const normalizedName = fileName?.trim();
  if (!normalizedName || normalizedName === file.name) return file;
  return new File([file], normalizedName, {
    type: file.type || getSavedFileMimeType(normalizedName),
    lastModified: file.lastModified,
  });
}

function getDefaultFileName(emptyType: OfficeEmptyType): string {
  return `New_Document.${emptyType}`;
}

function getSavedFileMimeType(fileName: string): string {
  const extension = getFileExtension(fileName);
  const mimeMap: Record<string, string> = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    doc: 'application/msword',
    docm: 'application/vnd.ms-word.document.macroEnabled.12',
    docxf: 'application/vnd.onlyoffice.docxf',
    dotx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
    epub: 'application/epub+zip',
    fb2: 'application/x-fictionbook+xml',
    html: 'text/html',
    md: 'text/markdown',
    oform: 'application/vnd.onlyoffice.oform',
    odt: 'application/vnd.oasis.opendocument.text',
    ott: 'application/vnd.oasis.opendocument.text-template',
    rtf: 'application/rtf',
    txt: 'text/plain',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    xls: 'application/vnd.ms-excel',
    xlsb: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12',
    xltm: 'application/vnd.ms-excel.template.macroEnabled.12',
    xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    ods: 'application/vnd.oasis.opendocument.spreadsheet',
    ots: 'application/vnd.oasis.opendocument.spreadsheet-template',
    csv: 'text/csv',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    ppt: 'application/vnd.ms-powerpoint',
    pptm: 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
    ppsm: 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
    ppsx: 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
    potm: 'application/vnd.ms-powerpoint.template.macroEnabled.12',
    potx: 'application/vnd.openxmlformats-officedocument.presentationml.template',
    odp: 'application/vnd.oasis.opendocument.presentation',
    otp: 'application/vnd.oasis.opendocument.presentation-template',
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
  const directUrl = name.trim();
  if (isOnlyOfficeDirectMediaUrl(directUrl)) {
    return directUrl;
  }

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
  const url = new URL(
    `${PRINT_PDF_ROUTE_PREFIX}${randomId}/${encodeURIComponent(pdfFileName)}`,
    window.location.origin,
  );
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

function assertNonEmptyPdfDocument(bytes: Uint8Array): void {
  if (!looksLikePdfDocument(bytes)) return;

  const source = bytesToBinaryString(bytes);
  const hasPageObject = /\/Type\s*\/Page(?!s)\b/.test(source);
  const hasZeroPageTree =
    /\/Type\s*\/Pages\b[\s\S]{0,300}\/Count\s+0\b/.test(source) ||
    /\/Count\s+0\b[\s\S]{0,300}\/Type\s*\/Pages\b/.test(source);

  if (hasZeroPageTree && !hasPageObject) {
    throw new Error('PDF export produced an empty 0-page document');
  }
}

function getPdfPageCount(bytes: Uint8Array): number | null {
  if (!looksLikePdfDocument(bytes)) return null;

  const source = bytesToBinaryString(bytes);
  const counts = [
    ...Array.from(source.matchAll(/\/Type\s*\/Pages\b[\s\S]{0,300}?\/Count\s+(\d+)\b/g)),
    ...Array.from(source.matchAll(/\/Count\s+(\d+)\b[\s\S]{0,300}?\/Type\s*\/Pages\b/g)),
  ]
    .map((match) => Number(match[1]))
    .filter((value) => Number.isFinite(value) && value >= 0);
  if (counts.length === 0) return null;
  return Math.max(...counts);
}

function hasBinaryNull(bytes: Uint8Array): boolean {
  return bytes.includes(0);
}

function looksLikeJpegDocument(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 4 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[bytes.length - 2] === 0xff &&
    bytes[bytes.length - 1] === 0xd9
  );
}

function decodeUtf8Loose(bytes: Uint8Array): string {
  return new TextDecoder('utf-8', { fatal: false })
    .decode(bytes)
    .replace(/^\ufeff/, '')
    .trimStart();
}

function looksLikePngDocument(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  );
}

function describeDownloadBytes(bytes: Uint8Array, fileName: string): string {
  const header = Array.from(bytes.slice(0, 16))
    .map((byte) => byte.toString(16).padStart(2, '0'))
    .join(' ');
  return `${fileName}; header: ${header || '<empty>'}`;
}

function assertTextLikeDownload(bytes: Uint8Array, targetExt: string): void {
  if (hasBinaryNull(bytes)) {
    throw new Error(`Download As ${targetExt.toUpperCase()} produced binary data`);
  }
  const text = decodeUtf8Loose(bytes);
  if (!text) {
    throw new Error(`Download As ${targetExt.toUpperCase()} produced an empty text document`);
  }
}

function assertValidDownloadFileBytes(bytes: Uint8Array, fileName: string, targetExt: string): void {
  const normalizedTargetExt = normalizeDownloadTargetExtension(targetExt || getFileExtension(fileName));
  const outputExt = getFileExtension(fileName) || normalizedTargetExt;

  if (bytes.byteLength === 0) {
    throw new Error(`Download As ${normalizedTargetExt.toUpperCase()} produced an empty file`);
  }

  if (normalizedTargetExt === 'pdf' || normalizedTargetExt === 'pdfa' || outputExt === 'pdf') {
    if (!looksLikePdfDocument(bytes)) {
      throw new Error(`Download As ${normalizedTargetExt.toUpperCase()} did not produce a PDF document`);
    }
    assertNonEmptyPdfDocument(bytes);
    return;
  }

  if (ZIP_DOWNLOAD_EXTENSIONS.has(outputExt)) {
    if (!looksLikeZipPackage(bytes)) {
      throw new Error(`Download As ${normalizedTargetExt.toUpperCase()} did not produce a ZIP-based package`);
    }
    return;
  }

  if (normalizedTargetExt === 'jpg' || normalizedTargetExt === 'jpeg' || outputExt === 'jpg' || outputExt === 'jpeg') {
    if (!looksLikeJpegDocument(bytes)) {
      throw new Error(
        `Download As ${normalizedTargetExt.toUpperCase()} did not produce a JPEG image (${describeDownloadBytes(
          bytes,
          fileName,
        )})`,
      );
    }
    return;
  }

  if (normalizedTargetExt === 'png' || outputExt === 'png') {
    if (!looksLikePngDocument(bytes)) {
      throw new Error(
        `Download As ${normalizedTargetExt.toUpperCase()} did not produce a PNG image (${describeDownloadBytes(
          bytes,
          fileName,
        )})`,
      );
    }
    return;
  }

  if (ZIP_DOWNLOAD_EXTENSIONS.has(normalizedTargetExt)) {
    if (!looksLikeZipPackage(bytes)) {
      throw new Error(`Download As ${normalizedTargetExt.toUpperCase()} did not produce a ZIP-based Office package`);
    }
    return;
  }

  if (normalizedTargetExt === 'rtf' || outputExt === 'rtf') {
    const text = decodeUtf8Loose(bytes);
    if (!text.startsWith('{\\rtf')) {
      throw new Error('Download As RTF did not produce an RTF document');
    }
    return;
  }

  if (normalizedTargetExt === 'html' || outputExt === 'html') {
    assertTextLikeDownload(bytes, normalizedTargetExt);
    const text = decodeUtf8Loose(bytes).toLowerCase();
    if (!/(<!doctype\s+html|<html\b|<body\b|<[a-z][\w:-]*\b)/i.test(text)) {
      throw new Error('Download As HTML did not produce an HTML document');
    }
    return;
  }

  if (normalizedTargetExt === 'fb2' || outputExt === 'fb2') {
    assertTextLikeDownload(bytes, normalizedTargetExt);
    if (!/fictionbook/i.test(decodeUtf8Loose(bytes))) {
      throw new Error('Download As FB2 did not produce a FictionBook document');
    }
    return;
  }

  if (['csv', 'md', 'txt'].includes(normalizedTargetExt) || ['csv', 'md', 'txt'].includes(outputExt)) {
    assertTextLikeDownload(bytes, normalizedTargetExt);
  }
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (char) => char.charCodeAt(0) & 0xff);
}

function utf8Bytes(value: string): Uint8Array {
  return new TextEncoder().encode(value);
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function crc32(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) {
    value = CRC32_TABLE[(value ^ byte) & 0xff] ^ (value >>> 8);
  }
  return (value ^ 0xffffffff) >>> 0;
}

function writeUint16LE(output: number[], value: number): void {
  output.push(value & 0xff, (value >>> 8) & 0xff);
}

function writeUint32LE(output: number[], value: number): void {
  output.push(value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff);
}

function createStoredZip(files: Array<{ path: string; data: Uint8Array }>): Uint8Array {
  const local: number[] = [];
  const central: number[] = [];
  const now = new Date();
  const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | Math.floor(now.getSeconds() / 2);
  const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();

  for (const file of files) {
    const name = utf8Bytes(file.path);
    const checksum = crc32(file.data);
    const localOffset = local.length;

    writeUint32LE(local, 0x04034b50);
    writeUint16LE(local, 20);
    writeUint16LE(local, 0x0800);
    writeUint16LE(local, 0);
    writeUint16LE(local, dosTime);
    writeUint16LE(local, dosDate);
    writeUint32LE(local, checksum);
    writeUint32LE(local, file.data.byteLength);
    writeUint32LE(local, file.data.byteLength);
    writeUint16LE(local, name.byteLength);
    writeUint16LE(local, 0);
    local.push(...name, ...file.data);

    writeUint32LE(central, 0x02014b50);
    writeUint16LE(central, 20);
    writeUint16LE(central, 20);
    writeUint16LE(central, 0x0800);
    writeUint16LE(central, 0);
    writeUint16LE(central, dosTime);
    writeUint16LE(central, dosDate);
    writeUint32LE(central, checksum);
    writeUint32LE(central, file.data.byteLength);
    writeUint32LE(central, file.data.byteLength);
    writeUint16LE(central, name.byteLength);
    writeUint16LE(central, 0);
    writeUint16LE(central, 0);
    writeUint16LE(central, 0);
    writeUint16LE(central, 0);
    writeUint32LE(central, 0);
    writeUint32LE(central, localOffset);
    central.push(...name);
  }

  const centralOffset = local.length;
  const centralSize = central.length;
  const end: number[] = [];
  writeUint32LE(end, 0x06054b50);
  writeUint16LE(end, 0);
  writeUint16LE(end, 0);
  writeUint16LE(end, files.length);
  writeUint16LE(end, files.length);
  writeUint32LE(end, centralSize);
  writeUint32LE(end, centralOffset);
  writeUint16LE(end, 0);

  return Uint8Array.from([...local, ...central, ...end]);
}

function normalizeMarkdownImageExtension(mimeSubtype: string): string {
  const normalized = mimeSubtype.toLowerCase().replace(/[^a-z0-9+.-]/g, '');
  if (normalized === 'jpeg' || normalized === 'pjpeg') return 'jpg';
  if (normalized === 'svg+xml') return 'svg';
  return normalized || 'bin';
}

function decodeDataUriBase64(value: string): Uint8Array | null {
  try {
    const binary = atob(value.replace(/\s+/g, ''));
    return Uint8Array.from(binary, (char) => char.charCodeAt(0));
  } catch {
    return null;
  }
}

function getExtensionFileNameSuffix(fileName: string): string {
  const extension = getFileExtension(fileName)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
  return extension ? `_${extension}` : '';
}

function packageMarkdownDataUriImages(
  fileName: string,
  bytes: Uint8Array,
  sourceFileName = fileName,
): { fileName: string; data: Uint8Array } | null {
  const markdown = decodeUtf8Loose(bytes);
  if (!markdown) return null;

  const baseName = (fileName.split(/[\\/]/).pop() || 'document').replace(/\.[^/.]+$/, '') || 'document';
  const files: Array<{ path: string; data: Uint8Array }> = [];
  let imageIndex = 0;

  const addImage = (mimeSubtype: string, base64: string): string | null => {
    const data = decodeDataUriBase64(base64);
    if (!data) return null;
    imageIndex += 1;
    const extension = normalizeMarkdownImageExtension(mimeSubtype);
    const imagePath = `assets/image-${String(imageIndex).padStart(3, '0')}.${extension}`;
    files.push({ path: imagePath, data });
    return imagePath;
  };

  let rewritten = markdown.replace(
    /(!\[[^\]]*]\()data:image\/([^;)\s]+);base64,([A-Za-z0-9+/=\s]+)(\))/gi,
    (match, prefix: string, mimeSubtype: string, base64: string, suffix: string) => {
      const imagePath = addImage(mimeSubtype, base64);
      return imagePath ? `${prefix}${imagePath}${suffix}` : match;
    },
  );

  rewritten = rewritten.replace(
    /(<img\b[^>]*\bsrc=["'])data:image\/([^;"']+);base64,([^"']+)(["'][^>]*>)/gi,
    (match, prefix: string, mimeSubtype: string, base64: string, suffix: string) => {
      const imagePath = addImage(mimeSubtype, base64);
      return imagePath ? `${prefix}${imagePath}${suffix}` : match;
    },
  );

  if (files.length === 0) return null;

  files.unshift({ path: `${baseName}.md`, data: utf8Bytes(rewritten) });
  return {
    fileName: `${baseName}${getExtensionFileNameSuffix(sourceFileName)}_md.zip`,
    data: createStoredZip(files),
  };
}

function hasNativeDownloadPrintOptions(options: NativeDownloadAsOptions | undefined): boolean {
  return Boolean(options && (options.advancedOptions !== undefined || options.printOptions !== undefined));
}

function getNativeDownloadAdvancedOptions(options: NativeDownloadAsOptions | undefined): unknown {
  return options?.advancedOptions ?? options?.printOptions;
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

function hideIframeForTeardown(iframe: HTMLIFrameElement): void {
  iframe.setAttribute('aria-hidden', 'true');
  iframe.style.visibility = 'hidden';
  iframe.style.opacity = '0';
  iframe.style.pointerEvents = 'none';
  iframe.style.background = 'transparent';
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

export function filterEditorFontsByVisibleNames(guiFonts: unknown[], visibleNames: ReadonlySet<string>): unknown[] {
  return guiFonts.filter((font) => {
    const name = getFontNameForFilter(font);
    return name ? visibleNames.has(name) : false;
  });
}

export function applyDefaultWordFont(
  frameWindow: OnlyOfficeFrameWindow,
  latinFontFamilyName = ONLYOFFICE_LATIN_DEFAULT_FONT_FAMILY,
  eastAsiaFontFamilyName = ONLYOFFICE_EAST_ASIA_DEFAULT_FONT_FAMILY,
): boolean {
  const defaultTextPr = frameWindow.AscBuilder?.Word?.Api?.GetDocument?.()?.GetDefaultTextPr?.();
  if (!defaultTextPr || typeof defaultTextPr.SetFontFamily !== 'function') return false;

  // ApiTextPr.SetFontFamily intentionally sets all four OOXML font slots.
  // Keep the public API for the Western default, then restore the EastAsia
  // slot on the underlying CTextPr. This mirrors OOXML rFonts semantics while
  // leaving ONLYOFFICE's native missing-font selection untouched.
  defaultTextPr.SetFontFamily(latinFontFamilyName);
  if (defaultTextPr.TextPr?.RFonts) {
    defaultTextPr.TextPr.RFonts.EastAsia = { Name: eastAsiaFontFamilyName, Index: -1 };
    defaultTextPr.TextPr.RFonts.EastAsiaTheme = undefined;
    defaultTextPr.private_OnChange?.();
  }
  frameWindow.Asc?.editor?.put_TextPrFontName?.(latinFontFamilyName);
  frameWindow.Asc?.editor?.UpdateInterfaceState?.();
  return true;
}

export function installFontPickerFilter(
  frameWindow: OnlyOfficeFrameWindow,
  configuredNames?: string[],
  fallbackNames?: string[],
  fallbackFamilyName = ONLYOFFICE_EAST_ASIA_DEFAULT_FONT_FAMILY,
): boolean {
  const visibleNames = new Set(configuredNames || []);
  // Keep unavailable names out of the picker, but do not rewrite document font
  // names or GetFontFileWeb. DocumentServer's generated font dictionary ranks
  // missing-font candidates by name, PANOSE, Unicode/code-page ranges, weight,
  // width and font metrics. Preserving that path gives old documents the same
  // substitution behavior as native ONLYOFFICE Docs.
  void fallbackNames;
  void fallbackFamilyName;

  const prototype = frameWindow.AscCommon?.baseEditorsApi?.prototype;
  const original = prototype?.sync_InitEditorFonts;
  if (!prototype || typeof original !== 'function') return false;
  prototype.__onlyOfficeBrowserVisibleFontNames = visibleNames;
  if (prototype.__onlyOfficeBrowserFontPickerFilter) return true;

  prototype.sync_InitEditorFonts = function syncInitEditorFontsWithFilter(this: unknown, guiFonts: unknown[]) {
    const currentVisibleNames = prototype.__onlyOfficeBrowserVisibleFontNames || new Set<string>();
    const filteredFonts = Array.isArray(guiFonts)
      ? filterEditorFontsByVisibleNames(guiFonts, currentVisibleNames)
      : guiFonts;
    return original.call(this, filteredFonts);
  };
  prototype.__onlyOfficeBrowserFontPickerFilter = true;
  return true;
}

export function installNestedFontPickerFilter(visibleFontNames?: string[], fallbackFontNames?: string[]): () => void {
  const startedAt = Date.now();
  const timeoutMs = 10_000;
  const intervalMs = 50;
  const runtimeDocument = document;
  const runtimeWindow = window;
  let cancelled = false;
  let retryId: number | null = null;

  const tryInstall = () => {
    retryId = null;
    if (cancelled) return;
    const frame = runtimeDocument.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameWindow = frame?.contentWindow as OnlyOfficeFrameWindow | null | undefined;
    if (frameWindow) {
      try {
        if (installFontPickerFilter(frameWindow, visibleFontNames, fallbackFontNames)) return;
      } catch {
        return;
      }
    }

    if (Date.now() - startedAt < timeoutMs) {
      retryId = runtimeWindow.setTimeout(tryInstall, intervalMs);
    }
  };

  tryInstall();
  return () => {
    cancelled = true;
    if (retryId !== null) {
      runtimeWindow.clearTimeout(retryId);
      retryId = null;
    }
  };
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
  private binData: BlobPart;
  private readonly originalFile?: File;
  private readonly sourceKind: OfficeEditorSourceKind;
  private readonly editorLang: string;
  private readonly readonlyMode: { value: boolean };
  private readonly previewEditAllowed: boolean;
  private editorMode: OfficeEditorMode;
  private destroyPromise: Promise<void> | null = null;
  private downloadAsInterceptorRetryId: number | null = null;
  private downloadAsInterceptorAttempts = 0;
  private switchingPreviewMode = false;
  private nativeEditModeBridgeCleanup: (() => void) | null = null;
  private nativeEditModeBridgeRetryId: number | null = null;
  private nestedFontPickerFilterCleanup: (() => void) | null = null;

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
    this.binData = prepared.binData;
    this.originalFile = prepared.originalFile;
    this.sourceKind = prepared.sourceKind;
    this.sourceObjectUrl = prepared.sourceObjectUrl || null;
    this.media = prepared.media || {};
    this.editorLang = options.lang || getOnlyOfficeLang();
    const initialMode = resolveInitialMode(options);
    this.editorMode = initialMode;
    this.previewEditAllowed =
      options.readonly !== true && (initialMode === 'preview' || options.canReturnToPreview === true);
    this.readonlyMode = { value: initialMode !== 'edit' };
  }

  private get previewMode(): boolean {
    return this.editorMode === 'preview';
  }

  static async create(container: HTMLElement, options: CreateOfficeEditorOptions): Promise<BrowserOfficeEditor> {
    const fontManifest = await assertGeneratedFontAssetsAvailable();
    await loadOfficeEditorApi();
    await initX2T();
    const resolvedOptions: CreateOfficeEditorOptions = {
      ...options,
      visibleFontNames: options.visibleFontNames || resolveAvailableFontFamilyNames(fontManifest, []),
      fallbackFontNames:
        options.fallbackFontNames ||
        (fontManifest.fontFamilies || [])
          .map((family) => family.name)
          .filter((name) => !resolveAvailableFontFamilyNames(fontManifest, []).includes(name)),
    };
    const prepared = await prepareDocument(resolvedOptions);
    const placeholder = document.createElement('div');
    const instance = new BrowserOfficeEditor(container, resolvedOptions, prepared, placeholder);
    await instance.mount();
    return instance;
  }

  private async mount(): Promise<void> {
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

    this.editorBinUrl = createObjectUrl(new Blob([this.binData], { type: 'application/octet-stream' }));
    registerMedia(this.id, this.media);
    activeInstances.set(this.id, this);

    try {
      this.scheduleNestedFontPickerFilter();
      const defaultZoom = getDefaultEditorModePreviewZoom(this.fileType, this.previewMode);
      persistDefaultEditorModePreviewZoom(this.fileType, this.previewMode);
      const uiTheme = normalizeOfficeInterfaceTheme(this.options.interfaceTheme);
      persistOnlyOfficeInterfaceTheme(uiTheme);
      const canEditInitially = !this.previewMode && !this.readonlyMode.value;
      const canRequestEditFromPreview = this.previewMode && this.previewEditAllowed;
      const editor = new runtimeWindow.DocsAPI.DocEditor(this.placeholder.id, {
        type: 'desktop',
        width: '100%',
        height: '100%',
        documentType: getDocumentType(this.fileType) || undefined,
        document: {
          title: this.fileName,
          url: this.editorBinUrl,
          fileType: this.fileType,
          key: `local-${this.id}-${Date.now()}`,
          permissions: {
            edit: canRequestEditFromPreview || canEditInitially,
            download: true,
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
          embedded: undefined,
          customization: {
            help: false,
            about: false,
            hideRightMenu: true,
            compactToolbar: true,
            uiTheme,
            zoom: defaultZoom,
            spellcheck: this.options.spellcheck ?? false,
            autosave: false,
            forcesave: false,
            plugins: Boolean(this.options.plugins?.configUrls.length),
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
          plugins: this.options.plugins?.configUrls.length
            ? {
                pluginsData: this.options.plugins.configUrls.map((url) => new URL(url, window.location.origin).href),
                autostart: this.options.plugins.autostart,
              }
            : undefined,
        },
        events: {
          onAppReady: () => {
            if (this.destroyed) return;
            this.applyNestedEditorFrameDefaults();
            this.installModernThemeFilter();
            this.scheduleNestedFontPickerFilter();
            this.installSpreadsheetPdfPrintPanelBridge();
            this.installNativeEditModeBridge();
            this.scheduleNativeDownloadAsInterceptor();
          },
          onDocumentReady: () => {
            if (this.destroyed) return;
            this.applyNestedEditorFrameDefaults();
            this.installModernThemeFilter();
            this.scheduleNestedFontPickerFilter();
            if (this.sourceKind === 'new-document' && getDocumentType(this.fileType) === 'word') {
              const frameWindow = this.getNestedEditorWindow() || this.getFrameEditorWindow();
              if (frameWindow) applyDefaultWordFont(frameWindow);
            }
            this.installSpreadsheetPdfPrintPanelBridge();
            this.installNativeEditModeBridge();
            this.scheduleNativeDownloadAsInterceptor();
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
          onRequestSaveAs: (event: RequestSaveAsEvent) => {
            void this.handleRequestSaveAs(event);
          },
          onRequestEditRights: () => {
            void this.handleRequestEditRights();
          },
          writeFile: (event: any) => {
            void this.handleWriteFile(event);
          },
        },
      });

      this.editor = editor;
      this.applyNestedEditorFrameDefaults();
      this.installModernThemeFilter();
      this.installNativeEditModeBridge();
      this.attachDestroyCleanup(editor);
      this.scheduleNativeDownloadAsInterceptor();

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

  setInterfaceTheme(theme: OfficeInterfaceTheme): void {
    const uiTheme = normalizeOfficeInterfaceTheme(theme);
    persistOnlyOfficeInterfaceTheme(uiTheme);

    const frameWindow = this.getNestedEditorWindow();
    installModernOnlyOfficeThemeFilter(frameWindow);
    const themes = frameWindow?.Common?.UI?.Themes;
    if (typeof themes?.setTheme === 'function') {
      themes.setTheme(uiTheme, 'host');
    }
  }

  private installModernThemeFilter(): void {
    const frameWindow = this.getNestedEditorWindow();
    installModernOnlyOfficeThemeFilter(frameWindow);
    if (frameWindow?.Common?.UI?.Themes) {
      this.setInterfaceTheme(this.options.interfaceTheme || 'system');
    }
  }

  private clearNativeEditModeBridgeRetry(): void {
    if (this.nativeEditModeBridgeRetryId === null) return;
    window.clearTimeout(this.nativeEditModeBridgeRetryId);
    this.nativeEditModeBridgeRetryId = null;
  }

  private scheduleNestedFontPickerFilter(): void {
    this.nestedFontPickerFilterCleanup?.();
    this.nestedFontPickerFilterCleanup = installNestedFontPickerFilter(
      this.options.visibleFontNames,
      this.options.fallbackFontNames,
    );
  }

  private scheduleNativeEditModeBridgeRetry(): void {
    if (this.nativeEditModeBridgeRetryId !== null || this.destroyed) return;
    this.nativeEditModeBridgeRetryId = window.setTimeout(() => {
      this.nativeEditModeBridgeRetryId = null;
      this.installNativeEditModeBridge();
    }, 200);
  }

  private installNativeEditModeBridge(): void {
    if (!this.previewEditAllowed) {
      this.clearNativeEditModeBridgeRetry();
      return;
    }

    const frameWindow = this.getFrameEditorWindow();
    const frameDocument = frameWindow?.document;
    if (!frameWindow || !frameDocument) {
      this.scheduleNativeEditModeBridgeRetry();
      return;
    }

    this.clearNativeEditModeBridgeRetry();
    this.nativeEditModeBridgeCleanup?.();
    this.nativeEditModeBridgeCleanup = null;

    let checkTimeoutId: number | null = null;
    const getNativeModeCaption = () =>
      frameDocument
        .querySelector('#slot-btn-edit-mode .btn.dropdown-toggle .caption')
        ?.textContent?.replace(/\s+/g, ' ')
        .trim()
        .toLowerCase() || '';
    const isVisibleElement = (element: Element | null) => {
      const FrameHTMLElement =
        (frameWindow as Window & { HTMLElement?: typeof HTMLElement }).HTMLElement || HTMLElement;
      if (!element || !(element instanceof FrameHTMLElement)) return false;
      const style = frameWindow.getComputedStyle(element);
      if (style.display === 'none' || style.visibility === 'hidden') return false;
      return element.getClientRects().length > 0 || /\bjsdom\b/i.test(frameWindow.navigator.userAgent);
    };
    const hasNativeEditToolbar = () => isVisibleElement(frameDocument.getElementById('home'));
    const returnPreviewTitle = this.editorLang.toLowerCase().startsWith('zh') ? '返回预览' : 'Return to preview';
    const ensureNativeReturnPreviewStyle = () => {
      if (frameDocument.getElementById('onlyoffice-browser-return-preview-mode-style')) return;
      const style = frameDocument.createElement('style');
      style.id = 'onlyoffice-browser-return-preview-mode-style';
      style.textContent = `
        #onlyoffice-browser-return-preview-mode {
          display: inline-flex;
          align-items: center;
          vertical-align: middle;
          margin: 0 4px 0 2px;
        }
        #onlyoffice-browser-return-preview-mode .onlyoffice-browser-return-preview-mode-button {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          min-width: 32px;
          width: 32px;
          height: 32px;
          padding: 0;
          border: 0;
          border-radius: 4px;
          background: transparent;
          color: var(--text-toolbar-header, var(--text-normal, #f2f2f2));
          font: inherit;
          line-height: 1;
          cursor: pointer;
        }
        #onlyoffice-browser-return-preview-mode .onlyoffice-browser-return-preview-mode-button:hover,
        #onlyoffice-browser-return-preview-mode .onlyoffice-browser-return-preview-mode-button:focus-visible {
          background: var(--highlight-header-button-hover, rgba(255, 255, 255, 0.12));
          outline: none;
        }
        #onlyoffice-browser-return-preview-mode .onlyoffice-browser-return-preview-mode-icon {
          display: block;
          width: 16px;
          height: 16px;
          flex: 0 0 auto;
        }
      `;
      frameDocument.head.appendChild(style);
    };
    const getButtonSlotAnchor = (element: Element) =>
      element.closest('.btn-slot, .btn-group, .toolbar-item, .toolbar-group') || element;
    const getNativeUserAvatarAnchor = () => {
      for (const id of RETURN_PREVIEW_AVATAR_ANCHOR_IDS) {
        const element = frameDocument.getElementById(id);
        if (element && isVisibleElement(element)) return getButtonSlotAnchor(element);
      }

      const isJSDOM = /\bjsdom\b/i.test(frameWindow.navigator.userAgent);
      const candidates = frameDocument.querySelectorAll(
        'button, [role="button"], .btn, .btn-slot, [id*="user"], [class*="user"], [id*="avatar"], [class*="avatar"], [id*="profile"], [class*="profile"], [id*="coauth"], [class*="coauth"]',
      );
      for (const element of Array.from(candidates)) {
        if (!isVisibleElement(element)) continue;
        const text = element.textContent?.replace(/\s+/g, '').trim() || '';
        if (!/^[A-Z]{1,3}$/.test(text)) continue;
        if (!isJSDOM) {
          const rect = element.getBoundingClientRect();
          if (rect.top > 64 || rect.width > 96 || rect.height > 48) continue;
        }
        return getButtonSlotAnchor(element);
      }

      return null;
    };
    const getNativeReturnPreviewAnchor = () =>
      getNativeUserAvatarAnchor() ||
      frameDocument.getElementById('slot-btn-search') ||
      frameDocument.getElementById('slot-btn-print') ||
      frameDocument.getElementById('title-doc-name') ||
      frameDocument.getElementById('toolbar');
    const removeNativeReturnPreviewButton = () => {
      frameDocument.getElementById('onlyoffice-browser-return-preview-mode')?.remove();
    };
    const syncNativeReturnPreviewButton = () => {
      if (this.destroyed || this.previewMode || this.readonlyMode.value) {
        removeNativeReturnPreviewButton();
        return;
      }

      const anchor = getNativeReturnPreviewAnchor();
      const host = anchor?.parentElement;
      if (!anchor || !host) return;

      ensureNativeReturnPreviewStyle();
      let buttonSlot = frameDocument.getElementById('onlyoffice-browser-return-preview-mode');
      if (!buttonSlot) {
        buttonSlot = frameDocument.createElement('span');
        buttonSlot.id = 'onlyoffice-browser-return-preview-mode';
        buttonSlot.className = 'btn-slot onlyoffice-browser-return-preview-mode';
        const button = frameDocument.createElement('button');
        button.type = 'button';
        button.className = 'btn btn-header onlyoffice-browser-return-preview-mode-button';
        button.title = returnPreviewTitle;
        button.setAttribute('aria-label', returnPreviewTitle);
        const icon = frameDocument.createElementNS('http://www.w3.org/2000/svg', 'svg');
        icon.classList.add('onlyoffice-browser-return-preview-mode-icon');
        icon.setAttribute('viewBox', '0 0 16 16');
        icon.setAttribute('width', '16');
        icon.setAttribute('height', '16');
        icon.setAttribute('aria-hidden', 'true');
        icon.setAttribute('focusable', 'false');
        icon.setAttribute('data-iconify-icon', RETURN_PREVIEW_ICON_NAME);
        icon.innerHTML = RETURN_PREVIEW_ICON_BODY;
        button.appendChild(icon);
        button.addEventListener('click', (event) => {
          event.preventDefault();
          event.stopPropagation();
          void this.handleReturnToPreview();
        });
        buttonSlot.appendChild(button);
      }

      if (buttonSlot.parentElement !== host || buttonSlot.nextSibling !== anchor) {
        host.insertBefore(buttonSlot, anchor);
      }
    };
    const isNativeEditMode = () => {
      const caption = getNativeModeCaption();
      return (caption.startsWith('edit') || caption.startsWith('编辑')) && hasNativeEditToolbar();
    };
    const isNativePreviewMode = () => {
      const caption = getNativeModeCaption();
      return (
        caption.startsWith('view') ||
        caption.startsWith('查看') ||
        caption.startsWith('preview') ||
        caption.startsWith('预览')
      );
    };
    const syncNativeEditMode = () => {
      checkTimeoutId = null;
      if (this.destroyed || !this.previewEditAllowed || this.switchingPreviewMode) return;
      if (this.previewMode) {
        if (!isNativeEditMode()) return;
        this.editorMode = 'edit';
        this.readonlyMode.value = false;
        this.notifyStateChange();
        syncNativeReturnPreviewButton();
        return;
      }
      syncNativeReturnPreviewButton();
      if (!this.readonlyMode.value && isNativePreviewMode()) {
        void this.handleReturnToPreview();
      }
    };
    const checkIntervalId = frameWindow.setInterval(() => syncNativeEditMode(), 250);
    const scheduleSync = () => {
      if (checkTimeoutId !== null) {
        frameWindow.clearTimeout(checkTimeoutId);
      }
      checkTimeoutId = frameWindow.setTimeout(syncNativeEditMode, 0);
    };

    const editModeSlot = frameDocument.getElementById('slot-btn-edit-mode');
    const observer = typeof MutationObserver === 'function' ? new MutationObserver(scheduleSync) : null;
    observer?.observe(editModeSlot || frameDocument.body, {
      childList: true,
      subtree: true,
      characterData: true,
      attributes: true,
      attributeFilter: ['class', 'style', 'title', 'aria-label'],
    });
    frameDocument.addEventListener('click', scheduleSync, true);
    frameDocument.addEventListener('keyup', scheduleSync, true);
    scheduleSync();

    this.nativeEditModeBridgeCleanup = () => {
      removeNativeReturnPreviewButton();
      observer?.disconnect();
      frameDocument.removeEventListener('click', scheduleSync, true);
      frameDocument.removeEventListener('keyup', scheduleSync, true);
      if (checkTimeoutId !== null) {
        frameWindow.clearTimeout(checkTimeoutId);
        checkTimeoutId = null;
      }
      frameWindow.clearInterval(checkIntervalId);
      this.clearNativeEditModeBridgeRetry();
    };
  }

  private getNestedEditorWindow(): OnlyOfficeFrameWindow | null {
    const editorWindow = this.editor?.getEditorWindow?.() as OnlyOfficeFrameWindow | null | undefined;
    if (editorWindow) return editorWindow;

    return this.getFrameEditorWindow();
  }

  private getFrameEditorWindow(): OnlyOfficeFrameWindow | null {
    const frame =
      this.placeholder.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]') ||
      this.container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameWindow = frame?.contentWindow as OnlyOfficeFrameWindow | null | undefined;
    return frameWindow || null;
  }

  private getNativeEditorApi(): OnlyOfficeEditorApi | null {
    const editorWithApi = this.editor as (DocEditor & { getEditorApi?: () => unknown }) | null;
    const directApi = editorWithApi?.getEditorApi?.();
    if (directApi && typeof directApi === 'object') return directApi as OnlyOfficeEditorApi;

    const frameWindow = this.getNestedEditorWindow() as
      | (OnlyOfficeFrameWindow & {
          __onlyOfficeBrowserApplication?: {
            getController?: (name: string) => { api?: unknown } | undefined;
          };
        })
      | null;
    const nestedApi = frameWindow?.__onlyOfficeBrowserApplication?.getController?.('Main')?.api;
    return nestedApi && typeof nestedApi === 'object' ? (nestedApi as OnlyOfficeEditorApi) : null;
  }

  private installNativeDownloadAsInterceptor(): boolean {
    const api = this.getNativeEditorApi();
    if (!api || (typeof api._downloadAs !== 'function' && typeof api.asc_DownloadAs !== 'function')) return false;
    if (api.__onlyOfficeBrowserDownloadAsPatched) return true;

    if (typeof api.asc_DownloadAs === 'function') {
      const originalAscDownloadAs = api.asc_DownloadAs.bind(api);
      api.__onlyOfficeBrowserAscDownloadAsOriginal = originalAscDownloadAs;
      api.asc_DownloadAs = (options?: NativeDownloadAsOptions | string | number) => {
        const nativeOptions = options && typeof options === 'object' ? options : undefined;
        if (isOnlyOfficeNativePrintRequest({ options: nativeOptions })) {
          return originalAscDownloadAs(options);
        }

        const fileType = getNativeDownloadAsFileType(options);
        const event = {
          data: { fileType },
          nativeOptions,
        };
        if (isOnlyOfficeNativeSaveAsRequest({ options: nativeOptions })) {
          void this.handleRequestSaveAs(event);
        } else {
          void this.handleDownloadAs(event);
        }
        return true;
      };
    }

    if (typeof api._downloadAs === 'function') {
      const originalDownloadAs = api._downloadAs.bind(api);
      api.__onlyOfficeBrowserDownloadAsOriginal = originalDownloadAs;
      api._downloadAs = (
        actionType?: unknown,
        options?: NativeDownloadAsOptions,
        additionalData?: NativeDownloadAsAdditionalData,
        dataContainer?: NativeDownloadAsDataContainer,
        downloadType?: unknown,
      ) => {
        if (isOnlyOfficeNativePrintRequest({ actionType, options, additionalData, downloadType })) {
          return originalDownloadAs(actionType, options, additionalData, dataContainer, downloadType);
        }

        const fileType =
          typeof additionalData?.outputformat === 'string' || typeof additionalData?.outputformat === 'number'
            ? additionalData.outputformat
            : getNativeDownloadAsFileType(options);
        const title = typeof additionalData?.title === 'string' ? additionalData.title : undefined;
        const event = {
          data: { fileType, title },
          nativeOptions: options,
          additionalData,
          dataContainer,
          actionType,
          downloadType,
        };
        if (isOnlyOfficeNativeSaveAsRequest({ options, additionalData })) {
          void this.handleRequestSaveAs(event);
        } else {
          void this.handleDownloadAs(event);
        }
        return true;
      };
    }

    api.__onlyOfficeBrowserDownloadAsPatched = true;
    return true;
  }

  private installSpreadsheetPdfPrintPanelBridge(): boolean {
    if (getDocumentType(this.fileType) !== 'cell') return false;

    const frameWindow = this.getNestedEditorWindow();
    const notificationCenter = frameWindow?.Common?.NotificationCenter;
    if (!frameWindow || !notificationCenter || typeof notificationCenter.trigger !== 'function') return false;

    this.installSpreadsheetPdfPrintControllerPatch(frameWindow);

    if (!notificationCenter.__onlyOfficeBrowserSpreadsheetPdfPrintTriggerPatched) {
      const originalTrigger = notificationCenter.trigger.bind(notificationCenter);
      notificationCenter.__onlyOfficeBrowserSpreadsheetPdfPrintTriggerOriginal = originalTrigger;
      notificationCenter.trigger = (eventName?: unknown, ...args: unknown[]) => {
        if (eventName === 'file:print' && !frameWindow.__onlyOfficeBrowserOpeningSpreadsheetPdfPrintPanel) {
          this.closeSpreadsheetPdfPrintPanel(frameWindow, { restoreSourcePanel: false });
          frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode = null;
        }

        if (eventName === 'download:cancel') {
          this.closeSpreadsheetPdfPrintPanel(frameWindow, { closePreview: true });
          frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode = null;
        }

        if (eventName === 'download:settings') {
          const fileType = getSpreadsheetPdfExportFileType(args[1]);
          if (fileType !== null) {
            const isSaveAs = isTruthyOnlyOfficeFlag(args[2]);
            const mode: SpreadsheetPdfPrintExportMode = {
              fileType,
              isSaveAs,
              returnPanel: isSaveAs ? 'save-copy' : 'saveas',
              wopiSaveAsPath: args[3],
            };
            this.openSpreadsheetPdfPrintPanel(frameWindow, mode);
            return undefined;
          }
        }

        return originalTrigger(eventName, ...args);
      };
      notificationCenter.__onlyOfficeBrowserSpreadsheetPdfPrintTriggerPatched = true;
    }

    return true;
  }

  private installSpreadsheetPdfPrintControllerPatch(frameWindow: OnlyOfficeFrameWindow): void {
    const printController = getOnlyOfficeController<OnlyOfficePrintControllerLike>(frameWindow, 'Print');
    if (!printController || typeof printController.querySavePrintSettings !== 'function') return;
    this.installSpreadsheetPdfPrintSettingsShowPatch(frameWindow, printController);
    if (printController.__onlyOfficeBrowserSpreadsheetPdfPrintQueryPatched) return;

    const originalQuerySavePrintSettings = printController.querySavePrintSettings.bind(printController);
    printController.__onlyOfficeBrowserSpreadsheetPdfPrintQueryOriginal = originalQuerySavePrintSettings;
    printController.querySavePrintSettings = (action: unknown, useSystemDialog?: unknown) => {
      const mode = frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode;
      if (!mode || (action !== 'print' && action !== 'print-pdf')) {
        return originalQuerySavePrintSettings(action, useSystemDialog);
      }

      this.exportSpreadsheetPdfFromPrintPanel(frameWindow, printController, mode, useSystemDialog);
      return undefined;
    };
    printController.__onlyOfficeBrowserSpreadsheetPdfPrintQueryPatched = true;
  }

  private installSpreadsheetPdfPrintSettingsShowPatch(
    frameWindow: OnlyOfficeFrameWindow,
    printController: OnlyOfficePrintControllerLike,
  ): void {
    const printSettings = printController.printSettings;
    if (!printSettings || typeof printSettings.show !== 'function') return;
    if (printSettings.__onlyOfficeBrowserSpreadsheetPdfPrintShowPatched) return;

    const originalShow = printSettings.show.bind(printSettings);
    printSettings.__onlyOfficeBrowserSpreadsheetPdfPrintShowOriginal = originalShow;
    printSettings.show = () => {
      const result = originalShow();
      if (!frameWindow.__onlyOfficeBrowserShowingSpreadsheetPdfPrintPanel) {
        const state = frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintPanel;
        if (state?.printPanel) {
          delete state.printPanel.dataset.onlyofficeBrowserSpreadsheetPdfPrintPanel;
        }
        frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintPanel = null;
        frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode = null;
        setSpreadsheetPdfPrintPanelTitle(
          frameWindow,
          getSpreadsheetPdfNativePrintPanelTitleLabel(this.editorLang, printSettings),
        );
      }
      return result;
    };
    printSettings.__onlyOfficeBrowserSpreadsheetPdfPrintShowPatched = true;
  }

  private openSpreadsheetPdfPrintPanel(frameWindow: OnlyOfficeFrameWindow, mode: SpreadsheetPdfPrintExportMode): void {
    this.installSpreadsheetPdfPrintControllerPatch(frameWindow);
    this.closeSpreadsheetPdfPrintPanel(frameWindow, { closePreview: true, restoreSourcePanel: false });
    frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode = mode;

    const panel = this.ensureSpreadsheetPdfPrintPanel(frameWindow, mode);
    if (!isSpreadsheetPdfPrintPanelReady(panel)) {
      this.closeSpreadsheetPdfPrintPanel(frameWindow, { closePreview: true });
      throw new Error('OnlyOffice spreadsheet print preview panel is not available');
    }

    const state = this.attachSpreadsheetPdfPrintPanel(frameWindow, mode, panel);
    if (!state) {
      this.closeSpreadsheetPdfPrintPanel(frameWindow, { closePreview: true });
      throw new Error('OnlyOffice spreadsheet print preview panel is not available');
    }
    frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintPanel = state;

    this.applySpreadsheetPdfPrintPanelLabels(frameWindow);
    this.refreshSpreadsheetPdfPrintPreview(frameWindow);
    frameWindow.setTimeout?.(() => {
      this.applySpreadsheetPdfPrintPanelLabels(frameWindow);
      this.refreshSpreadsheetPdfPrintPreview(frameWindow);
    }, 0);
  }

  private ensureSpreadsheetPdfPrintPanel(
    frameWindow: OnlyOfficeFrameWindow,
    mode: SpreadsheetPdfPrintExportMode,
  ): HTMLElement | null {
    this.restoreSpreadsheetPdfSourcePanel(frameWindow, mode);

    let panel = getSpreadsheetPdfPrintPanelElement(frameWindow);
    if (isSpreadsheetPdfPrintPanelReady(panel)) return panel;

    const leftMenuController = getOnlyOfficeController<{
      clickToolbarPrint?: () => unknown;
      leftMenu?: { showMenu?: (name: string) => unknown };
    }>(frameWindow, 'LeftMenu');

    frameWindow.__onlyOfficeBrowserOpeningSpreadsheetPdfPrintPanel = true;
    try {
      if (typeof leftMenuController?.leftMenu?.showMenu === 'function') {
        leftMenuController.leftMenu.showMenu('file:printpreview');
      } else if (typeof leftMenuController?.clickToolbarPrint === 'function') {
        leftMenuController.clickToolbarPrint();
      } else {
        frameWindow.Common?.NotificationCenter?.__onlyOfficeBrowserSpreadsheetPdfPrintTriggerOriginal?.(
          'file:print',
          this,
        );
      }
    } finally {
      frameWindow.__onlyOfficeBrowserOpeningSpreadsheetPdfPrintPanel = false;
    }

    this.restoreSpreadsheetPdfSourcePanel(frameWindow, mode);
    panel = getSpreadsheetPdfPrintPanelElement(frameWindow);
    return panel;
  }

  private attachSpreadsheetPdfPrintPanel(
    frameWindow: OnlyOfficeFrameWindow,
    mode: SpreadsheetPdfPrintExportMode,
    panel: HTMLElement,
  ): SpreadsheetPdfPrintPanelState | null {
    if (!panel || !panel.parentNode) return null;

    const sourcePanel = getSpreadsheetPdfSourcePanelElement(frameWindow, mode);
    const state: SpreadsheetPdfPrintPanelState = {
      printPanel: panel,
      sourcePanel,
      previousPrintDisplay: panel.style.display,
      previousSourceDisplay: sourcePanel?.style.display || '',
    };

    if (sourcePanel && sourcePanel !== panel) {
      sourcePanel.style.display = 'none';
    }
    const printSettings = getOnlyOfficeController<OnlyOfficePrintControllerLike>(frameWindow, 'Print')?.printSettings;
    frameWindow.__onlyOfficeBrowserShowingSpreadsheetPdfPrintPanel = true;
    try {
      printSettings?.show?.();
    } finally {
      frameWindow.__onlyOfficeBrowserShowingSpreadsheetPdfPrintPanel = false;
    }
    panel.dataset.onlyofficeBrowserSpreadsheetPdfPrintPanel = 'true';
    panel.style.display = 'block';
    return state;
  }

  private closeSpreadsheetPdfPrintPanel(
    frameWindow: OnlyOfficeFrameWindow,
    options: { closePreview?: boolean; restoreSourcePanel?: boolean } = {},
  ): void {
    const state = frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintPanel;
    const mode = frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode;
    if (!state) return;

    frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintPanel = null;
    if (state.printPanel) {
      delete state.printPanel.dataset.onlyofficeBrowserSpreadsheetPdfPrintPanel;
      state.printPanel.style.display = state.previousPrintDisplay;
    }
    const printSettings = getOnlyOfficeController<OnlyOfficePrintControllerLike>(frameWindow, 'Print')?.printSettings;
    setSpreadsheetPdfPrintPanelTitle(
      frameWindow,
      getSpreadsheetPdfNativePrintPanelTitleLabel(this.editorLang, printSettings),
    );
    if (state.sourcePanel) {
      state.sourcePanel.style.display = state.previousSourceDisplay;
    }
    if (options.closePreview) {
      getOnlyOfficeController<OnlyOfficePrintControllerLike>(frameWindow, 'Print')?.onHidePrintMenu?.();
    }
    if (options.restoreSourcePanel !== false && mode && !state.sourcePanel) {
      this.restoreSpreadsheetPdfSourcePanel(frameWindow, mode);
    }
  }

  private restoreSpreadsheetPdfSourcePanel(
    frameWindow: OnlyOfficeFrameWindow,
    mode: SpreadsheetPdfPrintExportMode,
  ): void {
    const leftMenuController = getOnlyOfficeController<{
      leftMenu?: {
        menuFile?: { show?: (name?: string) => unknown };
        showMenu?: (name: string) => unknown;
      };
    }>(frameWindow, 'LeftMenu');
    if (typeof leftMenuController?.leftMenu?.menuFile?.show === 'function') {
      leftMenuController.leftMenu.menuFile.show(mode.returnPanel);
    } else if (mode.returnPanel === 'saveas' && typeof leftMenuController?.leftMenu?.showMenu === 'function') {
      leftMenuController.leftMenu.showMenu('file:saveas');
    }
  }

  private refreshSpreadsheetPdfPrintPreview(frameWindow: OnlyOfficeFrameWindow): void {
    const printController = getOnlyOfficeController<OnlyOfficePrintControllerLike>(frameWindow, 'Print');
    printController?.updatePrintRenderContainerSize?.(true);
    printController?.printSettings?.printScroller?.update?.();
  }

  private applySpreadsheetPdfPrintPanelLabels(frameWindow: OnlyOfficeFrameWindow): void {
    const mode = frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode;
    if (!mode) return;
    const printController = getOnlyOfficeController<OnlyOfficePrintControllerLike>(frameWindow, 'Print');
    const printSettings = printController?.printSettings;
    if (!printSettings) return;

    const actionCaption = getSpreadsheetPdfPrintActionLabel(this.editorLang, mode.isSaveAs);
    const titleCaption = getSpreadsheetPdfPrintPanelTitleLabel(this.editorLang, mode.isSaveAs);
    const saveSettingsCaption = getPrintSettingsSaveCaption(this.editorLang);
    setSpreadsheetPdfPrintPanelTitle(frameWindow, titleCaption);
    for (const button of printSettings.btnsPrint || []) {
      setOnlyOfficeButtonCaption(button, actionCaption);
    }
    for (const button of printSettings.btnsSave || []) {
      setOnlyOfficeButtonCaption(button, saveSettingsCaption);
    }
  }

  private exportSpreadsheetPdfFromPrintPanel(
    frameWindow: OnlyOfficeFrameWindow,
    printController: OnlyOfficePrintControllerLike,
    mode: SpreadsheetPdfPrintExportMode,
    useSystemDialog: unknown,
  ): void {
    const printSettings = printController.printSettings;
    const adjustPrintParams = printController.adjPrintParams;
    const api = printController.api;
    const DownloadOptions = frameWindow.Asc?.asc_CDownloadOptions;
    if (!printSettings || !adjustPrintParams || !api || typeof api.asc_DownloadAs !== 'function' || !DownloadOptions) {
      throw new Error('OnlyOffice spreadsheet print preview export is not available');
    }

    printController.savePageOptions?.(printSettings);
    printSettings.applySettings?.();

    const range = printSettings.getRange?.();
    adjustPrintParams.asc_setPrintType?.(range);
    adjustPrintParams.asc_setPageOptionsMap?.(printController._changedProps);
    adjustPrintParams.asc_setIgnorePrintArea?.(printSettings.getIgnorePrintArea?.());

    const printTypes = frameWindow.Asc?.c_oAscPrintType;
    const statusbarController = getOnlyOfficeController<{ getSelectTabs?: () => unknown }>(frameWindow, 'Statusbar');
    const activeSheets =
      range === printTypes?.Selection || range === printTypes?.ActiveSheets
        ? (statusbarController?.getSelectTabs?.() ?? null)
        : null;
    adjustPrintParams.asc_setActiveSheetsArray?.(activeSheets);

    let pagesFrom = getOnlyOfficeNumber(printSettings.getPagesFrom?.(), 0);
    let pagesTo = getOnlyOfficeNumber(printSettings.getPagesTo?.(), 0);
    if (pagesFrom > pagesTo) {
      const previousFrom = pagesFrom;
      pagesFrom = pagesTo;
      pagesTo = previousFrom;
    }
    adjustPrintParams.asc_setStartPageIndex?.(pagesFrom > 0 ? pagesFrom - 1 : null);
    adjustPrintParams.asc_setEndPageIndex?.(pagesTo > 0 ? pagesTo - 1 : null);

    const activeWorksheetIndex =
      range === printTypes?.EntireWorkbook ? 0 : getOnlyOfficeNumber(api.asc_getActiveWorksheetIndex?.(), 0);
    const changedProps = printController._changedProps;
    const changedPageOptions = Array.isArray(changedProps)
      ? changedProps[activeWorksheetIndex]
      : changedProps?.[String(activeWorksheetIndex)];
    const pageOptions = changedPageOptions || api.asc_getPageOptions?.(activeWorksheetIndex);
    const pageSetup = getOnlyOfficePageSetup(pageOptions);
    const pageWidth = getOnlyOfficeNumber(pageSetup?.asc_getWidth?.(), 0);
    const pageHeight = getOnlyOfficeNumber(pageSetup?.asc_getHeight?.(), 0);
    const orientationRecord = printSettings.cmbPaperOrientation?.getSelectedRecord?.();
    const paperOrientation =
      orientationRecord?.value === 'auto' ? 'auto' : pageSetup?.asc_getOrientation?.() ? 'landscape' : 'portrait';
    const printer = printSettings.cmbPrinter?.getSelectedRecord?.();
    const colorMode = printSettings.cmbColorPrinting?.getValue?.();
    adjustPrintParams.asc_setNativeOptions?.({
      usesystemdialog: Boolean(useSystemDialog),
      printer: printer ? printer.value : null,
      colorMode: colorMode === 'color',
      paperSize: {
        w: pageWidth,
        h: pageHeight,
        preset: printController.findPagePreset?.(printSettings, pageWidth, pageHeight),
      },
      paperOrientation,
      copies: getOnlyOfficeNumber(printSettings.spnCopies?.getNumberValue?.(), 1),
      sides: printSettings.cmbSides?.getValue?.() || 'one',
    });

    const options = new DownloadOptions(mode.fileType, mode.isSaveAs);
    options.asc_setAdvancedOptions?.(adjustPrintParams);
    options.asc_setIsSaveAs?.(mode.isSaveAs);
    if (mode.wopiSaveAsPath !== undefined) {
      options.asc_setWopiSaveAsPath?.(mode.wopiSaveAsPath);
    }

    this.closeSpreadsheetPdfPrintPanel(frameWindow, { closePreview: true });
    frameWindow.__onlyOfficeBrowserSpreadsheetPdfPrintExportMode = null;
    api.asc_DownloadAs(options);
    const toolbarView = getOnlyOfficeController<{ getView?: (name: string) => unknown }>(
      frameWindow,
      'Toolbar',
    )?.getView?.('Toolbar');
    frameWindow.Common?.NotificationCenter?.trigger?.('edit:complete', toolbarView);
  }

  private scheduleNativeDownloadAsInterceptor(): void {
    if (this.destroyed || this.installNativeDownloadAsInterceptor()) return;
    if (this.downloadAsInterceptorRetryId !== null) return;
    this.downloadAsInterceptorAttempts = 0;
    const retry = () => {
      this.downloadAsInterceptorRetryId = null;
      if (this.destroyed || this.installNativeDownloadAsInterceptor()) return;
      this.downloadAsInterceptorAttempts += 1;
      if (this.downloadAsInterceptorAttempts >= 30) return;
      this.downloadAsInterceptorRetryId = window.setTimeout(retry, 200);
    };
    this.downloadAsInterceptorRetryId = window.setTimeout(retry, 50);
  }

  private clearNativeDownloadAsInterceptorRetry(): void {
    if (this.downloadAsInterceptorRetryId === null) return;
    window.clearTimeout(this.downloadAsInterceptorRetryId);
    this.downloadAsInterceptorRetryId = null;
  }

  private notifyStateChange(): void {
    void Promise.resolve(this.options.onStateChange?.(this.getState(), this)).catch((error) => {
      this.options.onError?.(toError(error), this);
    });
  }

  async confirmSaveToNewFormat(options: OfficeSaveToNewFormatConfirmationOptions = {}): Promise<boolean> {
    const frameWindow = this.getNestedEditorWindow();
    const warning = frameWindow?.Common?.UI?.warning;
    if (!frameWindow || typeof warning !== 'function') {
      throw new Error('OnlyOffice built-in warning dialog is not available');
    }

    return new Promise<boolean>((resolve) => {
      let settled = false;
      const settle = (confirmed: boolean) => {
        if (settled) return;
        settled = true;
        resolve(confirmed);
      };

      warning.call(frameWindow.Common?.UI, {
        closable: false,
        width: 600,
        title: options.title || getSaveToNewFormatDialogTitle(frameWindow, this.editorLang),
        msg: options.message || getSaveToNewFormatDialogMessage(frameWindow, this.editorLang),
        buttons: ['ok', 'cancel'],
        dontshow: options.dontshow ?? true,
        callback: (result) => settle(result === 'ok'),
      });
    });
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

  private revokeCurrentEditorBinUrl(): void {
    if (!this.editorBinUrl) return;
    URL.revokeObjectURL(this.editorBinUrl);
    this.editorBinUrl = null;
  }

  private releasePrintResourceUrls(): void {
    for (const url of this.printResourceUrls) {
      if (url.startsWith('blob:')) {
        URL.revokeObjectURL(url);
      } else {
        void deleteCachedPrintPdf(url);
      }
    }
    this.printResourceUrls.clear();
  }

  private async remountCurrentEditor(): Promise<void> {
    this.cleanupSaveRequest(new Error('Editor was reopened before save completed'));
    this.clearNativeDownloadAsInterceptorRetry();
    this.clearNativeEditModeBridgeRetry();
    this.nestedFontPickerFilterCleanup?.();
    this.nestedFontPickerFilterCleanup = null;
    this.nativeEditModeBridgeCleanup?.();
    this.nativeEditModeBridgeCleanup = null;
    try {
      this.editor?.destroyEditor?.();
    } catch (error) {
      console.warn('Failed to destroy editor before reopening:', error);
    }
    this.editor = null;
    await this.blankAndRemoveIframes();
    this.placeholder.remove();
    this.container.replaceChildren();
    this.revokeCurrentEditorBinUrl();
    this.releasePrintResourceUrls();
    await this.mount();
  }

  private async handleRequestEditRights(): Promise<void> {
    if (this.destroyed || !this.previewMode || !this.previewEditAllowed || this.switchingPreviewMode) return;
    this.switchingPreviewMode = true;
    this.status = 'opening';
    this.editorMode = 'edit';
    this.readonlyMode.value = false;
    this.notifyStateChange();

    try {
      await this.remountCurrentEditor();
    } catch (error) {
      if (!this.destroyed) {
        this.status = 'error';
        this.options.onError?.(toError(error), this);
      }
    } finally {
      this.switchingPreviewMode = false;
    }
  }

  private async handleReturnToPreview(): Promise<void> {
    if (this.destroyed || this.previewMode || !this.previewEditAllowed || this.switchingPreviewMode) return;
    this.switchingPreviewMode = true;

    try {
      if (this.dirty) {
        await this.save(getFileExtension(this.fileName).toUpperCase());
      }
      if (this.destroyed) return;
      this.status = 'opening';
      this.editorMode = 'preview';
      this.readonlyMode.value = true;
      this.notifyStateChange();
      await this.remountCurrentEditor();
    } catch (error) {
      if (!this.destroyed) {
        this.options.onError?.(toError(error), this);
      }
    } finally {
      this.switchingPreviewMode = false;
    }
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

  private getNativePackageTargetExtension(): string | null {
    const documentType = getDocumentType(this.fileType);
    if (documentType === 'word') return 'docx';
    if (documentType === 'cell') return 'xlsx';
    if (documentType === 'slide') return 'pptx';
    return null;
  }

  private async convertSavedBin(data: Uint8Array | ArrayBuffer, targetFormat: string): Promise<File> {
    const sourceBytes = toUint8Array(data);
    const normalizedTargetFormat = normalizeDownloadTargetExtension(targetFormat);

    const directPackageExtension = this.getNativePackageTargetExtension();
    if (
      directPackageExtension &&
      normalizedTargetFormat === directPackageExtension &&
      looksLikeZipPackage(sourceBytes)
    ) {
      const fileName = replaceFileExtension(this.fileName, directPackageExtension);
      assertValidDownloadFileBytes(sourceBytes, fileName, directPackageExtension);
      return new File([sourceBytes as BlobPart], fileName, { type: getSavedFileMimeType(fileName) });
    }

    try {
      const result: BinConversionResult = await convertBinToDocument(
        sourceBytes,
        this.fileName,
        targetFormat,
        this.media,
      );
      let fileName = result.fileName;
      let bytes = toUint8Array(result.data);
      let validationTargetExt = normalizedTargetFormat;
      const markdownPackage =
        normalizedTargetFormat === 'md' ? packageMarkdownDataUriImages(fileName, bytes, this.fileName) : null;
      if (markdownPackage) {
        fileName = markdownPackage.fileName;
        bytes = markdownPackage.data;
        validationTargetExt = 'zip';
      }
      assertValidDownloadFileBytes(bytes, fileName, validationTargetExt);
      return new File([bytes as BlobPart], fileName, { type: getSavedFileMimeType(fileName) });
    } catch (error) {
      const header = Array.from(sourceBytes.slice(0, 16))
        .map((byte) => byte.toString(16).padStart(2, '0'))
        .join(' ');
      throw new Error(`${toError(error).message}; native header: ${header}`);
    }
  }

  private commitWordEditingBeforeHtmlExport(api: OnlyOfficeEditorApi): void {
    const frameWindow = this.getNestedEditorWindow();
    const keyboardEventCtor = frameWindow?.AscCommon?.CKeyboardEvent;
    const logicDocument = (
      api as {
        WordControl?: {
          m_oLogicDocument?: {
            OnKeyDown?: (event: unknown) => unknown;
          };
        };
      }
    ).WordControl?.m_oLogicDocument;
    if (typeof keyboardEventCtor !== 'function' || typeof logicDocument?.OnKeyDown !== 'function') return;

    try {
      const event = new keyboardEventCtor();
      event.CtrlKey = false;
      event.KeyCode = 27;
      logicDocument.OnKeyDown(event);
    } catch {
      // Best-effort parity with sdkjs' built-in HTML download path.
    }
  }

  private getNativeHtmlData(nativeOptions?: NativeDownloadAsOptions): string {
    const api = this.getNativeEditorApi();
    if (!api || typeof api.asc_nativeGetHtml !== 'function') {
      throw new Error('OnlyOffice native HTML export is not available');
    }

    this.commitWordEditingBeforeHtmlExport(api);
    if (typeof api.asc_nativeCalculateFile === 'function') {
      api.asc_nativeCalculateFile(nativeOptions || null);
    }

    const html = api.asc_nativeGetHtml(nativeOptions || null);
    if (typeof html !== 'string' || !html.trim()) {
      throw new Error('OnlyOffice native HTML export produced an empty document');
    }
    return html;
  }

  private async convertNativeHtmlDownloadAsFile(
    targetExt: string,
    nativeOptions?: NativeDownloadAsOptions,
  ): Promise<File> {
    const normalizedTargetExt = normalizeDownloadTargetExtension(targetExt);
    const htmlBytes = utf8Bytes(this.getNativeHtmlData(nativeOptions));
    if (normalizedTargetExt === 'html') {
      const fileName = replaceFileExtension(this.fileName, 'html');
      assertValidDownloadFileBytes(htmlBytes, fileName, normalizedTargetExt);
      return new File([htmlBytes as BlobPart], fileName, { type: getSavedFileMimeType(fileName) });
    }

    const result = await convertHtmlToDocument(htmlBytes, this.fileName, normalizedTargetExt);
    let fileName = result.fileName;
    let bytes = toUint8Array(result.data);
    let validationTargetExt = normalizedTargetExt;

    const markdownPackage =
      normalizedTargetExt === 'md' ? packageMarkdownDataUriImages(fileName, bytes, this.fileName) : null;
    if (markdownPackage) {
      fileName = markdownPackage.fileName;
      bytes = markdownPackage.data;
      validationTargetExt = 'zip';
    }

    assertValidDownloadFileBytes(bytes, fileName, validationTargetExt);
    return new File([bytes as BlobPart], fileName, { type: getSavedFileMimeType(fileName) });
  }

  private async convertNativeOutputFile(
    nativeData: Uint8Array | ArrayBuffer,
    targetFormat: string,
    nativeOptions?: NativeDownloadAsOptions,
  ): Promise<File> {
    const normalizedTargetExt = normalizeDownloadTargetExtension(targetFormat);
    if (NATIVE_HTML_DOWNLOAD_EXTENSIONS.has(normalizedTargetExt)) {
      return this.convertNativeHtmlDownloadAsFile(normalizedTargetExt, nativeOptions);
    }
    return this.convertSavedBin(nativeData, targetFormat);
  }

  private async convertDownloadAsFile(targetExt: string, nativeOptions?: NativeDownloadAsOptions): Promise<File> {
    const normalizedTargetExt = normalizeDownloadTargetExtension(targetExt);
    if (NATIVE_HTML_DOWNLOAD_EXTENSIONS.has(normalizedTargetExt)) {
      return this.convertNativeHtmlDownloadAsFile(normalizedTargetExt, nativeOptions);
    }

    const usesNativePrintRenderer =
      normalizedTargetExt === 'pdf' ||
      normalizedTargetExt === 'pdfa' ||
      normalizedTargetExt === 'jpg' ||
      normalizedTargetExt === 'jpeg' ||
      normalizedTargetExt === 'png';
    const isPdfTarget = normalizedTargetExt === 'pdf' || normalizedTargetExt === 'pdfa';
    if (usesNativePrintRenderer) {
      const nativePrintBytes = this.getNativePrintData(isPdfTarget ? nativeOptions : undefined);
      if (!nativePrintBytes) {
        if (isPdfTarget && hasNativeDownloadPrintOptions(nativeOptions)) {
          throw new Error('OnlyOffice PDF export settings could not be applied');
        }
        throw new Error('OnlyOffice native print export is not available');
      }
      const outputName = getDownloadFileName(this.fileName, normalizedTargetExt);
      try {
        let pdfBytes = nativePrintBytes;
        if (!looksLikePdfDocument(pdfBytes)) {
          const pdfResult = await convertPrintDataToPdf(nativePrintBytes, this.fileName, this.media);
          pdfBytes = toUint8Array(pdfResult.data);
        }

        if (isPdfTarget) {
          assertValidDownloadFileBytes(pdfBytes, outputName, normalizedTargetExt);
          return new File([pdfBytes as BlobPart], outputName, { type: getSavedFileMimeType(outputName) });
        }

        const imageResult = await convertPdfToImage(pdfBytes, this.fileName, normalizedTargetExt, {
          allPages: (getPdfPageCount(pdfBytes) || 1) > 1,
        });
        const imageBytes = toUint8Array(imageResult.data);
        assertValidDownloadFileBytes(imageBytes, imageResult.fileName, normalizedTargetExt);
        return new File([imageBytes as BlobPart], imageResult.fileName, {
          type: getSavedFileMimeType(imageResult.fileName),
        });
      } catch (error) {
        if (isPdfTarget && hasNativeDownloadPrintOptions(nativeOptions)) {
          throw new Error(`OnlyOffice PDF export settings could not be applied: ${toError(error).message}`);
        }
        throw error;
      }
    }

    return this.convertNativeOutputFile(this.getNativeSaveData(), normalizedTargetExt.toUpperCase(), nativeOptions);
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
    this.binData = await file.arrayBuffer();
    this.setDirty(false);
    this.resolveSaveRequest(file);
  }

  private async emitDownloadedFile(file: File): Promise<void> {
    if (this.options.onDownload) {
      await this.options.onDownload(file, this);
      return;
    }
    downloadFile(file);
  }

  private async emitSaveAsFile(file: File): Promise<void> {
    if (this.options.onSaveAs) {
      await this.options.onSaveAs(file, this);
      return;
    }
    downloadFile(file);
  }

  private async createDownloadAsFileFromEvent(event: DownloadAsEvent): Promise<File> {
    const targetExt = getDownloadTargetExtension(
      event.data?.fileType,
      event.data?.url,
      getFileExtension(this.fileName),
    );
    const url = event.data?.url;
    if (!url) {
      const file = await this.convertDownloadAsFile(targetExt, event.nativeOptions);
      return withFileName(file, event.data?.title);
    }

    const response = await fetch(url);
    if (!response.ok) {
      return this.convertDownloadAsFile(targetExt, event.nativeOptions);
    }

    const blob = await response.blob();
    const fileName = event.data?.title || getDownloadFileName(this.fileName, targetExt);
    assertValidDownloadFileBytes(toUint8Array(await blob.arrayBuffer()), fileName, targetExt);
    return new File([blob], fileName, { type: blob.type || getSavedFileMimeType(fileName) });
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
      const file = await this.convertNativeOutputFile(nativeData, targetFormat);
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
      const file = await this.createDownloadAsFileFromEvent(event);
      await this.emitDownloadedFile(file);
    } catch (error) {
      const normalized = toError(error);
      this.rejectSaveRequest(normalized);
      this.options.onError?.(normalized, this);
    }
  }

  private async handleRequestSaveAs(event: RequestSaveAsEvent): Promise<void> {
    try {
      const file = await this.createDownloadAsFileFromEvent(event);
      await this.emitSaveAsFile(file);
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

  getNativePrintData(downloadOptions?: NativeDownloadAsOptions): Uint8Array | null {
    if (!this.editor || typeof this.editor.asc_nativeGetPDF !== 'function') return null;

    const runtimeWindow =
      (this.editor.getEditorWindow?.() as RuntimeWindow | null | undefined) || (window as RuntimeWindow);
    const hadNative = Object.prototype.hasOwnProperty.call(runtimeWindow, 'native');
    const nativeBridge = runtimeWindow.native || {};
    const hadSaveEnd = Object.prototype.hasOwnProperty.call(nativeBridge, 'Save_End');
    const originalSaveEnd = nativeBridge.Save_End;
    const advancedOptions = getNativeDownloadAdvancedOptions(downloadOptions);
    const useSpreadsheetNativePrintOptions = getDocumentType(this.fileType) === 'cell' && advancedOptions !== undefined;
    const hadDesktopPrintOptions = Object.prototype.hasOwnProperty.call(runtimeWindow, 'AscDesktopEditor_PrintOptions');
    const originalDesktopPrintOptions = runtimeWindow.AscDesktopEditor_PrintOptions;
    let nativeLength: number | undefined;

    runtimeWindow.native = nativeBridge;
    nativeBridge.Save_End = (header, length) => {
      if (Number.isFinite(length)) {
        nativeLength = Math.max(0, Math.floor(length));
      }
      return originalSaveEnd?.call(nativeBridge, header, length);
    };

    try {
      let printOptions: Record<string, unknown> | undefined = { isPrint: true };
      if (useSpreadsheetNativePrintOptions) {
        // Spreadsheet asc_nativePrint reads AscDesktopEditor_PrintOptions.advancedOptions;
        // passing the same object via asc_nativeGetPDF(options) makes it reset to EntireWorkbook.
        runtimeWindow.AscDesktopEditor_PrintOptions = { advancedOptions };
        printOptions = undefined;
      } else if (advancedOptions !== undefined) {
        printOptions.advancedOptions = advancedOptions;
        printOptions.printOptions = advancedOptions;
      }
      const payload: NativePrintPdfData = this.editor.asc_nativeGetPDF(printOptions);
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
      if (hadDesktopPrintOptions) {
        runtimeWindow.AscDesktopEditor_PrintOptions = originalDesktopPrintOptions;
      } else {
        delete runtimeWindow.AscDesktopEditor_PrintOptions;
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
      const file = await this.convertNativeOutputFile(this.getNativeSaveData(), targetFormat);
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
      void this.handleRequestEditRights();
      return;
    }
    if (!this.previewMode && readonly && this.previewEditAllowed) {
      void this.handleReturnToPreview();
      return;
    }
    this.readonlyMode.value = readonly;
    this.applyReadonlyState();
  }

  getState(): OfficeEditorState {
    const mode: OfficeEditorMode =
      this.editorMode === 'preview' ? 'preview' : this.readonlyMode.value ? 'readonly' : 'edit';
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
    this.clearNativeDownloadAsInterceptorRetry();
    this.clearNativeEditModeBridgeRetry();
    this.nestedFontPickerFilterCleanup?.();
    this.nestedFontPickerFilterCleanup = null;
    this.nativeEditModeBridgeCleanup?.();
    this.nativeEditModeBridgeCleanup = null;
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

    this.revokeCurrentEditorBinUrl();
    if (this.sourceObjectUrl) {
      URL.revokeObjectURL(this.sourceObjectUrl);
      this.sourceObjectUrl = null;
    }
    this.releasePrintResourceUrls();

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
        hideIframeForTeardown(frame);
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
