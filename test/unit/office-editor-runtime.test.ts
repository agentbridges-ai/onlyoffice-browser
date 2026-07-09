import { beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
	  assertGeneratedFontAssetsAvailable: vi.fn(),
	  convertBinToDocument: vi.fn(),
	  convertDocument: vi.fn(),
	  convertHtmlToDocument: vi.fn(),
	  convertPdfToImage: vi.fn(),
	  convertPrintDataToPdf: vi.fn(),
	  initX2T: vi.fn(),
}));

vi.mock('../../src/lib/font-assets', () => ({
  assertGeneratedFontAssetsAvailable: mocks.assertGeneratedFontAssetsAvailable,
}));

vi.mock('../../src/lib/converter', () => ({
	  convertBinToDocument: mocks.convertBinToDocument,
	  convertDocument: mocks.convertDocument,
	  convertHtmlToDocument: mocks.convertHtmlToDocument,
	  convertPdfToImage: mocks.convertPdfToImage,
	  convertPrintDataToPdf: mocks.convertPrintDataToPdf,
	  initX2T: mocks.initX2T,
}));

import { createOfficeEditor } from '../../src/lib/office-editor-runtime';

function flush(): Promise<void> {
  return Promise.resolve().then(() => undefined);
}

function waitForMessage(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

async function waitForCondition(predicate: () => boolean, timeoutMs = 1_000): Promise<void> {
  const startedAt = Date.now();
  while (!predicate()) {
    if (Date.now() - startedAt > timeoutMs) {
      throw new Error('Timed out waiting for condition');
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

function asciiBytes(value: string): Uint8Array {
  return Uint8Array.from(value, (char) => char.charCodeAt(0) & 0xff);
}

function binaryString(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  let output = '';
  for (let index = 0; index < bytes.length; index += chunkSize) {
    output += String.fromCharCode(...bytes.slice(index, index + chunkSize));
  }
  return output;
}

function pdfUtf16BeHex(value: string): string {
  const bytes = [0xfe, 0xff];
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes.push((code >> 8) & 0xff, code & 0xff);
  }
  return bytes.map((byte) => byte.toString(16).padStart(2, '0').toUpperCase()).join('');
}

function validOnePagePdfFixture(): Uint8Array {
  return asciiBytes(`%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Count 1 /Kids [3 0 R] >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 10 10] >>
endobj
trailer
<< /Root 1 0 R /Size 4 >>
startxref
0
%%EOF
`);
}

function zipFixture(): Uint8Array {
  return new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14, 0, 0, 0]);
}

function pngFixture(): Uint8Array {
  return new Uint8Array([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x00,
  ]);
}

function jpegFixture(): Uint8Array {
  return new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 0x00, 0x43, 0x00, 0xff, 0xd9]);
}

type CapturedDocEditorConfig = {
  type?: 'desktop' | 'embedded' | 'mobile';
  document: {
    permissions: {
      edit: boolean;
      download?: boolean;
    };
  };
  editorConfig: {
    mode?: 'edit' | 'view';
    user?: {
      id: string;
      name: string;
    };
    coEditing?: {
      mode?: 'fast' | 'strict';
      change?: boolean;
    };
    embedded?: {
      autostart?: 'document' | 'player';
      toolbarDocked?: 'top' | 'bottom';
    };
    customization: {
      zoom?: number;
      spellcheck?: boolean;
      uiTheme?: 'theme-system' | 'theme-white' | 'theme-night';
      autosave?: boolean;
      forcesave?: boolean;
      features: {
        spellcheck: {
          change: boolean;
        };
      };
      anonymous: {
        request: boolean;
        label: string;
      };
    };
  };
  events: {
    onAppReady: () => void;
    onDocumentReady: () => void;
    onSave?: (event: {
      data: {
        data: {
          data: Uint8Array;
        };
        option: {
          outputformat: number;
        };
      };
    }) => void;
    onDocumentStateChange?: (event: boolean | { data?: boolean }) => void;
    onDownloadAs?: (event: { data?: { url?: string; fileType?: string | number; title?: string } }) => void;
    onRequestSaveAs?: (event: { data?: { url?: string; fileType?: string | number; title?: string } }) => void;
    onRequestEditRights?: () => void;
  };
};

type CapturedDocEditorInstance = {
  connectMockServer: ReturnType<typeof vi.fn>;
  destroyEditor: ReturnType<typeof vi.fn>;
  downloadAs: ReturnType<typeof vi.fn>;
  nativeDownloadAs: ReturnType<typeof vi.fn>;
  nativeDownloadAsHandler: ReturnType<typeof vi.fn>;
  getEditorApi: ReturnType<typeof vi.fn>;
	    nativeApi: {
	      asc_DownloadAs: (options?: unknown) => unknown;
	      asc_nativeCalculateFile: ReturnType<typeof vi.fn>;
	      asc_nativeGetHtml: ReturnType<typeof vi.fn>;
	      _downloadAs: (
      actionType?: unknown,
      options?: unknown,
      additionalData?: { outputformat?: unknown; title?: unknown; inline?: unknown; isSaveAs?: unknown },
      dataContainer?: unknown,
      downloadType?: unknown,
	    ) => boolean | undefined;
	    __onlyOfficeBrowserDownloadAsPatched?: boolean;
	    __onlyOfficeBrowserAscDownloadAsOriginal?: unknown;
	    __onlyOfficeBrowserDownloadAsOriginal?: unknown;
	  };
  processRightsChange: ReturnType<typeof vi.fn>;
	  asc_nativeGetFile3: ReturnType<typeof vi.fn>;
	  asc_nativeGetPDF: ReturnType<typeof vi.fn>;
	  asc_nativeCalculateFile: ReturnType<typeof vi.fn>;
	  asc_nativeGetHtml: ReturnType<typeof vi.fn>;
	  zoomFitToWidth: ReturnType<typeof vi.fn>;
	};

describe('office editor runtime', () => {
  const docEditorConfigs: CapturedDocEditorConfig[] = [];
  const docEditorInstances: CapturedDocEditorInstance[] = [];
  const themeSetMocks: ReturnType<typeof vi.fn>[] = [];
  const themeControllers: Array<{ map: () => Record<string, unknown>; setTheme: ReturnType<typeof vi.fn> }> = [];

  beforeEach(() => {
    document.head.innerHTML = '';
    document.body.innerHTML = '';
    window.localStorage.clear();
    docEditorConfigs.length = 0;
    docEditorInstances.length = 0;
    themeSetMocks.length = 0;
    themeControllers.length = 0;
    vi.clearAllMocks();
    Object.defineProperty(URL, 'createObjectURL', {
      configurable: true,
      value: vi.fn(() => 'blob:onlyoffice-browser-download'),
    });
    Object.defineProperty(URL, 'revokeObjectURL', {
      configurable: true,
      value: vi.fn(),
    });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
    mocks.assertGeneratedFontAssetsAvailable.mockResolvedValue(undefined);
    mocks.initX2T.mockResolvedValue(undefined);
    mocks.convertDocument.mockImplementation(async (file: File) => ({
      fileName: file.name,
      bin: new Uint8Array([1, 2, 3]),
      media: {},
    }));
	    mocks.convertBinToDocument.mockImplementation(async (_bin: Uint8Array, fileName: string, targetExt = 'DOCX') => {
      const normalizedTargetExt = targetExt.toLowerCase() === 'pdfa' ? 'pdf' : targetExt.toLowerCase();
      const outputExt = normalizedTargetExt === 'jpeg' ? 'jpg' : normalizedTargetExt;
      const dataByExtension: Record<string, Uint8Array> = {
        docm: zipFixture(),
        docx: zipFixture(),
        docxf: zipFixture(),
        dotx: zipFixture(),
        epub: zipFixture(),
        odp: zipFixture(),
        ods: zipFixture(),
        odt: zipFixture(),
        oform: zipFixture(),
        otp: zipFixture(),
        ots: zipFixture(),
        ott: zipFixture(),
        potm: zipFixture(),
        potx: zipFixture(),
        ppsm: zipFixture(),
        ppsx: zipFixture(),
        pptm: zipFixture(),
        pptx: zipFixture(),
        xlsb: zipFixture(),
        xlsm: zipFixture(),
        xlsx: zipFixture(),
        xltm: zipFixture(),
        xltx: zipFixture(),
        pdf: validOnePagePdfFixture(),
        jpg: jpegFixture(),
        jpeg: jpegFixture(),
        png: pngFixture(),
        rtf: asciiBytes('{\\rtf1\\ansi alpha}'),
        html: asciiBytes('<!doctype html><html><body>alpha</body></html>'),
        fb2: asciiBytes('<?xml version="1.0"?><FictionBook></FictionBook>'),
        md: asciiBytes('# alpha\n'),
        txt: asciiBytes('alpha\n'),
        csv: asciiBytes('alpha\n'),
      };
      return {
        fileName: fileName.replace(/\.[^/.]+$/, `.${outputExt}`),
        data: dataByExtension[normalizedTargetExt] || asciiBytes('alpha\n'),
	      };
	    });
	    mocks.convertHtmlToDocument.mockImplementation(async (_html: Uint8Array, fileName: string, targetExt: string) => {
	      const normalizedTargetExt = targetExt.toLowerCase();
	      const dataByExtension: Record<string, Uint8Array> = {
	        epub: zipFixture(),
	        fb2: asciiBytes('<?xml version="1.0"?><FictionBook></FictionBook>'),
	        html: asciiBytes('<!doctype html><html><body>alpha</body></html>'),
	        md: asciiBytes('# alpha\n'),
	      };
	      return {
	        fileName: fileName.replace(/\.[^/.]+$/, `.${normalizedTargetExt}`),
	        data: dataByExtension[normalizedTargetExt] || asciiBytes('alpha\n'),
	      };
	    });
    mocks.convertPrintDataToPdf.mockImplementation(async (_printData: Uint8Array, fileName: string) => ({
      fileName: fileName.replace(/\.[^/.]+$/, '.pdf'),
      data: asciiBytes('%PDF-1.7\n%%EOF\n'),
    }));
    mocks.convertPdfToImage.mockImplementation(async (_pdfData: Uint8Array, fileName: string, targetExt: string) => {
      const normalizedTargetExt = targetExt.toLowerCase() === 'jpeg' ? 'jpg' : targetExt.toLowerCase();
      return {
        fileName: fileName.replace(/\.[^/.]+$/, `.${normalizedTargetExt}`),
        data: normalizedTargetExt === 'png' ? pngFixture() : jpegFixture(),
      };
    });

    const cacheEntries = new Map<string, Response>();
    const printCache = {
      put: vi.fn(async (request: RequestInfo | URL, response: Response) => {
        const url = typeof request === 'string' ? request : request instanceof URL ? request.href : request.url;
        cacheEntries.set(url, response.clone());
      }),
      match: vi.fn(async (request: RequestInfo | URL) => {
        const url = typeof request === 'string' ? request : request instanceof URL ? request.href : request.url;
        return cacheEntries.get(url);
      }),
      delete: vi.fn(async (request: RequestInfo | URL) => {
        const url = typeof request === 'string' ? request : request instanceof URL ? request.href : request.url;
        return cacheEntries.delete(url);
      }),
    };
    Object.defineProperty(window, 'caches', {
      configurable: true,
      value: {
        open: vi.fn(async () => printCache),
      },
    });
    Object.defineProperty(navigator, 'serviceWorker', {
      configurable: true,
      value: {
        controller: {},
        register: vi.fn(async () => ({})),
        ready: Promise.resolve({}),
        addEventListener: vi.fn(),
        removeEventListener: vi.fn(),
      },
    });

	    class MockDocEditor {
	      connectMockServer = vi.fn();
	      destroyEditor = vi.fn();
	      downloadAs = vi.fn();
	      nativeDownloadAs = vi.fn();
	      nativeDownloadAsHandler = vi.fn();
	      asc_nativeCalculateFile = vi.fn();
	      asc_nativeGetHtml = vi.fn(() => '<!doctype html><html><body>alpha</body></html>');
	      nativeApi = {
	        asc_DownloadAs: (options?: unknown) =>
	          this.nativeApi._downloadAs(1, options, {
	            outputformat:
              options && typeof options === 'object' && 'fileType' in options
                ? (options as { fileType?: unknown }).fileType
                : undefined,
	            title: 'alpha.out',
	            isSaveAs:
              options && typeof options === 'object' && 'isSaveAs' in options
                ? (options as { isSaveAs?: unknown }).isSaveAs
                : undefined,
	          }),
	        asc_nativeCalculateFile: this.asc_nativeCalculateFile,
	        asc_nativeGetHtml: this.asc_nativeGetHtml,
	        _downloadAs: this.nativeDownloadAsHandler,
	      };
      getEditorApi = vi.fn(() => this.nativeApi);
      processRightsChange = vi.fn();
      asc_nativeGetFile3 = vi.fn(() => ({ data: new Uint8Array([9, 8, 7]) }));
      asc_nativeGetPDF = vi.fn();
      zoomFitToWidth = vi.fn();

      constructor(elementId: string, config: CapturedDocEditorConfig) {
        docEditorConfigs.push(config);
        docEditorInstances.push(this);
        const frame = document.createElement('iframe');
        frame.name = 'frameEditor';
        document.getElementById(elementId)?.appendChild(frame);
        const setTheme = vi.fn();
        const themeMap = {
          'theme-system': { text: 'System' },
          'theme-light': { text: 'Light' },
          'theme-dark': { text: 'Dark' },
          'theme-white': { text: 'Modern light' },
          'theme-night': { text: 'Modern dark' },
        };
        const themes = {
          map: vi.fn(() => themeMap),
          get: vi.fn((theme: string) => themeMap[theme as keyof typeof themeMap]),
          setTheme,
          defaultThemeId: vi.fn(() => 'theme-light'),
          defaultTheme: vi.fn(() => themeMap['theme-light']),
        };
        themeSetMocks.push(setTheme);
        themeControllers.push(themes);
        Object.defineProperty(frame.contentWindow, 'Common', {
          configurable: true,
          writable: true,
          value: {
            UI: {
              Themes: themes,
            },
          },
        });
        queueMicrotask(() => {
          config.events.onAppReady();
          config.events.onDocumentReady();
        });
      }
    }

    (window as typeof window & { DocsAPI?: unknown }).DocsAPI = {
      DocEditor: MockDocEditor,
    };
  });

  function installSpreadsheetPrintPreviewHarness(frameWindow: Window) {
    class MockDownloadOptions {
      fileType: unknown;
      asUrl: unknown;
      isSaveAs = false;
      wopiSaveAsPath: unknown;
      advancedOptions: unknown;

      constructor(fileType?: unknown, asUrl?: unknown) {
        this.fileType = fileType;
        this.asUrl = asUrl;
      }

      asc_setAdvancedOptions = vi.fn((options: unknown) => {
        this.advancedOptions = options;
      });

      asc_setIsSaveAs = vi.fn((isSaveAs: boolean) => {
        this.isSaveAs = isSaveAs;
      });

      asc_setWopiSaveAsPath = vi.fn((path: unknown) => {
        this.wopiSaveAsPath = path;
      });

      asc_getFileType = vi.fn(() => this.fileType);
      asc_getIsSaveAs = vi.fn(() => this.isSaveAs);
    }

    const originalTrigger = vi.fn();
    const fileMenuPanel = frameWindow.document.createElement('div');
    fileMenuPanel.id = 'file-menu-panel';
    const panelContext = frameWindow.document.createElement('div');
    panelContext.className = 'panel-context';
    const panelSaveAs = frameWindow.document.createElement('div');
    panelSaveAs.id = 'panel-saveas';
    panelSaveAs.className = 'content-box';
    panelSaveAs.style.display = 'block';
    const panelSaveCopy = frameWindow.document.createElement('div');
    panelSaveCopy.id = 'panel-savecopy';
    panelSaveCopy.className = 'content-box';
    panelSaveCopy.style.display = 'none';
    const printPanel = frameWindow.document.createElement('div');
    printPanel.id = 'panel-print';
    printPanel.className = 'content-box';
    printPanel.style.display = 'none';
    const printSettingsHost = frameWindow.document.createElement('div');
    printSettingsHost.id = 'id-print-settings';
    const printSettingsContainer = frameWindow.document.createElement('div');
    printSettingsContainer.className = 'print-settings';
    const printMainHeader = frameWindow.document.createElement('div');
    printMainHeader.className = 'main-header';
    printMainHeader.textContent = '打印';
    const settingsContainer = frameWindow.document.createElement('div');
    settingsContainer.className = 'settings-container';
    printSettingsContainer.appendChild(printMainHeader);
    printSettingsContainer.appendChild(settingsContainer);
    printSettingsHost.appendChild(printSettingsContainer);
    const printPreviewBox = frameWindow.document.createElement('div');
    printPreviewBox.id = 'print-preview-box';
    const printPreviewWrapper = frameWindow.document.createElement('div');
    printPreviewWrapper.id = 'print-preview-wrapper';
    const printPreview = frameWindow.document.createElement('div');
    printPreview.id = 'print-preview';
    printPreviewWrapper.appendChild(printPreview);
    const printNavigation = frameWindow.document.createElement('div');
    printNavigation.id = 'print-navigation';
    printPreviewBox.append(printPreviewWrapper, printNavigation);
    const printPreviewEmpty = frameWindow.document.createElement('div');
    printPreviewEmpty.id = 'print-preview-empty';
    printPanel.append(printSettingsHost, printPreviewBox, printPreviewEmpty);
    panelContext.append(panelSaveAs, panelSaveCopy, printPanel);
    fileMenuPanel.appendChild(panelContext);
    frameWindow.document.body.appendChild(fileMenuPanel);
    const showMenu = vi.fn();
    const showFilePanel = vi.fn((name?: string) => {
      for (const panel of [panelSaveAs, panelSaveCopy, printPanel]) {
        panel.style.display = 'none';
      }
      if (name === 'save-copy') {
        panelSaveCopy.style.display = 'block';
      } else if (name === 'printpreview') {
        printPanel.style.display = 'block';
      } else {
        panelSaveAs.style.display = 'block';
      }
    });
    showMenu.mockImplementation((name: string) => {
      if (name === 'file:printpreview') showFilePanel('printpreview');
      if (name === 'file:saveas') showFilePanel('saveas');
    });
    const ascDownloadAs = vi.fn();
    const pageSetup = {
      asc_getWidth: vi.fn(() => 297),
      asc_getHeight: vi.fn(() => 420),
      asc_getOrientation: vi.fn(() => true),
    };
    const pageOptions = {
      asc_getPageSetup: vi.fn(() => pageSetup),
    };
    const printSettingsShow = vi.fn(() => {
      printPanel.style.display = 'block';
    });
    const printSettings = {
      show: printSettingsShow,
      txtPrint: '打印',
      btnsPrint: [{ setCaption: vi.fn() }],
      btnsSave: [{ setCaption: vi.fn() }],
      applySettings: vi.fn(),
      getRange: vi.fn(() => 2),
      getIgnorePrintArea: vi.fn(() => true),
      getPagesFrom: vi.fn(() => 1),
      getPagesTo: vi.fn(() => 4),
      cmbPaperOrientation: { getSelectedRecord: vi.fn(() => ({ value: 'landscape' })) },
      cmbPrinter: { getSelectedRecord: vi.fn(() => ({ value: 'printer-1' })) },
      cmbColorPrinting: { getValue: vi.fn(() => 'color') },
      spnCopies: { getNumberValue: vi.fn(() => 1) },
      cmbSides: { getValue: vi.fn(() => 'one') },
      printScroller: { update: vi.fn() },
    };
    const adjustPrintParams = {
      asc_setPrintType: vi.fn(),
      asc_setPageOptionsMap: vi.fn(),
      asc_setIgnorePrintArea: vi.fn(),
      asc_setActiveSheetsArray: vi.fn(),
      asc_setStartPageIndex: vi.fn(),
      asc_setEndPageIndex: vi.fn(),
      asc_setNativeOptions: vi.fn(),
    };
    const printController = {
      printSettings,
      adjPrintParams: adjustPrintParams,
      _changedProps: [] as unknown[],
      api: {
        asc_getActiveWorksheetIndex: vi.fn(() => 0),
        asc_getPageOptions: vi.fn(() => pageOptions),
        asc_DownloadAs: ascDownloadAs,
      },
      savePageOptions: vi.fn(),
      findPagePreset: vi.fn(() => 'A3'),
      querySavePrintSettings: vi.fn(),
      onHidePrintMenu: vi.fn(),
      updatePrintRenderContainerSize: vi.fn(),
    };
    const leftMenuController = {
      leftMenu: { showMenu, menuFile: { show: showFilePanel } },
      clickToolbarPrint: vi.fn(() => showMenu('file:printpreview')),
    };
    const toolbarController = {
      getView: vi.fn(() => ({ id: 'toolbar' })),
    };
    const statusbarController = {
      getSelectTabs: vi.fn(() => [0]),
    };

    Object.assign(frameWindow, {
      Common: {
        NotificationCenter: {
          trigger: originalTrigger,
        },
      },
      Asc: {
        c_oAscPrintType: {
          Selection: 1,
          ActiveSheets: 2,
          EntireWorkbook: 3,
        },
        asc_CDownloadOptions: MockDownloadOptions,
      },
      SSE: {
        getController: vi.fn((name: string) => {
          if (name === 'Print') return printController;
          if (name === 'LeftMenu') return leftMenuController;
          if (name === 'Toolbar') return toolbarController;
          if (name === 'Statusbar') return statusbarController;
          return null;
        }),
      },
    });

    return {
      MockDownloadOptions,
      originalTrigger,
      showMenu,
      showFilePanel,
      printSettingsShow,
      ascDownloadAs,
      printSettings,
      adjustPrintParams,
      printController,
      statusbarController,
      printPanel,
      printMainHeader,
      panelContext,
      panelSaveAs,
      panelSaveCopy,
    };
  }

  it('disables autosave, forcesave, and spellcheck by default', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });

    expect(docEditorConfigs).toHaveLength(1);
    expect(docEditorConfigs[0].editorConfig.customization).toMatchObject({
      autosave: false,
      compactToolbar: true,
      forcesave: false,
      spellcheck: false,
      features: {
        featuresTips: false,
        spellcheck: {
          change: false,
        },
      },
    });
    expect(docEditorConfigs[0].editorConfig.coEditing).toEqual({
      mode: 'strict',
      change: false,
    });

    await instance.destroy();
  });

  it('defaults the OnlyOffice interface to the modern system theme', async () => {
    window.localStorage.setItem('ui-theme-id', 'theme-classic-light');
    window.localStorage.setItem('ui-theme', JSON.stringify({ id: 'theme-classic-light', source: 'static' }));
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });

    expect(docEditorConfigs[0].editorConfig.customization.uiTheme).toBe('theme-system');
    expect(window.localStorage.getItem('ui-theme-id')).toBe('theme-system');
    expect(window.localStorage.getItem('ui-theme')).toBeNull();
    expect(Object.keys(themeControllers[0].map())).toEqual(['theme-system', 'theme-white', 'theme-night']);

    await instance.destroy();
  });

  it('maps supported interface theme choices to modern OnlyOffice theme ids', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      interfaceTheme: 'dark',
    });

    expect(docEditorConfigs[0].editorConfig.customization.uiTheme).toBe('theme-night');
    expect(window.localStorage.getItem('ui-theme-id')).toBe('theme-night');

    await instance.destroy();
  });

  it('migrates legacy interface theme ids to modern light or dark themes', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      interfaceTheme: 'theme-contrast-dark' as unknown as 'dark',
    });

    expect(docEditorConfigs[0].editorConfig.customization.uiTheme).toBe('theme-night');
    expect(window.localStorage.getItem('ui-theme-id')).toBe('theme-night');

    await instance.destroy();
  });

  it('updates the nested OnlyOffice theme at runtime', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });

    instance.setInterfaceTheme('light');
    expect(window.localStorage.getItem('ui-theme-id')).toBe('theme-white');
    expect(themeSetMocks[0]).toHaveBeenLastCalledWith('theme-white', 'host');

    instance.setInterfaceTheme('dark');
    expect(window.localStorage.getItem('ui-theme-id')).toBe('theme-night');
    expect(themeSetMocks[0]).toHaveBeenLastCalledWith('theme-night', 'host');

    await instance.destroy();
  });

  it('passes direct image URLs through the OnlyOffice parent image resolver', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['pptx'], 'slides.pptx'),
      fileName: 'slides.pptx',
      mode: 'preview',
    });
    await flush();

    const dataUrl = 'data:image/png;base64,iVBORw0KGgo=';
    const app = window as typeof window & {
      APP?: {
        getImageURL?: (name: string, callback: (url: string) => void) => void;
      };
    };
    const resolved = await new Promise<string>((resolve) => {
      app.APP?.getImageURL?.(dataUrl, resolve);
    });

    expect(resolved).toBe(dataUrl);

    await instance.destroy();
  });

  it('hides editor iframes before blank teardown to avoid white-frame flashes', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
    });
    await flush();

    const frame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    expect(frame).toBeTruthy();
    const sourceDescriptor = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'src');
    expect(sourceDescriptor?.get).toEqual(expect.any(Function));
    expect(sourceDescriptor?.set).toEqual(expect.any(Function));
    const assignments: Array<{
      src: string;
      visibility: string;
      opacity: string;
      pointerEvents: string;
      ariaHidden: string | null;
    }> = [];
    Object.defineProperty(frame!, 'src', {
      configurable: true,
      get() {
        return sourceDescriptor!.get!.call(this);
      },
      set(value: string) {
        assignments.push({
          src: value,
          visibility: frame!.style.visibility,
          opacity: frame!.style.opacity,
          pointerEvents: frame!.style.pointerEvents,
          ariaHidden: frame!.getAttribute('aria-hidden'),
        });
        sourceDescriptor!.set!.call(this, value);
        frame!.dispatchEvent(new Event('load'));
      },
    });

    await instance.destroy();

    expect(assignments).toContainEqual({
      src: 'about:blank',
      visibility: 'hidden',
      opacity: '0',
      pointerEvents: 'none',
      ariaHidden: 'true',
    });
  });

  it('opens preview mode as the desktop common viewer while allowing edit-rights requests', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      readonly: false,
      lang: 'zh',
    });
    await flush();

    expect(docEditorConfigs[0].type).toBe('desktop');
    expect(docEditorConfigs[0].editorConfig.mode).toBe('view');
    expect(docEditorConfigs[0].editorConfig.embedded).toBeUndefined();
    expect(docEditorConfigs[0].document.permissions.edit).toBe(true);
    expect(docEditorConfigs[0].document.permissions.download).toBe(true);
    expect(docEditorConfigs[0].events.onRequestEditRights).toEqual(expect.any(Function));
    expect(instance.getState()).toMatchObject({ mode: 'preview', readonly: true });

    await instance.destroy();
  });

  it('keeps readonly preview files in the desktop common viewer without edit rights', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      readonly: true,
    });
    await flush();

    expect(docEditorConfigs[0].type).toBe('desktop');
    expect(docEditorConfigs[0].editorConfig.mode).toBe('view');
    expect(docEditorConfigs[0].document.permissions.edit).toBe(false);
    expect(docEditorConfigs[0].document.permissions.download).toBe(true);
    expect(instance.getState()).toMatchObject({ mode: 'preview', readonly: true });

    await instance.destroy();
  });

  it('reopens a preview document in edit mode when OnlyOffice requests edit rights', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      readonly: false,
      lang: 'zh',
    });
    await flush();

    docEditorConfigs[0].events.onRequestEditRights?.();
    await waitForCondition(() => docEditorConfigs.length === 2);
    await flush();

    expect(docEditorInstances[1]).not.toBe(docEditorInstances[0]);
    expect(container.querySelectorAll('iframe[name="frameEditor"]')).toHaveLength(1);
    expect(docEditorConfigs[1].type).toBe('desktop');
    expect(docEditorConfigs[1].editorConfig.mode).toBe('edit');
    expect(docEditorConfigs[1].editorConfig.embedded).toBeUndefined();
    expect(docEditorConfigs[1].document.permissions.edit).toBe(true);
    expect(docEditorConfigs[1].document.permissions.download).toBe(true);
    expect(instance.getState()).toMatchObject({ mode: 'edit', readonly: false, status: 'ready' });

    await instance.destroy();
  });

  it('tracks native common-viewer edit mode changes without requiring a remount', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onStateChange = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      readonly: false,
      onStateChange,
    });
    await flush();

    const frame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameWindow = frame?.contentWindow;
    const frameDocument = frameWindow?.document;
    expect(frameWindow).toBeTruthy();
    expect(frameDocument).toBeTruthy();

    const slot = frameDocument!.createElement('div');
    slot.id = 'slot-btn-edit-mode';
    const button = frameDocument!.createElement('button');
    button.className = 'btn dropdown-toggle';
    const caption = frameDocument!.createElement('span');
    caption.className = 'caption';
    caption.textContent = '查看';
    button.appendChild(caption);
    slot.appendChild(button);
    frameDocument!.body.appendChild(slot);
    const homeTab = frameDocument!.createElement('button');
    homeTab.id = 'home';
    homeTab.textContent = '开始';
    frameDocument!.body.appendChild(homeTab);
    await waitForMessage();

    expect(instance.getState()).toMatchObject({ mode: 'preview', readonly: true });

    caption.textContent = '编辑';
    button.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    await waitForCondition(() => instance.getState().mode === 'edit');

    expect(docEditorConfigs).toHaveLength(1);
    expect(instance.getState()).toMatchObject({ mode: 'edit', readonly: false, status: 'ready' });
    expect(onStateChange).toHaveBeenLastCalledWith(
      expect.objectContaining({ mode: 'edit', readonly: false }),
      instance,
    );

    await instance.destroy();
  });

  it('shows an integrated OnlyOffice preview button for returning a preview-derived edit session', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      readonly: false,
      lang: 'zh',
    });
    await flush();

    const frame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameWindow = frame?.contentWindow;
    const frameDocument = frameWindow?.document;
    expect(frameWindow).toBeTruthy();
    expect(frameDocument).toBeTruthy();

    const slot = frameDocument!.createElement('div');
    slot.id = 'slot-btn-edit-mode';
    const button = frameDocument!.createElement('button');
    button.className = 'btn dropdown-toggle';
    const caption = frameDocument!.createElement('span');
    caption.className = 'caption';
    caption.textContent = '编辑';
    button.appendChild(caption);
    slot.appendChild(button);
    frameDocument!.body.appendChild(slot);
    const searchSlot = frameDocument!.createElement('span');
    searchSlot.id = 'slot-btn-search';
    searchSlot.className = 'btn-slot';
    frameDocument!.body.appendChild(searchSlot);
    const avatarSlot = frameDocument!.createElement('span');
    avatarSlot.id = 'slot-btn-user';
    avatarSlot.className = 'btn-slot';
    avatarSlot.textContent = 'LU';
    frameDocument!.body.appendChild(avatarSlot);
    const homeTab = frameDocument!.createElement('button');
    homeTab.id = 'home';
    homeTab.textContent = '开始';
    frameDocument!.body.appendChild(homeTab);

    await waitForCondition(() => instance.getState().mode === 'edit');
    await waitForCondition(() => !!frameDocument!.querySelector('#onlyoffice-browser-return-preview-mode button'));
    expect(docEditorConfigs).toHaveLength(1);

    const returnPreviewButton = frameDocument!.querySelector<HTMLButtonElement>(
      '#onlyoffice-browser-return-preview-mode button',
    );
    expect(returnPreviewButton?.textContent?.trim()).toBe('');
    expect(
      returnPreviewButton?.querySelector('svg[data-iconify-icon="qlementine-icons:preview-16"]'),
    ).toBeTruthy();
    expect(caption.textContent).toBe('编辑');
    expect(slot.dataset.officeBrowserModeAction).toBeUndefined();
    expect(frameDocument!.getElementById('onlyoffice-browser-return-preview-mode')?.nextSibling).toBe(avatarSlot);

    returnPreviewButton!.dispatchEvent(new MouseEvent('click', { bubbles: true }));

    await waitForCondition(() => docEditorConfigs.length === 2);
    await flush();

    expect(docEditorConfigs[1].type).toBe('desktop');
    expect(docEditorConfigs[1].editorConfig.mode).toBe('view');
    expect(instance.getState()).toMatchObject({ mode: 'preview', readonly: true, status: 'ready' });

    await instance.destroy();
  });

  it('reopens a preview-derived editor back in view mode when readonly is restored', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
      readonly: false,
    });
    await flush();

    docEditorConfigs[0].events.onRequestEditRights?.();
    await waitForCondition(() => docEditorConfigs.length === 2);
    await flush();

    instance.setReadonly(true);
    await waitForCondition(() => docEditorConfigs.length === 3);
    await flush();

    expect(docEditorConfigs[2].type).toBe('desktop');
    expect(docEditorConfigs[2].editorConfig.mode).toBe('view');
    expect(docEditorConfigs[2].editorConfig.embedded).toBeUndefined();
    expect(docEditorConfigs[2].document.permissions.edit).toBe(true);
    expect(instance.getState()).toMatchObject({ mode: 'preview', readonly: true, status: 'ready' });

    await instance.destroy();
  });

  it('passes converted media URLs into the common viewer documentOpen map', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const media = {
      'media/image1.png': 'blob:image-1',
    };
    mocks.convertDocument.mockResolvedValueOnce({
      fileName: 'slides.pptx',
      bin: new Uint8Array([1, 2, 3]),
      media,
    });

    const instance = await createOfficeEditor(container, {
      file: new File(['pptx'], 'slides.pptx'),
      fileName: 'slides.pptx',
      mode: 'preview',
    });
    await flush();

    const server = docEditorInstances[0].connectMockServer.mock.calls[0][0] as {
      getDocumentOpenData?: (documentUrl: string) => Record<string, string>;
    };
    expect(docEditorConfigs[0].type).toBe('desktop');
    expect(server.getDocumentOpenData?.('blob:document-bin')).toEqual({
      'Editor.bin': 'blob:document-bin',
      ...media,
    });

    await instance.destroy();
  });

  it('uses the OnlyOffice built-in warning dialog for save-to-new-format confirmation', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'legacy.doc'),
      fileName: 'legacy.doc',
      mode: 'edit',
    });
    const editorFrame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const warning = vi.fn((options: { callback?: (result: string, dontShow?: boolean) => void }) => {
      options.callback?.('ok', false);
    });
    Object.assign(editorFrame?.contentWindow as any, {
      Common: {
        UI: {
          warning,
          Window: {
            prototype: {
              textWarning: 'Warning',
            },
          },
        },
      },
      DE: {
        Controllers: {
          LeftMenu: {
            prototype: {
              notcriticalErrorTitle: 'Warning',
              txtCompatible: 'The document will be saved to the new format.',
            },
          },
        },
      },
    });

    await expect(instance.confirmSaveToNewFormat()).resolves.toBe(true);
    expect(warning).toHaveBeenCalledWith(expect.objectContaining({
      closable: false,
      width: 600,
      title: 'Warning',
      msg: 'The document will be saved to the new format.',
      buttons: ['ok', 'cancel'],
      dontshow: true,
      callback: expect.any(Function),
    }));

    await instance.destroy();
  });

  it('resolves false when the OnlyOffice save-to-new-format dialog is cancelled', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'legacy.xls'),
      fileName: 'legacy.xls',
      mode: 'edit',
    });
    const editorFrame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    Object.assign(editorFrame?.contentWindow as any, {
      Common: {
        UI: {
          warning: vi.fn((options: { callback?: (result: string, dontShow?: boolean) => void }) => {
            options.callback?.('cancel', false);
          }),
        },
      },
    });

    await expect(instance.confirmSaveToNewFormat()).resolves.toBe(false);

    await instance.destroy();
  });

  it('defaults runtime containers and the nested OnlyOffice frame to fill available space', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });
    const placeholder = container.querySelector<HTMLElement>('.office-editor-frame');
    const frame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');

    expect(container.classList.contains('office-editor-host')).toBe(true);
    expect(container.style.width).toBe('100%');
    expect(container.style.height).toBe('100%');
    expect(container.style.minWidth).toBe('0px');
    expect(container.style.minHeight).toBe('0px');
    expect(placeholder?.style.width).toBe('100%');
    expect(placeholder?.style.height).toBe('100%');
    expect(placeholder?.style.minWidth).toBe('0px');
    expect(placeholder?.style.minHeight).toBe('0px');
    expect(frame?.style.display).toBe('block');
    expect(frame?.style.width).toBe('100%');
    expect(frame?.style.height).toBe('100%');
    expect(frame?.style.minWidth).toBe('0px');
    expect(frame?.style.minHeight).toBe('0px');

    await instance.destroy();
  });

  it('can opt into spellcheck while keeping the in-editor toggle disabled', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      spellcheck: true,
    });

    expect(docEditorConfigs[0].editorConfig.customization).toMatchObject({
      compactToolbar: true,
      spellcheck: true,
      features: {
        featuresTips: false,
        spellcheck: {
          change: false,
        },
      },
    });

    await instance.destroy();
  });

  it('sets a stable local user identity for editor notifications', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });

    expect(docEditorConfigs[0].editorConfig.user).toEqual({
      id: 'local-browser-user',
      name: 'Local Browser User',
    });
    expect(docEditorConfigs[0].editorConfig.customization.anonymous).toEqual({
      request: false,
      label: 'Local Browser User',
    });

    await instance.destroy();
  });

  it('saves manually from native editor bin without triggering downloadAs', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onSave = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      onSave,
    });
    await flush();

    const savedFile = await instance.save('XLSX');

    expect(docEditorInstances[0].asc_nativeGetFile3).toHaveBeenCalledTimes(1);
    expect(docEditorInstances[0].downloadAs).not.toHaveBeenCalled();
    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.xlsx', 'XLSX', {});
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onSave.mock.calls[0][0]).toMatchObject({
      name: 'alpha.xlsx',
      size: zipFixture().byteLength,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    expect(savedFile).toMatchObject({
      name: 'alpha.xlsx',
      size: zipFixture().byteLength,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    await instance.destroy();
  });

  it('passes loaded media to native save export', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const media = {
      'media/image1.png': 'blob:image-1',
    };
    mocks.convertDocument.mockResolvedValueOnce({
      fileName: 'with-image.docx',
      bin: new Uint8Array([1, 2, 3]),
      media,
    });

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'with-image.docx'),
      fileName: 'with-image.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    await instance.save('DOCX');

    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'with-image.docx', 'DOCX', media);

    await instance.destroy();
  });

  it('downloads as a selected non-PDF format through native bin conversion when OnlyOffice has no URL', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    mocks.convertBinToDocument.mockResolvedValueOnce({
      fileName: 'alpha.odt',
      data: zipFixture(),
    });

    docEditorConfigs[0].events.onDownloadAs?.({ data: { fileType: 67 } });
    await waitForMessage();

    expect(docEditorInstances[0].asc_nativeGetFile3).toHaveBeenCalledTimes(1);
    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.docx', 'ODT', {});
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it('passes Save Copy As files to the host save-as callback without downloading', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onDownload = vi.fn();
    const onSaveAs = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
      onSaveAs,
      onDownload,
    });
    await flush();

    mocks.convertBinToDocument.mockResolvedValueOnce({
      fileName: 'alpha.odt',
      data: zipFixture(),
    });

    docEditorConfigs[0].events.onRequestSaveAs?.({ data: { fileType: 67, title: 'alpha.odt' } });
    await waitForMessage();

    expect(onSaveAs).toHaveBeenCalledTimes(1);
    expect(onSaveAs.mock.calls[0][0]).toMatchObject({ name: 'alpha.odt' });
    expect(onDownload).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('handles built-in Save As requests from the File Download As panel', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    mocks.convertBinToDocument.mockResolvedValueOnce({
      fileName: 'alpha.rtf',
      data: asciiBytes('{\\rtf1\\ansi alpha}'),
    });

    docEditorConfigs[0].events.onRequestSaveAs?.({ data: { fileType: 68, title: 'alpha.rtf' } });
    await waitForMessage();

    expect(docEditorInstances[0].asc_nativeGetFile3).toHaveBeenCalledTimes(1);
    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.docx', 'RTF', {});
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it('uses the native PDF renderer for built-in Download As PDF', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    const pdfBytes = validOnePagePdfFixture();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }).native?.Save_End?.(
        '',
        printStream.byteLength,
      );
      return printStream;
    });
    mocks.convertPrintDataToPdf.mockResolvedValueOnce({
      fileName: 'alpha.pdf',
      data: pdfBytes,
    });

    docEditorConfigs[0].events.onDownloadAs?.({ data: { fileType: 513 } });
    await waitForMessage();

    expect(docEditorInstances[0].asc_nativeGetPDF).toHaveBeenCalledWith(expect.objectContaining({ isPrint: true }));
    expect(mocks.convertPrintDataToPdf).toHaveBeenCalledWith(printStream, 'alpha.docx', {});
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it('uses the native PDF renderer plus x2t PDF rasterization for built-in Download As image', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    const pdfBytes = validOnePagePdfFixture();
    const jpgBytes = jpegFixture();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }).native?.Save_End?.(
        '',
        printStream.byteLength,
      );
      return printStream;
    });
    mocks.convertPrintDataToPdf.mockResolvedValueOnce({
      fileName: 'alpha.pdf',
      data: pdfBytes,
    });
    mocks.convertPdfToImage.mockResolvedValueOnce({
      fileName: 'alpha.jpg',
      data: jpgBytes,
    });

    docEditorConfigs[0].events.onDownloadAs?.({ data: { fileType: 1025 } });
    await waitForMessage();

    expect(docEditorInstances[0].asc_nativeGetPDF).toHaveBeenCalledWith(expect.objectContaining({ isPrint: true }));
    expect(mocks.convertPrintDataToPdf).toHaveBeenCalledWith(printStream, 'alpha.xlsx', {});
    expect(mocks.convertPdfToImage).toHaveBeenCalledWith(pdfBytes, 'alpha.xlsx', 'jpg', { allPages: false });
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it('does not fall back to the doctrenderer-backed save-bin route for PDF downloads', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onError = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
      onError,
    });
    await flush();

    docEditorInstances[0].asc_nativeGetPDF.mockReturnValueOnce(null);

    docEditorConfigs[0].events.onDownloadAs?.({ data: { fileType: 513 } });
    await waitForMessage();

    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(onError).toHaveBeenCalledWith(
      expect.objectContaining({ message: expect.stringContaining('native print export is not available') }),
      instance,
    );
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('does not fall back to save-bin PDF conversion when spreadsheet PDF settings cannot be applied', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onError = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      saveBehavior: 'download',
      onError,
    });
    await flush();

    docEditorInstances[0].asc_nativeGetPDF.mockReturnValueOnce(null);
    docEditorInstances[0].nativeApi.asc_DownloadAs({ fileType: 513, advancedOptions: { range: 'active-sheets' } });
    await waitForMessage();

    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(onError).toHaveBeenCalledWith(
      expect.objectContaining({ message: expect.stringContaining('settings could not be applied') }),
      instance,
    );
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('passes spreadsheet PDF sheet and page settings through the native CAdjustPrint bridge', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const pdfBytes = validOnePagePdfFixture();
    const adjustPrint = {
      asc_getPrintType: vi.fn(() => 2),
      asc_getActiveSheetsArray: vi.fn(() => [0, 2]),
      asc_getStartPageIndex: vi.fn(() => 1),
      asc_getEndPageIndex: vi.fn(() => 3),
      asc_getPageOptionsMap: vi.fn(() => ({ 0: { id: 'sheet-0-options' }, 2: { id: 'sheet-2-options' } })),
    };
    let capturedPrintOptions: unknown = 'not-called';
    let capturedDesktopPrintOptions: unknown;

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce((options?: unknown) => {
      capturedPrintOptions = options;
      capturedDesktopPrintOptions = (window as typeof window & { AscDesktopEditor_PrintOptions?: unknown })
        .AscDesktopEditor_PrintOptions;
      (window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }).native?.Save_End?.(
        '',
        pdfBytes.byteLength,
      );
      return pdfBytes;
    });

    docEditorInstances[0].nativeApi.asc_DownloadAs({ fileType: 513, advancedOptions: adjustPrint });
    await waitForMessage();

    expect(capturedPrintOptions).toBeUndefined();
    expect(capturedDesktopPrintOptions).toEqual({ advancedOptions: adjustPrint });
    expect((window as typeof window & { AscDesktopEditor_PrintOptions?: unknown }).AscDesktopEditor_PrintOptions).toBeUndefined();
    expect(docEditorInstances[0].asc_nativeGetPDF).toHaveBeenCalledTimes(1);
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

	  it('intercepts final native asc_DownloadAs calls after upstream Download As dialogs are confirmed', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    mocks.convertBinToDocument.mockResolvedValueOnce({
      fileName: 'alpha.odt',
      data: zipFixture(),
    });

    docEditorInstances[0].nativeApi.asc_DownloadAs({ fileType: 67 });
    await waitForMessage();

	    expect(docEditorInstances[0].nativeApi.__onlyOfficeBrowserDownloadAsPatched).toBe(true);
	    expect(docEditorInstances[0].nativeApi.__onlyOfficeBrowserAscDownloadAsOriginal).toBeTruthy();
	    expect(docEditorInstances[0].nativeDownloadAs).not.toHaveBeenCalled();
	    expect(docEditorInstances[0].nativeDownloadAsHandler).not.toHaveBeenCalled();
    expect(docEditorInstances[0].asc_nativeGetFile3).toHaveBeenCalledTimes(1);
    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.docx', 'ODT', {});
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

	    await instance.destroy();
	  });

	  it('still intercepts prepared native _downloadAs calls that bypass asc_DownloadAs', async () => {
	    const container = document.createElement('div');
	    document.body.appendChild(container);

	    const instance = await createOfficeEditor(container, {
	      file: new File(['hello'], 'alpha.docx'),
	      fileName: 'alpha.docx',
	      mode: 'edit',
	      saveBehavior: 'download',
	    });
	    await flush();

	    mocks.convertBinToDocument.mockResolvedValueOnce({
	      fileName: 'alpha.ott',
	      data: zipFixture(),
	    });

	    docEditorInstances[0].nativeApi._downloadAs(1, { fileType: 79 }, { outputformat: 79, title: 'alpha.ott' });
	    await waitForMessage();

	    expect(docEditorInstances[0].nativeDownloadAsHandler).not.toHaveBeenCalled();
	    expect(docEditorInstances[0].asc_nativeGetFile3).toHaveBeenCalledTimes(1);
	    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.docx', 'OTT', {});
	    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

	    await instance.destroy();
	  });

	  it('routes native Save Copy As _downloadAs calls to onSaveAs instead of browser download', async () => {
	    const container = document.createElement('div');
	    document.body.appendChild(container);
	    const onSaveAs = vi.fn();
	    const onDownload = vi.fn();

	    const instance = await createOfficeEditor(container, {
	      file: new File(['hello'], 'alpha.docx'),
	      fileName: 'alpha.docx',
	      mode: 'edit',
	      saveBehavior: 'callback',
	      onSaveAs,
	      onDownload,
	    });
	    await flush();

	    mocks.convertBinToDocument.mockResolvedValueOnce({
	      fileName: 'alpha.odt',
	      data: zipFixture(),
	    });

	    docEditorInstances[0].nativeApi._downloadAs(
	      1,
	      { fileType: 67, isSaveAs: true },
	      { outputformat: 67, title: 'alpha.odt', isSaveAs: true },
	    );
	    await waitForMessage();

	    expect(docEditorInstances[0].nativeDownloadAsHandler).not.toHaveBeenCalled();
	    expect(docEditorInstances[0].asc_nativeGetFile3).toHaveBeenCalledTimes(1);
	    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.docx', 'ODT', {});
	    expect(onSaveAs).toHaveBeenCalledTimes(1);
	    expect(onSaveAs.mock.calls[0][0]).toMatchObject({ name: 'alpha.odt' });
	    expect(onDownload).not.toHaveBeenCalled();
	    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

	    await instance.destroy();
	  });

  it('lets native print _downloadAs requests continue into OnlyOffice APP.printPdf flow', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    docEditorInstances[0].nativeApi._downloadAs(
      7,
      { fileType: 513, isDownloadEvent: true },
      { outputformat: 513, title: 'alpha.pdf', inline: 1 },
      { data: new Uint8Array([1, 2, 3]) },
      'asc_onPrintUrl',
    );
    await waitForMessage();

    expect(docEditorInstances[0].nativeDownloadAsHandler).toHaveBeenCalledTimes(1);
    expect(docEditorInstances[0].nativeDownloadAsHandler).toHaveBeenCalledWith(
      7,
      expect.objectContaining({ fileType: 513, isDownloadEvent: true }),
      expect.objectContaining({ outputformat: 513, inline: 1 }),
      expect.objectContaining({ data: expect.any(Uint8Array) }),
      'asc_onPrintUrl',
    );
    expect(docEditorInstances[0].asc_nativeGetFile3).not.toHaveBeenCalled();
    expect(docEditorInstances[0].asc_nativeGetPDF).not.toHaveBeenCalled();
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(mocks.convertPrintDataToPdf).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('lets asc_DownloadAs print markers reach upstream instead of downloading a PDF', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    docEditorInstances[0].nativeApi.asc_DownloadAs({ fileType: 513, isPrint: true });
    await waitForMessage();

    expect(docEditorInstances[0].nativeDownloadAsHandler).toHaveBeenCalledTimes(1);
    expect(docEditorInstances[0].nativeDownloadAsHandler).toHaveBeenCalledWith(
      1,
      expect.objectContaining({ fileType: 513, isPrint: true }),
      expect.objectContaining({ outputformat: 513 }),
      undefined,
      undefined,
    );
    expect(docEditorInstances[0].asc_nativeGetFile3).not.toHaveBeenCalled();
    expect(docEditorInstances[0].asc_nativeGetPDF).not.toHaveBeenCalled();
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(mocks.convertPrintDataToPdf).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('lets built-in Download As panel tile clicks reach the upstream OnlyOffice UI', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    const editorFrame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const formatButton = editorFrame!.contentDocument!.createElement('div');
    formatButton.className = 'btn-doc-format';
    formatButton.setAttribute('format', '67');
    const icon = editorFrame!.contentDocument!.createElement('div');
    icon.className = 'svg-format-odt';
    formatButton.appendChild(icon);
    editorFrame!.contentDocument!.body.appendChild(formatButton);

    const clickEvent = new MouseEvent('click', { bubbles: true, cancelable: true });
    const notCancelled = icon.dispatchEvent(clickEvent);
    await flush();

    expect(notCancelled).toBe(true);
    expect(clickEvent.defaultPrevented).toBe(false);
    expect(docEditorInstances[0].asc_nativeGetFile3).not.toHaveBeenCalled();
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    await instance.destroy();
  });

  it('opens spreadsheet PDF Download As and Save Copy As in the native file-menu side panel', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      saveBehavior: 'download',
      lang: 'zh',
    });
    await flush();

    const editorFrame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameWindow = editorFrame!.contentWindow!;
    const harness = installSpreadsheetPrintPreviewHarness(frameWindow);
    docEditorConfigs[0].events.onAppReady();
    await flush();

    const trigger = (frameWindow as typeof frameWindow & {
      Common: { NotificationCenter: { trigger: (...args: unknown[]) => unknown } };
    }).Common.NotificationCenter.trigger;

    trigger('download:settings', {}, 513, false);
    expect(harness.originalTrigger).not.toHaveBeenCalledWith(
      'download:settings',
      expect.anything(),
      expect.anything(),
    );
    expect(harness.showMenu).not.toHaveBeenCalledWith('file:printpreview');
    const downloadPanel = frameWindow.document.querySelector(
      '[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]',
    );
    expect(downloadPanel).toBe(harness.printPanel);
    expect(harness.printPanel.querySelector('button[aria-label="关闭"]')).toBeNull();
    expect(harness.panelContext.contains(harness.printPanel)).toBe(true);
    expect(harness.printPanel.style.display).toBe('block');
    expect(harness.panelSaveAs.style.display).toBe('none');
    expect(harness.printSettingsShow).toHaveBeenCalledTimes(1);
    expect(harness.printMainHeader.textContent).toBe('下载');
    expect(harness.printPanel.style.position).toBe('');
    expect(harness.printPanel.style.top).toBe('');
    expect(harness.printPanel.style.bottom).toBe('');
    expect(harness.printPanel.querySelector<HTMLElement>('#id-print-settings')?.style.position).toBe('');
    expect(harness.printPanel.querySelector<HTMLElement>('#print-preview-box')?.style.left).toBe('');
    expect(harness.showFilePanel).toHaveBeenLastCalledWith('saveas');
    expect(harness.printSettings.btnsPrint[0].setCaption).toHaveBeenCalledWith('下载');
    expect(harness.printSettings.btnsSave[0].setCaption).toHaveBeenCalledWith('保存设置');
    expect(harness.printController.updatePrintRenderContainerSize).toHaveBeenCalledWith(true);
    expect(harness.printSettings.printScroller.update).toHaveBeenCalled();

    harness.showFilePanel.mockClear();
    trigger('download:cancel');
    expect(
      frameWindow.document.querySelector('[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]'),
    ).toBeNull();
    expect(harness.panelContext.contains(harness.printPanel)).toBe(true);
    expect(harness.printPanel.style.display).toBe('none');
    expect(harness.panelSaveAs.style.display).toBe('block');
    expect(harness.printMainHeader.textContent).toBe('打印');
    expect(harness.printPanel.style.position).toBe('');
    expect(harness.printPanel.querySelector('button[aria-label="关闭"]')).toBeNull();
    expect(harness.printPanel.querySelector<HTMLElement>('#print-preview-box')?.style.left).toBe('');
    expect(harness.printController.onHidePrintMenu).toHaveBeenCalledTimes(1);
    expect(harness.showFilePanel).not.toHaveBeenCalled();

    harness.showMenu.mockClear();
    harness.showFilePanel.mockClear();
    harness.printController.onHidePrintMenu.mockClear();
    harness.printSettings.btnsPrint[0].setCaption.mockClear();
    harness.printSettings.btnsSave[0].setCaption.mockClear();
    harness.printSettingsShow.mockClear();

    trigger('download:settings', {}, 513, false);
    expect(harness.showMenu).not.toHaveBeenCalledWith('file:printpreview');
    expect(harness.showFilePanel).toHaveBeenLastCalledWith('saveas');
    expect(harness.printSettingsShow).toHaveBeenCalledTimes(1);
    expect(
      frameWindow.document.querySelector('[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]'),
    ).not.toBeNull();

    harness.printController.querySavePrintSettings('print', false);

    expect(harness.printController.savePageOptions).toHaveBeenCalledWith(harness.printSettings);
    expect(harness.printSettings.applySettings).toHaveBeenCalledTimes(1);
    expect(harness.adjustPrintParams.asc_setPrintType).toHaveBeenCalledWith(2);
    expect(harness.adjustPrintParams.asc_setPageOptionsMap).toHaveBeenCalledWith([]);
    expect(harness.adjustPrintParams.asc_setIgnorePrintArea).toHaveBeenCalledWith(true);
    expect(harness.adjustPrintParams.asc_setActiveSheetsArray).toHaveBeenCalledWith([0]);
    expect(harness.adjustPrintParams.asc_setStartPageIndex).toHaveBeenCalledWith(0);
    expect(harness.adjustPrintParams.asc_setEndPageIndex).toHaveBeenCalledWith(3);
    expect(harness.adjustPrintParams.asc_setNativeOptions).toHaveBeenCalledWith(
      expect.objectContaining({
        colorMode: true,
        copies: 1,
        paperOrientation: 'landscape',
        paperSize: expect.objectContaining({ h: 420, preset: 'A3', w: 297 }),
        printer: 'printer-1',
        sides: 'one',
        usesystemdialog: false,
      }),
    );
    expect(harness.ascDownloadAs).toHaveBeenCalledTimes(1);
    const downloadOptions = harness.ascDownloadAs.mock.calls[0][0] as {
      advancedOptions?: unknown;
      fileType?: unknown;
      isSaveAs?: unknown;
      wopiSaveAsPath?: unknown;
    };
    expect(downloadOptions.fileType).toBe(513);
    expect(downloadOptions.isSaveAs).toBe(false);
    expect(downloadOptions.wopiSaveAsPath).toBeUndefined();
    expect(downloadOptions.advancedOptions).toBe(harness.adjustPrintParams);
    expect(
      frameWindow.document.querySelector('[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]'),
    ).toBeNull();
    expect(harness.panelContext.contains(harness.printPanel)).toBe(true);
    expect(harness.printPanel.style.display).toBe('none');
    expect(harness.panelSaveAs.style.display).toBe('block');
    expect(harness.printPanel.style.position).toBe('');
    expect(harness.printController.onHidePrintMenu).toHaveBeenCalledTimes(1);
    expect(harness.showFilePanel).toHaveBeenLastCalledWith('saveas');
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    harness.showMenu.mockClear();
    harness.showFilePanel.mockClear();
    harness.ascDownloadAs.mockClear();
    harness.printSettings.btnsPrint[0].setCaption.mockClear();
    harness.printSettings.btnsSave[0].setCaption.mockClear();
    harness.printSettingsShow.mockClear();
    harness.printController.onHidePrintMenu.mockClear();

    trigger('download:settings', {}, 521, true, '/alpha-copy.pdf');
    expect(harness.showMenu).not.toHaveBeenCalledWith('file:printpreview');
    expect(harness.showFilePanel).toHaveBeenLastCalledWith('save-copy');
    expect(harness.printSettingsShow).toHaveBeenCalledTimes(1);
    expect(harness.printMainHeader.textContent).toBe('另存副本');
    const saveCopyPanel = frameWindow.document.querySelector(
      '[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]',
    );
    expect(saveCopyPanel).toBe(harness.printPanel);
    expect(harness.printPanel.querySelector('button[aria-label="关闭"]')).toBeNull();
    expect(harness.panelContext.contains(harness.printPanel)).toBe(true);
    expect(harness.printPanel.style.display).toBe('block');
    expect(harness.panelSaveCopy.style.display).toBe('none');
    expect(harness.printPanel.style.position).toBe('');
    expect(harness.printPanel.style.top).toBe('');
    expect(harness.printSettings.btnsPrint[0].setCaption).toHaveBeenCalledWith('保存副本');
    expect(harness.printSettings.btnsSave[0].setCaption).toHaveBeenCalledWith('保存设置');

    harness.printSettingsShow.mockClear();
    harness.printSettings.show?.();
    expect(harness.printSettingsShow).toHaveBeenCalledTimes(1);
    expect(harness.printMainHeader.textContent).toBe('打印');
    expect(
      frameWindow.document.querySelector('[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]'),
    ).toBeNull();

    harness.printSettingsShow.mockClear();
    trigger('download:settings', {}, 521, true, '/alpha-copy.pdf');
    expect(harness.printSettingsShow).toHaveBeenCalledTimes(1);
    expect(harness.printMainHeader.textContent).toBe('另存副本');

    harness.printController.querySavePrintSettings('print-pdf', false);

    expect(harness.ascDownloadAs).toHaveBeenCalledTimes(1);
    const saveCopyOptions = harness.ascDownloadAs.mock.calls[0][0] as {
      advancedOptions?: unknown;
      fileType?: unknown;
      isSaveAs?: unknown;
      wopiSaveAsPath?: unknown;
    };
    expect(saveCopyOptions.fileType).toBe(521);
    expect(saveCopyOptions.isSaveAs).toBe(true);
    expect(saveCopyOptions.wopiSaveAsPath).toBe('/alpha-copy.pdf');
    expect(saveCopyOptions.advancedOptions).toBe(harness.adjustPrintParams);
    expect(
      frameWindow.document.querySelector('[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]'),
    ).toBeNull();
    expect(harness.panelContext.contains(harness.printPanel)).toBe(true);
    expect(harness.printPanel.style.display).toBe('none');
    expect(harness.panelSaveCopy.style.display).toBe('block');
    expect(harness.printPanel.style.position).toBe('');
    expect(harness.printController.onHidePrintMenu).toHaveBeenCalledTimes(1);
    expect(harness.showFilePanel).toHaveBeenLastCalledWith('save-copy');
    expect(harness.originalTrigger).toHaveBeenCalledWith(
      'edit:complete',
      expect.objectContaining({ id: 'toolbar' }),
    );
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    harness.originalTrigger.mockClear();
    trigger('file:print', {});
    expect(
      frameWindow.document.querySelector('[data-onlyoffice-browser-spreadsheet-pdf-print-panel="true"]'),
    ).toBeNull();
    expect(harness.originalTrigger).toHaveBeenCalledWith('file:print', {});

    await instance.destroy();
  });

  it('extracts native asc_CDownloadOptions file types through accessors', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

	    mocks.convertHtmlToDocument.mockResolvedValueOnce({
	      fileName: 'alpha.html',
	      data: asciiBytes('<!doctype html><html><body>alpha</body></html>'),
	    });

	    docEditorInstances[0].nativeApi.asc_DownloadAs({ asc_getFileType: () => 70 });
	    await waitForMessage();

	    expect(docEditorInstances[0].asc_nativeCalculateFile).toHaveBeenCalledWith(expect.objectContaining({
	      asc_getFileType: expect.any(Function),
	    }));
	    expect(docEditorInstances[0].asc_nativeGetHtml).toHaveBeenCalledWith(expect.objectContaining({
	      asc_getFileType: expect.any(Function),
	    }));
	    const htmlCall = mocks.convertHtmlToDocument.mock.calls.at(-1);
	    expect(ArrayBuffer.isView(htmlCall?.[0])).toBe(true);
	    expect(htmlCall?.[1]).toBe('alpha.docx');
	    expect(htmlCall?.[2]).toBe('html');
	    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
	    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it('maps the current OnlyOffice MD file type constant for Download As', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

	    mocks.convertHtmlToDocument.mockResolvedValueOnce({
	      fileName: 'alpha.md',
	      data: asciiBytes('# alpha\n'),
	    });

	    docEditorInstances[0].nativeApi.asc_DownloadAs({ asc_getFileType: () => 92 });
	    await waitForMessage();

	    expect(docEditorInstances[0].asc_nativeCalculateFile).toHaveBeenCalledWith(expect.objectContaining({
	      asc_getFileType: expect.any(Function),
	    }));
	    expect(docEditorInstances[0].asc_nativeGetHtml).toHaveBeenCalledWith(expect.objectContaining({
	      asc_getFileType: expect.any(Function),
	    }));
	    const markdownCall = mocks.convertHtmlToDocument.mock.calls.at(-1);
	    expect(ArrayBuffer.isView(markdownCall?.[0])).toBe(true);
	    expect(markdownCall?.[1]).toBe('alpha.docx');
	    expect(markdownCall?.[2]).toBe('md');
	    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
	    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it('packages Markdown data URI images into a ZIP with separate assets', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const pngBase64 = window.btoa('\x89PNG\r\n\x1a\n');

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

	    mocks.convertHtmlToDocument.mockResolvedValueOnce({
	      fileName: 'alpha.md',
	      data: asciiBytes(`# alpha\n\n![diagram](data:image/png;base64,${pngBase64})\n`),
	    });

    docEditorConfigs[0].events.onDownloadAs?.({ data: { fileType: 92 } });
    await waitForMessage();

    const downloadedFile = (URL.createObjectURL as ReturnType<typeof vi.fn>).mock.calls.at(-1)?.[0] as File;
    expect(downloadedFile).toMatchObject({ name: 'alpha_docx_md.zip' });
    const bytes = new Uint8Array(await downloadedFile.arrayBuffer());
    const latin1 = Array.from(bytes, (byte) => String.fromCharCode(byte)).join('');
    expect(Array.from(bytes.slice(0, 4))).toEqual([0x50, 0x4b, 0x03, 0x04]);
    expect(latin1).toContain('alpha.md');
    expect(latin1).toContain('assets/image-001.png');
    expect(latin1).not.toContain('data:image/png;base64');

    await instance.destroy();
  });

  it('uses already packaged native OOXML bytes directly when the editor returns the requested package format', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onSave = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      onSave,
    });
    await flush();

    const zipBytes = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]);
    docEditorInstances[0].asc_nativeGetFile3.mockReturnValueOnce({ data: zipBytes });

    const savedFile = await instance.save('XLSX');

    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(savedFile).toMatchObject({
      name: 'alpha.xlsx',
      size: zipBytes.byteLength,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    expect(onSave).toHaveBeenCalledWith(savedFile, instance);

    await instance.destroy();
  });

  it('saves native built-in save events without an external save request', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onSave = vi.fn();
    const onDirtyChange = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      onSave,
      onDirtyChange,
    });
    await flush();

    docEditorConfigs[0].events.onDocumentStateChange?.({ data: true });
    expect(instance.getState().dirty).toBe(true);

    docEditorConfigs[0].events.onSave?.({
      data: {
        data: {
          data: new Uint8Array([4, 5, 6]),
        },
        option: {
          outputformat: 257,
        },
      },
    });
    await waitForMessage();

    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.xlsx', 'XLSX', {});
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onSave.mock.calls[0][0]).toMatchObject({
      name: 'alpha.xlsx',
      size: zipFixture().byteLength,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    expect(instance.getState().dirty).toBe(false);
    expect(onDirtyChange).toHaveBeenLastCalledWith(false, instance);

    await instance.destroy();
  });

  it('returns same-origin cached PDF URLs for OnlyOffice built-in print', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    document.title = 'Workbench';
    const pdfFixture = asciiBytes(`%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R >>
endobj
xref
0 4
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
trailer
<< /Size 4 /Root 1 0 R >>
startxref
160
%%EOF
`);
    mocks.convertPrintDataToPdf.mockResolvedValueOnce({
      fileName: 'alpha.pdf',
      data: pdfFixture,
    });

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });
    await flush();
    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (
        window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }
      ).native?.Save_End?.('', printStream.byteLength);
      return printStream;
    });

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void })
      .printPdf;
    expect(printPdf).toEqual(expect.any(Function));

    printPdf?.({ data: new Uint8Array([1, 2, 3]) }, callback);
    await vi.waitFor(() => expect(callback).toHaveBeenCalled());

    expect(mocks.convertPrintDataToPdf).toHaveBeenCalledWith(printStream, 'alpha.docx', {});
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    const result = callback.mock.calls[0][0] as { type: string; status: string; data: string; filetype: number };
    expect(result).toMatchObject({
      type: 'save',
      status: 'ok',
      filetype: 513,
    });
    expect(result.data).toContain('/__onlyoffice-browser-print__/');
    expect(new URL(result.data).pathname).toMatch(/\.pdf$/);
    expect(new URL(result.data).searchParams.get('filename')).toBe('alpha.pdf');
    expect(document.title).toBe('alpha');

    const cache = await window.caches.open('onlyoffice-browser-print-pdfs');
    const cachedPdf = await cache.match(result.data);
    expect(cachedPdf?.headers.get('accept-ranges')).toBe('bytes');
    expect(cachedPdf?.headers.get('content-disposition')).toBe(
      `inline; filename="alpha.pdf"; filename*=UTF-8''alpha.pdf`,
    );
    expect(Number(cachedPdf?.headers.get('content-length'))).toBeGreaterThan(pdfFixture.byteLength);
    expect(cachedPdf?.headers.get('content-type')).toBe('application/pdf');
    expect(cachedPdf?.headers.get('x-content-type-options')).toBe('nosniff');
    const cachedPdfBytes = new Uint8Array(await cachedPdf!.arrayBuffer());
    expect(binaryString(cachedPdfBytes)).toContain(`/Title <${pdfUtf16BeHex('alpha')}>`);
    window.dispatchEvent(new Event('afterprint'));
    expect(document.title).toBe('Workbench');

    await instance.destroy();
  });

  it('preserves loaded media and UTF-8 filenames for built-in print', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    document.title = 'Workbench';
    const media = {
      'media/image1.png': 'blob:image-1',
    };
    const pdfFixture = asciiBytes(`%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R >>
endobj
%%EOF
`);
    mocks.convertDocument.mockResolvedValueOnce({
      fileName: 'vLLM模型.docx',
      bin: new Uint8Array([1, 2, 3]),
      media,
    });
    mocks.convertPrintDataToPdf.mockResolvedValueOnce({
      fileName: 'vLLM模型.pdf',
      data: pdfFixture,
    });

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'vLLM模型.docx'),
      fileName: 'vLLM模型.docx',
      mode: 'edit',
    });
    await flush();
    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (
        window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }
      ).native?.Save_End?.('', printStream.byteLength);
      return printStream;
    });

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void })
      .printPdf;
    printPdf?.({ data: new Uint8Array([1, 2, 3]) }, callback);
    await vi.waitFor(() => expect(callback).toHaveBeenCalled());

    expect(mocks.convertPrintDataToPdf).toHaveBeenCalledWith(printStream, 'vLLM模型.docx', media);
    const result = callback.mock.calls[0][0] as { data: string };
    expect(result.data).toContain(encodeURIComponent('vLLM模型.pdf'));
    expect(new URL(result.data).searchParams.get('filename')).toBe('vLLM模型.pdf');
    expect(document.title).toBe('vLLM模型');

    const cache = await window.caches.open('onlyoffice-browser-print-pdfs');
    const cachedPdf = await cache.match(result.data);
    const contentDisposition = cachedPdf?.headers.get('content-disposition') || '';
    expect(contentDisposition).toContain('filename="vLLM_.pdf"');
    expect(contentDisposition).toContain(`filename*=UTF-8''${encodeURIComponent('vLLM模型.pdf')}`);
    expect(new Uint8Array(await cachedPdf!.arrayBuffer())).toEqual(pdfFixture);
    window.dispatchEvent(new Event('afterprint'));
    expect(document.title).toBe('Workbench');

    await instance.destroy();
  });

  it('keeps the host-provided UTF-8 file name even when conversion returns a sanitized name', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    document.title = 'Workbench';
    mocks.convertDocument.mockResolvedValueOnce({
      fileName: 'vLLM.docx',
      bin: new Uint8Array([1, 2, 3]),
      media: {},
    });

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'vLLM模型.docx'),
      fileName: 'vLLM模型.docx',
      mode: 'edit',
    });
    await flush();

    expect(instance.getState().fileName).toBe('vLLM模型.docx');

    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (
        window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }
      ).native?.Save_End?.('', printStream.byteLength);
      return printStream;
    });

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void })
      .printPdf;
    printPdf?.({ data: new Uint8Array([1, 2, 3]) }, callback);
    await vi.waitFor(() => expect(callback).toHaveBeenCalled());

    expect(mocks.convertPrintDataToPdf).toHaveBeenCalledWith(printStream, 'vLLM模型.docx', {});
    const result = callback.mock.calls[0][0] as { data: string };
    expect(result.data).toContain(encodeURIComponent('vLLM模型.pdf'));
    expect(document.title).toBe('vLLM模型');

    const cache = await window.caches.open('onlyoffice-browser-print-pdfs');
    const cachedPdf = await cache.match(result.data);
    const contentDisposition = cachedPdf?.headers.get('content-disposition') || '';
    expect(contentDisposition).toContain(`filename*=UTF-8''${encodeURIComponent('vLLM模型.pdf')}`);

    await instance.destroy();
  });

  it('adds a UTF-16 PDF title so all-CJK print names do not fall back to document.pdf', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    document.title = 'Workbench';
    const pdfFixture = asciiBytes(`%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R >>
endobj
xref
0 4
0000000000 65535 f 
0000000009 00000 n 
0000000058 00000 n 
0000000115 00000 n 
trailer
<< /Size 4 /Root 1 0 R >>
startxref
160
%%EOF
`);
    mocks.convertDocument.mockResolvedValueOnce({
      fileName: '文档.docx',
      bin: new Uint8Array([1, 2, 3]),
      media: {},
    });
    mocks.convertPrintDataToPdf.mockResolvedValueOnce({
      fileName: '文档.pdf',
      data: pdfFixture,
    });

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], '文档.docx'),
      fileName: '文档.docx',
      mode: 'edit',
    });
    await flush();
    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (
        window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }
      ).native?.Save_End?.('', printStream.byteLength);
      return printStream;
    });

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void })
      .printPdf;
    printPdf?.({ data: new Uint8Array([1, 2, 3]) }, callback);
    await vi.waitFor(() => expect(callback).toHaveBeenCalled());

    const result = callback.mock.calls[0][0] as { data: string };
    const url = new URL(result.data);
    expect(url.pathname).toContain(encodeURIComponent('文档.pdf'));
    expect(url.searchParams.get('filename')).toBe('文档.pdf');
    expect(document.title).toBe('文档');

    const cache = await window.caches.open('onlyoffice-browser-print-pdfs');
    const cachedPdf = await cache.match(result.data);
    const contentDisposition = cachedPdf?.headers.get('content-disposition') || '';
    expect(contentDisposition).toContain('filename="_.pdf"');
    expect(contentDisposition).toContain(`filename*=UTF-8''${encodeURIComponent('文档.pdf')}`);
    const cachedPdfBytes = new Uint8Array(await cachedPdf!.arrayBuffer());
    expect(binaryString(cachedPdfBytes)).toContain(`/Title <${pdfUtf16BeHex('文档')}>`);

    await instance.destroy();
  });

  it('does not run x2t when OnlyOffice print already returns PDF bytes', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const pdfFixture = asciiBytes(`%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R >>
endobj
%%EOF
`);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
    });
    await flush();

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void })
      .printPdf;

    printPdf?.({ data: pdfFixture }, callback);
    await vi.waitFor(() => expect(callback).toHaveBeenCalled());

    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(mocks.convertPrintDataToPdf).not.toHaveBeenCalled();
    const result = callback.mock.calls[0][0] as { data: string };
    expect(new URL(result.data).pathname).toMatch(/alpha\.pdf$/);
    expect(new URL(result.data).searchParams.get('filename')).toBe('alpha.pdf');

    const cache = await window.caches.open('onlyoffice-browser-print-pdfs');
    const cachedPdf = await cache.match(result.data);
    expect(new Uint8Array(await cachedPdf!.arrayBuffer())).toEqual(pdfFixture);

    await instance.destroy();
  });

  it('prefers the editor native PDF exporter for built-in print', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const pdfFixture = asciiBytes(`%PDF-1.7
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R >>
endobj
%%EOF
`);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
    });
    await flush();
    const paddedPdf = new Uint8Array(pdfFixture.byteLength + 4);
    paddedPdf.set(pdfFixture);
    paddedPdf.set([1, 2, 3, 4], pdfFixture.byteLength);
    docEditorInstances[0].asc_nativeGetPDF.mockImplementationOnce(() => {
      (
        window as typeof window & { native?: { Save_End?: (header: string, length: number) => void } }
      ).native?.Save_End?.('', pdfFixture.byteLength);
      return paddedPdf;
    });

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void })
      .printPdf;

    printPdf?.({ data: new Uint8Array([1, 2, 3]) }, callback);
    await vi.waitFor(() => expect(callback).toHaveBeenCalled());

    expect(docEditorInstances[0].asc_nativeGetPDF).toHaveBeenCalledWith(expect.objectContaining({ isPrint: true }));
    expect(mocks.convertBinToDocument).not.toHaveBeenCalled();
    expect(mocks.convertPrintDataToPdf).not.toHaveBeenCalled();
    const result = callback.mock.calls[0][0] as { data: string };
    expect(new URL(result.data).pathname).toMatch(/alpha\.pdf$/);
    expect(new URL(result.data).searchParams.get('filename')).toBe('alpha.pdf');

    const cache = await window.caches.open('onlyoffice-browser-print-pdfs');
    const cachedPdf = await cache.match(result.data);
    expect(new Uint8Array(await cachedPdf!.arrayBuffer())).toEqual(pdfFixture);
    expect('native' in window).toBe(false);

    await instance.destroy();
  });

  it('tracks dirty state changes and clears dirty after manual save', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onDirtyChange = vi.fn();
    const onSave = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      onDirtyChange,
      onSave,
    });
    await flush();

    docEditorConfigs[0].events.onDocumentStateChange?.({ data: true });

    expect(instance.getState().dirty).toBe(true);
    expect(onDirtyChange).toHaveBeenCalledWith(true, instance);

    await instance.save('XLSX');

    expect(instance.getState().dirty).toBe(false);
    expect(onDirtyChange).toHaveBeenLastCalledWith(false, instance);

    await instance.destroy();
  });

  it('opens initial readonly mode through native edit permissions without a rights-change toast', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      readonly: true,
    });
    await flush();

    expect(docEditorConfigs[0].document.permissions.edit).toBe(false);
    expect(docEditorConfigs[0].editorConfig.mode).toBe('edit');
    expect(docEditorInstances[0].processRightsChange).not.toHaveBeenCalled();
    expect(instance.getState().readonly).toBe(true);

    await instance.destroy();
  });

  it('still uses processRightsChange for runtime readonly toggles after the editor is ready', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });
    await flush();

    expect(docEditorConfigs[0].document.permissions.edit).toBe(true);
    expect(docEditorInstances[0].processRightsChange).not.toHaveBeenCalled();

    instance.setReadonly(true);

    expect(docEditorInstances[0].processRightsChange).toHaveBeenCalledWith(false, 'Readonly mode');

    await instance.destroy();
  });

  it.each([
    ['alpha.docx', 'edit', 'de-settings-zoom'],
    ['alpha.docx', 'preview', 'de-settings-zoom'],
    ['slides.pptx', 'readonly', 'pe-settings-zoom'],
  ] as const)('fits %s in %s mode to width by default', async (fileName, mode, storageKey) => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], fileName),
      fileName,
      mode,
    });
    await flush();

    expect(docEditorConfigs[0].editorConfig.customization.zoom).toBe(-2);
    expect(window.localStorage.getItem(storageKey)).toBe('-2');
    expect(docEditorInstances[0].zoomFitToWidth).toHaveBeenCalledTimes(1);

    await instance.destroy();
  });

  it.each([
    ['sheet.xlsx', 'edit'],
    ['slides.pptx', 'preview'],
  ] as const)('does not force fit-to-width for %s in %s mode', async (fileName, mode) => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], fileName),
      fileName,
      mode,
    });
    await flush();

    expect(docEditorConfigs[0].editorConfig.customization.zoom).toBeUndefined();
    expect(window.localStorage.getItem('de-settings-zoom')).toBeNull();
    expect(window.localStorage.getItem('pe-settings-zoom')).toBeNull();
    expect(docEditorInstances[0].zoomFitToWidth).not.toHaveBeenCalled();

    await instance.destroy();
  });
});
