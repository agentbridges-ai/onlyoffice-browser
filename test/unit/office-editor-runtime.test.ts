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
    customization: {
      zoom?: number;
      spellcheck?: boolean;
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
      additionalData?: { outputformat?: unknown; title?: unknown },
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

  beforeEach(() => {
    document.head.innerHTML = '';
    document.body.innerHTML = '';
    window.localStorage.clear();
    docEditorConfigs.length = 0;
    docEditorInstances.length = 0;
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

  it('passes Download As files to the host callback in embedded mode', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);
    const onDownload = vi.fn();

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
      saveBehavior: 'download',
      onDownload,
    });
    await flush();

    mocks.convertBinToDocument.mockResolvedValueOnce({
      fileName: 'alpha.odt',
      data: zipFixture(),
    });

    docEditorConfigs[0].events.onRequestSaveAs?.({ data: { fileType: 67, title: 'alpha.odt' } });
    await waitForMessage();

    expect(onDownload).toHaveBeenCalledTimes(1);
    expect(onDownload.mock.calls[0][0]).toMatchObject({ name: 'alpha.odt' });
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

  it('lets Spreadsheet PDF settings confirmations reach upstream before the final asc_DownloadAs call', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.xlsx'),
      fileName: 'alpha.xlsx',
      mode: 'edit',
      saveBehavior: 'download',
    });
    await flush();

    const editorFrame = container.querySelector<HTMLIFrameElement>('iframe[name="frameEditor"]');
    const frameDocument = editorFrame!.contentDocument!;
    const formatButton = frameDocument.createElement('div');
    formatButton.className = 'btn-doc-format';
    formatButton.setAttribute('format', '513');
    frameDocument.body.appendChild(formatButton);
    formatButton.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, cancelable: true }));

    const dialog = frameDocument.createElement('div');
    dialog.setAttribute('role', 'dialog');
    dialog.append('PDF settings');
    const saveAndDownload = frameDocument.createElement('button');
    saveAndDownload.textContent = 'Save & Download';
    dialog.appendChild(saveAndDownload);
    frameDocument.body.appendChild(dialog);

    const clickEvent = new MouseEvent('click', { bubbles: true, cancelable: true });
    const notCancelled = saveAndDownload.dispatchEvent(clickEvent);
    await flush();

    expect(notCancelled).toBe(true);
    expect(clickEvent.defaultPrevented).toBe(false);
    expect(docEditorInstances[0].asc_nativeGetPDF).not.toHaveBeenCalled();
    expect(mocks.convertPrintDataToPdf).not.toHaveBeenCalled();
    expect(HTMLAnchorElement.prototype.click).not.toHaveBeenCalled();

    const printStream = new Uint8Array([0xa3, 0, 0, 0, 1, 2, 3, 4]);
    const pdfBytes = validOnePagePdfFixture();
    const advancedOptions = { range: 'active-sheets' };
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

    docEditorInstances[0].nativeApi.asc_DownloadAs({ fileType: 513, advancedOptions });
    await waitForMessage();

    expect(docEditorInstances[0].asc_nativeGetPDF).toHaveBeenCalledWith(
      expect.objectContaining({
        isPrint: true,
        advancedOptions,
        printOptions: advancedOptions,
      }),
    );
    expect(mocks.convertPrintDataToPdf).toHaveBeenCalledWith(printStream, 'alpha.xlsx', {});
    expect(HTMLAnchorElement.prototype.click).toHaveBeenCalledTimes(1);

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
    expect(downloadedFile).toMatchObject({ name: 'alpha.zip' });
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
    ['slides.pptx', 'readonly', 'pe-settings-zoom'],
  ] as const)('fits %s editor-mode preview to width by default', async (fileName, mode, storageKey) => {
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
    ['alpha.docx', 'preview'],
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
