import { beforeEach, describe, expect, it, vi } from 'vitest';

const mocks = vi.hoisted(() => ({
  assertGeneratedFontAssetsAvailable: vi.fn(),
  convertBinToDocument: vi.fn(),
  convertDocument: vi.fn(),
  initX2T: vi.fn(),
}));

vi.mock('../../src/lib/font-assets', () => ({
  assertGeneratedFontAssetsAvailable: mocks.assertGeneratedFontAssetsAvailable,
}));

vi.mock('../../src/lib/converter', () => ({
  convertBinToDocument: mocks.convertBinToDocument,
  convertDocument: mocks.convertDocument,
  initX2T: mocks.initX2T,
}));

import { createOfficeEditor } from '../../src/lib/office-editor-runtime';

function flush(): Promise<void> {
  return Promise.resolve().then(() => undefined);
}

function waitForMessage(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
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
  };
};

type CapturedDocEditorInstance = {
  connectMockServer: ReturnType<typeof vi.fn>;
  destroyEditor: ReturnType<typeof vi.fn>;
  downloadAs: ReturnType<typeof vi.fn>;
  processRightsChange: ReturnType<typeof vi.fn>;
  asc_nativeGetFile3: ReturnType<typeof vi.fn>;
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
    mocks.assertGeneratedFontAssetsAvailable.mockResolvedValue(undefined);
    mocks.initX2T.mockResolvedValue(undefined);
    mocks.convertDocument.mockImplementation(async (file: File) => ({
      fileName: file.name,
      bin: new Uint8Array([1, 2, 3]),
      media: {},
    }));
    mocks.convertBinToDocument.mockImplementation(async (bin: Uint8Array, fileName: string, targetExt = 'DOCX') => ({
      fileName: fileName.replace(/\.[^/.]+$/, `.${targetExt.toLowerCase()}`),
      data: bin,
    }));

    class MockDocEditor {
      connectMockServer = vi.fn();
      destroyEditor = vi.fn();
      downloadAs = vi.fn();
      processRightsChange = vi.fn();
      asc_nativeGetFile3 = vi.fn(() => ({ data: new Uint8Array([9, 8, 7]) }));
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
    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.xlsx', 'XLSX');
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onSave.mock.calls[0][0]).toMatchObject({
      name: 'alpha.xlsx',
      size: 3,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    expect(savedFile).toMatchObject({
      name: 'alpha.xlsx',
      size: 3,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });

    await instance.destroy();
  });

  it('falls back to native OOXML bytes when x2t rejects an already zipped save payload', async () => {
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
    mocks.convertBinToDocument.mockRejectedValueOnce(new Error('Conversion failed with code: 88'));

    const savedFile = await instance.save('XLSX');

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

    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.xlsx', 'XLSX');
    expect(onSave).toHaveBeenCalledTimes(1);
    expect(onSave.mock.calls[0][0]).toMatchObject({
      name: 'alpha.xlsx',
      size: 3,
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    expect(instance.getState().dirty).toBe(false);
    expect(onDirtyChange).toHaveBeenLastCalledWith(false, instance);

    await instance.destroy();
  });

  it('exports print payloads to PDF object URLs through the OnlyOffice parent app bridge', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });
    await flush();

    const callback = vi.fn();
    const printPdf = (window.APP as { printPdf?: (data: unknown, callback: (result: unknown) => void) => void }).printPdf;
    expect(printPdf).toEqual(expect.any(Function));

    printPdf?.({ data: new Uint8Array([1, 2, 3]) }, callback);
    await waitForMessage();

    expect(mocks.convertBinToDocument).toHaveBeenCalledWith(expect.any(Uint8Array), 'alpha.docx', 'PDF');
    expect(callback).toHaveBeenCalledWith({
      type: 'save',
      status: 'ok',
      data: expect.stringMatching(/^blob:/),
      filetype: 513,
    });

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
