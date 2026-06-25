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
    customization: {
      zoom?: number;
      spellcheck?: boolean;
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
  };
};

type CapturedDocEditorInstance = {
  connectMockServer: ReturnType<typeof vi.fn>;
  destroyEditor: ReturnType<typeof vi.fn>;
  processRightsChange: ReturnType<typeof vi.fn>;
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

    class MockDocEditor {
      connectMockServer = vi.fn();
      destroyEditor = vi.fn();
      processRightsChange = vi.fn();
      zoomFitToWidth = vi.fn();

      constructor(_elementId: string, config: CapturedDocEditorConfig) {
        docEditorConfigs.push(config);
        docEditorInstances.push(this);
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

  it('disables OnlyOffice spellcheck by default and prevents turning it on from the UI', async () => {
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['hello'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'edit',
    });

    expect(docEditorConfigs).toHaveLength(1);
    expect(docEditorConfigs[0].editorConfig.customization).toMatchObject({
      compactToolbar: true,
      spellcheck: false,
      features: {
        featuresTips: false,
        spellcheck: {
          change: false,
        },
      },
    });

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
