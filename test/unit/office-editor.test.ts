import { beforeEach, describe, expect, it, vi } from 'vitest';

const { mockConvertDocument, mockConvertBinToDocument, mockInitX2T } = vi.hoisted(() => ({
  mockConvertDocument: vi.fn(),
  mockConvertBinToDocument: vi.fn(),
  mockInitX2T: vi.fn(),
}));

vi.mock('../../src/lib/converter', () => ({
  convertDocument: mockConvertDocument,
  convertBinToDocument: mockConvertBinToDocument,
  initX2T: mockInitX2T,
}));

import { createOfficeEditor, loadOfficeEditorApi } from '../../src/lib/office-editor';

type EditorRecord = {
  elementId: string;
  config: any;
  editor: {
    connectMockServer: ReturnType<typeof vi.fn>;
    destroyEditor: () => void;
    downloadAs: ReturnType<typeof vi.fn>;
    processRightsChange: ReturnType<typeof vi.fn>;
    serviceCommand: ReturnType<typeof vi.fn>;
  };
  destroyEditorSpy: ReturnType<typeof vi.fn>;
};

function setupDocsApi(): EditorRecord[] {
  const records: EditorRecord[] = [];
  (window as any).DocsAPI = {
    DocEditor: vi.fn(function (elementId: string, config: any) {
      const destroyEditorSpy = vi.fn();
      const editor = {
        connectMockServer: vi.fn(),
        destroyEditor: destroyEditorSpy,
        downloadAs: vi.fn(),
        processRightsChange: vi.fn(),
        serviceCommand: vi.fn(),
      };
      records.push({ elementId, config, editor, destroyEditorSpy });
      return editor;
    }),
  };
  return records;
}

function makeSaveEvent(bytes: number[], outputformat = 65) {
  return {
    data: {
      data: {
        data: new Uint8Array(bytes),
      },
      option: {
        outputformat,
      },
    },
  };
}

describe('office-editor component API', () => {
  let objectUrlCounter = 0;

  beforeEach(() => {
    document.head.innerHTML = '';
    document.body.innerHTML = '';
    delete (window as any).DocsAPI;
    delete (window as any).APP;
    objectUrlCounter = 0;

    vi.mocked(URL.createObjectURL).mockImplementation(() => `blob:test-${++objectUrlCounter}`);
    vi.mocked(URL.revokeObjectURL).mockClear();
    mockInitX2T.mockResolvedValue({ FS: {}, ccall: vi.fn(), onRuntimeInitialized: vi.fn() });
    mockConvertDocument.mockImplementation(async (file: File) => ({
      fileName: file.name,
      type: file.name.endsWith('.pptx') ? 'slide' : file.name.endsWith('.xlsx') ? 'cell' : 'word',
      bin: new Uint8Array([1, 2, 3]),
      media: {
        [`media/${file.name}.png`]: `blob:media-${file.name}`,
      },
    }));
    mockConvertBinToDocument.mockImplementation(async (_bin: Uint8Array, fileName: string, targetExt = 'DOCX') => ({
      fileName: `${fileName.replace(/\.[^/.]+$/, '')}.${targetExt.toLowerCase()}`,
      data: new Uint8Array([9, 8, 7]),
    }));
  });

  it('loads the OnlyOffice API script only once', async () => {
    const first = loadOfficeEditorApi();
    const second = loadOfficeEditorApi();
    const scripts = document.querySelectorAll('script[data-office-editor-api="true"]');

    expect(scripts).toHaveLength(1);
    scripts[0].dispatchEvent(new Event('load'));

    await expect(first).resolves.toBeUndefined();
    await expect(second).resolves.toBeUndefined();
  });

  it('creates isolated editor instances with independent media and mock servers', async () => {
    const records = setupDocsApi();
    const firstContainer = document.createElement('div');
    const secondContainer = document.createElement('div');
    document.body.append(firstContainer, secondContainer);

    const first = await createOfficeEditor(firstContainer, {
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
    });
    const second = await createOfficeEditor(secondContainer, {
      file: new File(['b'], 'deck.pptx'),
      fileName: 'deck.pptx',
      readonly: true,
    });

    expect(first.getState()).toMatchObject({ fileName: 'alpha.docx', fileType: 'docx', readonly: false });
    expect(second.getState()).toMatchObject({ fileName: 'deck.pptx', fileType: 'pptx', readonly: true });
    expect(records).toHaveLength(2);
    expect(records[0].elementId).not.toBe(records[1].elementId);
    expect(records[0].editor.connectMockServer).toHaveBeenCalledTimes(1);
    expect(records[1].editor.connectMockServer).toHaveBeenCalledTimes(1);

    const firstServer = records[0].editor.connectMockServer.mock.calls[0][0];
    const secondServer = records[1].editor.connectMockServer.mock.calls[0][0];
    await expect(firstServer.getImageURL('alpha.docx.png')).resolves.toBe('blob:media-alpha.docx');
    await expect(secondServer.getImageURL('deck.pptx.png')).resolves.toBe('blob:media-deck.pptx');

    const app = (window as any).APP;
    await new Promise<void>((resolve) => {
      app.getImageURL('deck.pptx.png', (url: string) => {
        expect(url).toBe('blob:media-deck.pptx');
        resolve();
      });
    });

    first.destroy();
    second.destroy();
  });

  it('keeps simultaneous save requests scoped to their own instance', async () => {
    const records = setupDocsApi();
    const first = await createOfficeEditor(document.createElement('div'), {
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
    });
    const second = await createOfficeEditor(document.createElement('div'), {
      file: new File(['b'], 'deck.pptx'),
      fileName: 'deck.pptx',
    });

    const firstSave = first.save('DOCX');
    const secondSave = second.save('PPTX');

    records[1].config.events.onSave(makeSaveEvent([2], 129));
    records[0].config.events.onSave(makeSaveEvent([1], 65));

    await expect(firstSave).resolves.toMatchObject({ name: 'alpha.docx' });
    await expect(secondSave).resolves.toMatchObject({ name: 'deck.pptx' });
    expect(records[0].editor.downloadAs).toHaveBeenCalledWith('DOCX');
    expect(records[1].editor.downloadAs).toHaveBeenCalledWith('PPTX');

    first.destroy();
    second.destroy();
  });

  it('keeps readonly as a reversible edit-runtime mode', async () => {
    const records = setupDocsApi();
    const instance = await createOfficeEditor(document.createElement('div'), {
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      readonly: true,
    });

    expect(records[0].config.editorConfig.mode).toBe('edit');
    expect(records[0].config.document.permissions.edit).toBe(true);
    expect(instance.getState()).toMatchObject({ mode: 'readonly', readonly: true });

    records[0].config.events.onDocumentReady();
    expect(records[0].editor.processRightsChange).toHaveBeenLastCalledWith(false, 'Readonly mode');

    instance.setReadonly(false);
    expect(instance.getState()).toMatchObject({ mode: 'edit', readonly: false });
    expect(records[0].editor.processRightsChange).toHaveBeenLastCalledWith(true, '');

    instance.setReadonly(true);
    expect(instance.getState()).toMatchObject({ mode: 'readonly', readonly: true });
    expect(records[0].editor.processRightsChange).toHaveBeenLastCalledWith(false, 'Readonly mode');

    instance.destroy();
  });

  it('uses the upstream embedded viewer only for explicit preview instances', async () => {
    const records = setupDocsApi();
    const instance = await createOfficeEditor(document.createElement('div'), {
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
      mode: 'preview',
    });

    expect(records[0].config.type).toBe('embedded');
    expect(records[0].config.editorConfig.mode).toBe('view');
    expect(records[0].config.editorConfig.embedded).toMatchObject({
      autostart: 'document',
      toolbarDocked: 'top',
    });
    expect(records[0].config.document.permissions.edit).toBe(false);
    expect(instance.getState()).toMatchObject({ mode: 'preview', readonly: true });

    records[0].config.events.onDocumentReady();
    expect(records[0].editor.processRightsChange).not.toHaveBeenCalled();
    expect(() => instance.setReadonly(false)).toThrow('Preview mode cannot be switched to editing');
    await expect(instance.save()).rejects.toThrow('Current document is readonly');

    instance.destroy();
  });

  it('destroys idempotently and revokes all local object URLs', async () => {
    const records = setupDocsApi();
    const container = document.createElement('div');
    document.body.appendChild(container);

    const instance = await createOfficeEditor(container, {
      file: new File(['a'], 'alpha.docx'),
      fileName: 'alpha.docx',
    });

    instance.destroy();
    instance.destroy();

    expect(records[0].destroyEditorSpy).toHaveBeenCalledTimes(1);
    expect(container.childElementCount).toBe(0);
    expect(instance.getState()).toMatchObject({ destroyed: true, status: 'destroyed' });
    expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:test-1');
    expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:test-2');
    expect(URL.revokeObjectURL).toHaveBeenCalledWith('blob:media-alpha.docx');
  });
});
