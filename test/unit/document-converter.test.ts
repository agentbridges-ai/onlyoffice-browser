import { afterEach, describe, expect, it, vi } from 'vitest';

import { X2TConverter } from '../../src/lib/document-converter';
import type { EmscriptenModule } from '../../src/lib/document-types';

type FakeX2TModule = EmscriptenModule & {
  files: Map<string, Uint8Array<ArrayBuffer> | string>;
  lastInputBytes: Uint8Array<ArrayBuffer> | null;
  lastMediaFiles: Map<string, Uint8Array<ArrayBuffer> | string>;
  lastParamsXml: string;
};

function createFakeX2TModule(): FakeX2TModule {
  const files = new Map<string, Uint8Array<ArrayBuffer> | string>();
  const dirs = new Set(['/working', '/working/media', '/working/fonts', '/working/themes']);

  const module: FakeX2TModule = {
    files,
    lastInputBytes: null,
    lastMediaFiles: new Map(),
    lastParamsXml: '',
    FS: {
      mkdir(path: string) {
        dirs.add(path);
      },
      readdir(path: string) {
        const prefix = path.endsWith('/') ? path : `${path}/`;
        const entries = new Set<string>(['.', '..']);

        for (const dir of dirs) {
          if (dir.startsWith(prefix)) {
            const child = dir.slice(prefix.length).split('/')[0];
            if (child) entries.add(child);
          }
        }

        for (const file of files.keys()) {
          if (file.startsWith(prefix)) {
            const child = file.slice(prefix.length).split('/')[0];
            if (child) entries.add(child);
          }
        }

        return Array.from(entries);
      },
      readFile(path: string) {
        const file = files.get(path);
        if (!file) throw new Error(`Missing fake file: ${path}`);
        return file;
      },
      writeFile(path: string, data: Uint8Array | string) {
        files.set(path, typeof data === 'string' ? data : new Uint8Array(Array.from(data)));
        if (path === '/working/params.xml' && typeof data === 'string') {
          module.lastParamsXml = data;
        }
      },
      unlink(path: string) {
        files.delete(path);
      },
    },
    ccall: vi.fn((_funcName: string, _returnType: string, _argTypes: string[], args: string[]) => {
      const params = files.get(args[0]);
      if (typeof params !== 'string') return 88;

      const outputPath = params.match(/<m_sFileTo>([^<]+)<\/m_sFileTo>/)?.[1];
      if (!outputPath) return 88;

      const inputPath = params.match(/<m_sFileFrom>([^<]+)<\/m_sFileFrom>/)?.[1];
      const input = inputPath ? files.get(inputPath) : undefined;
      module.lastInputBytes = input instanceof Uint8Array ? input : null;
      module.lastMediaFiles = new Map(
        Array.from(files.entries())
          .filter(([path]) => path.startsWith('/working/media/'))
          .map(([path, value]) => [path, value instanceof Uint8Array ? new Uint8Array(Array.from(value)) : value]),
      );
      files.set(outputPath, new Uint8Array(new ArrayBuffer(4)));
      return 0;
    }),
    onRuntimeInitialized: vi.fn(),
  };

  return module;
}

function mockFetchMedia(mediaByUrl: Record<string, Uint8Array>): void {
  vi.stubGlobal(
    'fetch',
    vi.fn(async (url: string) => {
      const bytes = mediaByUrl[String(url)];
      if (!bytes) {
        return {
          ok: false,
          status: 404,
          statusText: 'Not Found',
          arrayBuffer: async () => new ArrayBuffer(0),
        };
      }
      const copy = new Uint8Array(Array.from(bytes));
      return {
        ok: true,
        status: 200,
        statusText: 'OK',
        arrayBuffer: async () => copy.buffer,
      };
    }),
  );
}

describe('X2TConverter', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it.each([
    { fileName: 'legacy.doc', sourceFormat: 66, targetFormat: 8193, type: 'word' },
    { fileName: 'legacy.xls', sourceFormat: 258, targetFormat: 8194, type: 'cell' },
    { fileName: 'legacy.ppt', sourceFormat: 130, targetFormat: 8195, type: 'slide' },
  ] as const)(
    'passes explicit x2t format ids when converting $fileName to editor bin',
    async ({ fileName, sourceFormat, targetFormat, type }) => {
      const converter = new X2TConverter();
      const x2tModule = createFakeX2TModule();

      (converter as unknown as { x2tModule: EmscriptenModule }).x2tModule = x2tModule;
      vi.spyOn(converter, 'initialize').mockResolvedValue(x2tModule);

      const result = await converter.convertDocument(new File([new Uint8Array([1, 2, 3])], fileName));
      const params = x2tModule.lastParamsXml;

      expect(result.type).toBe(type);
      expect(params).toEqual(expect.stringContaining(`<m_nFormatFrom>${sourceFormat}</m_nFormatFrom>`));
      expect(params).toEqual(expect.stringContaining(`<m_nFormatTo>${targetFormat}</m_nFormatTo>`));
      expect(params).toEqual(expect.stringContaining('<m_sFontDir>/working/fonts/</m_sFontDir>'));
      expect(x2tModule.ccall).toHaveBeenCalledWith('main1', 'number', ['string'], ['/working/params.xml']);
    },
  );

  it('lets x2t infer native save-bin formats when exporting to PDF', async () => {
    const converter = new X2TConverter();
    const x2tModule = createFakeX2TModule();

    (converter as unknown as { x2tModule: EmscriptenModule }).x2tModule = x2tModule;
    vi.spyOn(converter, 'initialize').mockResolvedValue(x2tModule);

    const nativeBytes = new Uint8Array([0x44, 0x4f, 0x43, 0x59, 0x3b, 0x76, 0x31, 0x30, 0x3b, 0x30, 0x3b, 1, 2, 3]);

    await converter.convertBinToDocument(nativeBytes, 'local.docx', 'PDF');
    const params = x2tModule.lastParamsXml;

    expect(params).not.toContain('<m_nFormatFrom>');
    expect(params).not.toContain('<m_nFormatTo>');
    expect(params).toEqual(expect.stringContaining('<m_bIsNoBase64>true</m_bIsNoBase64>'));
    expect(params).toEqual(expect.stringContaining('<m_sFontDir>/working/fonts/</m_sFontDir>'));
    expect(Array.from(x2tModule.lastInputBytes || [])).toEqual(Array.from(nativeBytes));
    expect(x2tModule.ccall).toHaveBeenCalledWith('main1', 'number', ['string'], ['/working/params.xml']);
  });

  it('passes native base64 save-bin payloads through for x2t inference', async () => {
    const converter = new X2TConverter();
    const x2tModule = createFakeX2TModule();
    const nativeBase64 = btoa('XLSY;v10;0;\x07\x06\x8b\x02');
    const nativeBase64Bytes = Uint8Array.from(nativeBase64, (char) => char.charCodeAt(0) & 0xff);

    (converter as unknown as { x2tModule: EmscriptenModule }).x2tModule = x2tModule;
    vi.spyOn(converter, 'initialize').mockResolvedValue(x2tModule);

    await converter.convertBinToDocument(nativeBase64Bytes, 'legacy.xls', 'PDF');
    const params = x2tModule.lastParamsXml;

    expect(params).not.toContain('<m_nFormatFrom>');
    expect(params).not.toContain('<m_nFormatTo>');
    expect(params).toEqual(expect.stringContaining('<m_bIsNoBase64>false</m_bIsNoBase64>'));
    expect(params).toEqual(expect.stringContaining('<m_sFontDir>/working/fonts/</m_sFontDir>'));
    expect(Array.from(x2tModule.lastInputBytes || [])).toEqual(Array.from(nativeBase64Bytes));
  });

  it('decodes native base64 save-bin payloads for document exports', async () => {
    const converter = new X2TConverter();
    const x2tModule = createFakeX2TModule();
    const nativeText = 'XLSY;v10;0;\x07\x06\x8b\x02';
    const nativeBase64 = btoa(nativeText);
    const nativeBase64Bytes = Uint8Array.from(nativeBase64, (char) => char.charCodeAt(0) & 0xff);

    (converter as unknown as { x2tModule: EmscriptenModule }).x2tModule = x2tModule;
    vi.spyOn(converter, 'initialize').mockResolvedValue(x2tModule);

    const result = await converter.convertBinToDocument(nativeBase64Bytes, 'legacy.xls', 'XLS');
    const params = x2tModule.lastParamsXml;

    expect(params).not.toContain('<m_nFormatFrom>');
    expect(params).not.toContain('<m_nFormatTo>');
    expect(params).toEqual(expect.stringContaining('<m_bIsNoBase64>true</m_bIsNoBase64>'));
    expect(result.fileName).toBe('legacy.xlsx');
    expect(Array.from(x2tModule.lastInputBytes || [])).toEqual(
      Array.from(Uint8Array.from(nativeText, (char) => char.charCodeAt(0) & 0xff)),
    );
  });

  it('restores media files into the x2t workspace before native document export', async () => {
    const converter = new X2TConverter();
    const x2tModule = createFakeX2TModule();
    const imageBytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

    (converter as unknown as { x2tModule: EmscriptenModule }).x2tModule = x2tModule;
    vi.spyOn(converter, 'initialize').mockResolvedValue(x2tModule);
    mockFetchMedia({ 'blob:image-1': imageBytes });

    const nativeBytes = new Uint8Array([0x44, 0x4f, 0x43, 0x59, 0x3b, 0x76, 0x31, 0x30, 0x3b, 0x30, 0x3b, 1, 2, 3]);

    await converter.convertBinToDocument(nativeBytes, 'with-image.docx', 'DOCX', {
      'media/image%201.png': 'blob:image-1',
    });

    expect(x2tModule.lastMediaFiles.get('/working/media/image 1.png')).toEqual(imageBytes);
    expect(globalThis.fetch).toHaveBeenCalledWith('blob:image-1');
  });

  it('restores media files into the x2t workspace before native print export', async () => {
    const converter = new X2TConverter();
    const x2tModule = createFakeX2TModule();
    const chartBytes = new Uint8Array([0x43, 0x48, 0x41, 0x52, 0x54]);

    (converter as unknown as { x2tModule: EmscriptenModule }).x2tModule = x2tModule;
    vi.spyOn(converter, 'initialize').mockResolvedValue(x2tModule);
    mockFetchMedia({ 'blob:chart': chartBytes });

    await converter.convertPrintDataToPdf(new Uint8Array([1, 2, 3]), 'with-chart.docx', {
      'media/charts/chart1.png': 'blob:chart',
    });

    expect(x2tModule.lastMediaFiles.get('/working/media/charts/chart1.png')).toEqual(chartBytes);
    expect(globalThis.fetch).toHaveBeenCalledWith('blob:chart');
  });
});
