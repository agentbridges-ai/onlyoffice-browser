import { describe, expect, it, vi } from 'vitest';

import { X2TConverter } from '../../src/lib/document-converter';
import type { EmscriptenModule } from '../../src/lib/document-types';

type FakeX2TModule = EmscriptenModule & {
  files: Map<string, Uint8Array<ArrayBuffer> | string>;
  lastParamsXml: string;
};

function createFakeX2TModule(): FakeX2TModule {
  const files = new Map<string, Uint8Array<ArrayBuffer> | string>();
  const dirs = new Set(['/working', '/working/media', '/working/fonts', '/working/themes']);

  const module: FakeX2TModule = {
    files,
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

      files.set(outputPath, new Uint8Array(new ArrayBuffer(4)));
      return 0;
    }),
    onRuntimeInitialized: vi.fn(),
  };

  return module;
}

describe('X2TConverter', () => {
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
});
