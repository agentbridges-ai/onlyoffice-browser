import type {
  BinConversionResult,
  ConversionResult,
  DocumentMediaMap,
  DocumentType,
  EmscriptenModule,
} from './document-types';
import { BASE_PATH, DOCUMENT_TYPE_MAP } from './document-utils';
import { oAscFileType } from './file-types';
import {
  fetchGeneratedFontAssetsManifest,
  fetchGeneratedFontSourceMap,
  fetchRuntimeBinaryAsset,
  getAssetFileName,
} from './font-assets';

function createObjectURL(blob: Blob): string {
  return URL.createObjectURL(blob);
}

function getExtensions(mimeType = ''): string[] {
  const map: Record<string, string[]> = {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['docx'],
    'application/msword': ['doc'],
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['xlsx'],
    'application/vnd.ms-excel': ['xls'],
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['pptx'],
    'application/vnd.ms-powerpoint': ['ppt'],
    'text/csv': ['csv'],
  };
  return map[mimeType] || [];
}

function loadScriptOnce(src: string): Promise<void> {
  const existing = Array.from(document.scripts).find((script) => script.src === src);
  if (existing) {
    return Promise.resolve();
  }

  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Failed to load script: ${src}`));
    document.head.appendChild(script);
  });
}

function decodeLatin1(data: Uint8Array): string {
  const chunkSize = 0x8000;
  let result = '';
  for (let index = 0; index < data.length; index += chunkSize) {
    result += String.fromCharCode(...data.slice(index, index + chunkSize));
  }
  return result;
}

function safeDecodeURIComponent(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

const X2T_SOURCE_FORMAT_BY_EXTENSION: Record<string, number> = {
  csv: oAscFileType.CSV,
  doc: oAscFileType.DOC,
  docm: oAscFileType.DOCM,
  docx: oAscFileType.DOCX,
  docxf: oAscFileType.DOCXF,
  dotx: oAscFileType.DOTX,
  epub: oAscFileType.EPUB,
  fb2: oAscFileType.FB2,
  html: oAscFileType.HTML,
  md: oAscFileType.MD,
  odp: oAscFileType.ODP,
  ods: oAscFileType.ODS,
  oform: oAscFileType.OFORM,
  odt: oAscFileType.ODT,
  ott: oAscFileType.OTT,
  otp: oAscFileType.OTP,
  ots: oAscFileType.OTS,
  potm: oAscFileType.POTM,
  potx: oAscFileType.POTX,
  ppt: oAscFileType.PPT,
  pptm: oAscFileType.PPTM,
  pptx: oAscFileType.PPTX,
  ppsm: oAscFileType.PPSM,
  ppsx: oAscFileType.PPSX,
  rtf: oAscFileType.RTF,
  txt: oAscFileType.TXT,
  xls: oAscFileType.XLS,
  xlsb: oAscFileType.XLSB,
  xlsm: oAscFileType.XLSM,
  xlsx: oAscFileType.XLSX,
  xltm: oAscFileType.XLTM,
  xltx: oAscFileType.XLTX,
};

const X2T_TARGET_FORMAT_BY_EXTENSION: Record<string, number> = {
  ...X2T_SOURCE_FORMAT_BY_EXTENSION,
  jpeg: oAscFileType.JPG,
  jpg: oAscFileType.JPG,
  pdf: oAscFileType.PDF,
  pdfa: oAscFileType.PDFA,
  png: oAscFileType.PNG,
};

const X2T_OUTPUT_EXTENSION_BY_TARGET_EXTENSION: Record<string, string> = {
  jpeg: 'jpg',
  pdfa: 'pdf',
};

const NATIVE_BASE64_PASSTHROUGH_TARGETS = new Set(['pdf', 'pdfa', 'jpg', 'jpeg', 'png']);
const CONVERTER_ALL_FONTS_PATH = '/server/FileConverter/bin/AllFonts.js';
const GENERATED_CONVERTER_ALL_FONTS_ASSET = 'server/FileConverter/bin/AllFonts.js';
const CONVERTER_TEMP_DIR = '/tmp/x2t-conversion';

const LEGACY_TARGET_INTERMEDIATE_EXTENSION: Record<string, string> = {
  doc: 'docx',
  ppt: 'pptx',
  xls: 'xlsx',
};

const X2T_NATIVE_FORMAT_BY_SIGNATURE: Record<string, number> = {
  DOCY: oAscFileType.CANVAS_WORD,
  XLSY: oAscFileType.CANVAS_SPREADSHEET,
  PPTY: oAscFileType.CANVAS_PRESENTATION,
};

const X2T_CANVAS_FORMAT_BY_DOCUMENT_TYPE: Record<DocumentType, number> = {
  cell: oAscFileType.CANVAS_SPREADSHEET,
  slide: oAscFileType.CANVAS_PRESENTATION,
  word: oAscFileType.CANVAS_WORD,
};

function blobPartStartsWith(data: BlobPart, signature: number[]): boolean {
  if (!(data instanceof Uint8Array) && !(data instanceof ArrayBuffer)) return false;
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.byteLength < signature.length) return false;
  return signature.every((value, index) => bytes[index] === value);
}

function getExtensionFileNameSuffix(fileName: string): string {
  const extension = fileName.split(/[\\/]/).pop()?.match(/\.([^.]+)$/)?.[1] || '';
  const normalized = extension.toLowerCase().replace(/[^a-z0-9]+/g, '');
  return normalized ? `_${normalized}` : '';
}

function getTargetExtensionSuffix(targetExt: string): string {
  const normalizedTargetExt = targetExt.toLowerCase() === 'jpeg' ? 'jpg' : targetExt.toLowerCase();
  const normalized = normalizedTargetExt.replace(/[^a-z0-9]+/g, '');
  return normalized ? `_${normalized}` : '';
}

function resolveGeneratedOutputFileName(
  fileName: string,
  data: BlobPart,
  targetExt: string,
  sourceFileName = fileName,
): string {
  const normalizedTargetExt = targetExt.toLowerCase();
  if (!['jpg', 'jpeg', 'png'].includes(normalizedTargetExt)) return fileName;
  if (!blobPartStartsWith(data, [0x50, 0x4b, 0x03, 0x04])) return fileName;
  return fileName.replace(
    /\.[^/.]+$/,
    `${getExtensionFileNameSuffix(sourceFileName)}${getTargetExtensionSuffix(normalizedTargetExt)}.zip`,
  );
}

export class X2TConverter {
  private x2tModule: EmscriptenModule | null = null;
  private isReady = false;
  private initPromise: Promise<EmscriptenModule> | null = null;
  private hasScriptLoaded = false;
  private hasGeneratedFontAssetsLoaded = false;
  private hasThemeAssetsLoaded = false;

  // Supported file type mapping
  private readonly DOCUMENT_TYPE_MAP: Record<string, DocumentType> = DOCUMENT_TYPE_MAP;

  private readonly WORKING_DIRS = [
    '/tmp',
    CONVERTER_TEMP_DIR,
    '/working',
    '/working/media',
    '/working/fonts',
    '/working/themes',
  ];
  private readonly SCRIPT_PATH = `${BASE_PATH}wasm/x2t/x2t.js`;
  private readonly INIT_TIMEOUT = 300000;

  /**
   * Load X2T script file.
   */
  async loadScript(): Promise<void> {
    if (this.hasScriptLoaded) return;

    try {
      const absolutePath = new URL(this.SCRIPT_PATH, window.location.href).href;
      await loadScriptOnce(absolutePath);
      this.hasScriptLoaded = true;
      console.log('X2T WASM script loaded successfully');
    } catch (error) {
      const errorMsg = 'Failed to load X2T WASM script';
      console.error(errorMsg, error);
      throw new Error(errorMsg);
    }
  }

  /**
   * Initialize X2T module
   */
  async initialize(): Promise<EmscriptenModule> {
    if (this.isReady && this.x2tModule) {
      return this.x2tModule;
    }

    // Prevent duplicate initialization
    if (this.initPromise) {
      return this.initPromise;
    }

    this.initPromise = this.doInitialize();
    return this.initPromise;
  }

  private async doInitialize(): Promise<EmscriptenModule> {
    try {
      await this.loadScript();
      return new Promise((resolve, reject) => {
        const x2t = window.Module;
        if (!x2t) {
          reject(new Error('X2T module not found after script loading'));
          return;
        }

        // Set timeout handling
        const timeoutId = setTimeout(() => {
          if (!this.isReady) {
            reject(new Error(`X2T initialization timeout after ${this.INIT_TIMEOUT}ms`));
          }
        }, this.INIT_TIMEOUT);

        x2t.onRuntimeInitialized = () => {
          void (async () => {
            try {
              clearTimeout(timeoutId);
              this.createWorkingDirectories(x2t);
              this.x2tModule = x2t;
              await this.loadGeneratedFontAssets();
              await this.loadThemeAssets();
              this.isReady = true;
              console.log('X2T module initialized successfully');
              resolve(x2t);
            } catch (error) {
              reject(error);
            }
          })();
        };
      });
    } catch (error) {
      this.initPromise = null; // Reset to allow retry
      throw error;
    }
  }

  /**
   * Create working directories
   */
  private createWorkingDirectories(x2t: EmscriptenModule): void {
    this.WORKING_DIRS.forEach((dir) => {
      try {
        x2t.FS.mkdir(dir);
      } catch (error) {
        // Directory may already exist, ignore error
        console.warn(`Directory ${dir} may already exist:`, error);
      }
    });
  }

  private ensureDirectory(dirPath: string): void {
    if (!this.x2tModule || !dirPath || dirPath === '/') return;

    const parts = dirPath.split('/').filter(Boolean);
    let current = dirPath.startsWith('/') ? '' : '.';
    for (const part of parts) {
      current = current === '' ? `/${part}` : `${current}/${part}`;
      try {
        this.x2tModule.FS.mkdir(current);
      } catch {
        // Existing directories are fine.
      }
    }
  }

  private writeBinaryFile(targetPath: string, data: Uint8Array): void {
    try {
      const normalizedPath = targetPath.replace(/\\/g, '/');
      const dirPath = normalizedPath.slice(0, normalizedPath.lastIndexOf('/'));
      this.ensureDirectory(dirPath);
      this.x2tModule?.FS.writeFile(targetPath, data);
    } catch (error) {
      console.warn(`Failed to write generated font asset to ${targetPath}:`, error);
    }
  }

  private firstFont(fontDataByName: Map<string, Uint8Array>, names: string[]): Uint8Array | undefined {
    for (const name of names) {
      const data = fontDataByName.get(name);
      if (data) return data;
    }
    return undefined;
  }

  private writeFontAliases(fontDataByName: Map<string, Uint8Array>): void {
    const regularSans = this.firstFont(fontDataByName, [
      'arial.ttf',
      'calibri.ttf',
      'aptos.ttf',
      'dejavusans.ttf',
      'dejavu_sans.ttf',
    ]);
    const boldSans = this.firstFont(fontDataByName, [
      'arial_bold.ttf',
      'arial-bold.ttf',
      'arialbd.ttf',
      'calibrib.ttf',
      'aptos-bold.ttf',
      'dejavusans-bold.ttf',
      'dejavu_sans_bold.ttf',
    ]);
    const italicSans =
      this.firstFont(fontDataByName, [
        'arial_italic.ttf',
        'arial-italic.ttf',
        'ariali.ttf',
        'calibrii.ttf',
        'aptos-italic.ttf',
        'dejavusans-oblique.ttf',
        'dejavu_sans_oblique.ttf',
      ]) || regularSans;
    const boldItalicSans =
      this.firstFont(fontDataByName, [
        'arial_bold_italic.ttf',
        'arial-bold-italic.ttf',
        'arialbi.ttf',
        'calibriz.ttf',
        'aptos-bold-italic.ttf',
        'dejavusans-boldoblique.ttf',
        'dejavu_sans_bold_oblique.ttf',
      ]) ||
      boldSans ||
      italicSans;
    const serif =
      this.firstFont(fontDataByName, [
        'times_new_roman.ttf',
        'times-new-roman.ttf',
        'times.ttf',
        'cambria.ttc',
        'cambria.ttf',
      ]) || regularSans;

    const aliases: Array<[string, Uint8Array | undefined]> = [
      ['/usr/share/fonts/truetype/msttcorefonts/Arial.ttf', regularSans],
      ['/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf', boldSans],
      ['/usr/share/fonts/truetype/msttcorefonts/Arial_Italic.ttf', italicSans],
      ['/usr/share/fonts/truetype/msttcorefonts/Arial_Bold_Italic.ttf', boldItalicSans],
      ['/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf', regularSans],
      ['/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf', boldSans],
      ['/usr/share/fonts/truetype/liberation/LiberationSans-Italic.ttf', italicSans],
      ['/usr/share/fonts/truetype/liberation/LiberationSans-BoldItalic.ttf', boldItalicSans],
      ['/var/www/onlyoffice/documentserver/core-fonts/dejavu/DejaVuSans.ttf', regularSans],
      ['/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', boldSans],
      ['/usr/share/fonts/onlyoffice-browser-extra/Times New Roman.ttf', serif],
      ['/usr/share/fonts/truetype/msttcorefonts/Times_New_Roman.ttf', serif],
    ];

    for (const [targetPath, data] of aliases) {
      if (data) this.writeBinaryFile(targetPath, data);
    }
  }

  private async loadGeneratedFontAssets(): Promise<void> {
    if (this.hasGeneratedFontAssetsLoaded || !this.x2tModule) return;
    this.hasGeneratedFontAssetsLoaded = true;

    const manifest = await fetchGeneratedFontAssetsManifest();
    const sourceMap = await fetchGeneratedFontSourceMap(manifest.fontSourceMap);
    const sourceByAssetPath = new Map((sourceMap?.fonts || []).map((font) => [font.file, font.source]));
    const fontDataByName = new Map<string, Uint8Array>();

    for (const fontPath of manifest.fonts) {
      try {
        const data = await fetchRuntimeBinaryAsset(fontPath);
        this.writeBinaryFile(`/working/fonts/${getAssetFileName(fontPath)}`, data);
        const sourcePath = sourceByAssetPath.get(fontPath);
        if (sourcePath) {
          this.writeBinaryFile(sourcePath, data);
          fontDataByName.set(getAssetFileName(sourcePath).toLowerCase(), data);
        }
      } catch (error) {
        throw new Error(
          `Failed to load generated font asset ${fontPath}: ${error instanceof Error ? error.message : String(error)}`,
        );
      }
    }
    this.writeFontAliases(fontDataByName);

    try {
      let allFontsData: Uint8Array;
      try {
        allFontsData = await fetchRuntimeBinaryAsset(GENERATED_CONVERTER_ALL_FONTS_ASSET);
      } catch {
        allFontsData = await fetchRuntimeBinaryAsset(manifest.allFonts);
      }
      this.writeBinaryFile(CONVERTER_ALL_FONTS_PATH, allFontsData);
    } catch (error) {
      throw new Error(
        `Failed to load generated AllFonts asset: ${error instanceof Error ? error.message : String(error)}`,
      );
    }

    try {
      const data = await fetchRuntimeBinaryAsset(manifest.fontSelection);
      this.writeBinaryFile('/font_selection.bin', data);
      this.writeBinaryFile('/working/font_selection.bin', data);
      this.writeBinaryFile('/working/fonts/font_selection.bin', data);
    } catch (error) {
      throw new Error(
        `Failed to load generated font selection asset ${manifest.fontSelection}: ${
          error instanceof Error ? error.message : String(error)
        }`,
      );
    }
  }

  private async loadThemeAssets(): Promise<void> {
    if (this.hasThemeAssetsLoaded || !this.x2tModule) return;
    this.hasThemeAssetsLoaded = true;

    const themeAssets = [
      'sdkjs/slide/themes/themes.js',
      ...Array.from({ length: 6 }, (_, index) => `sdkjs/slide/themes/theme${index + 1}/theme.bin`),
    ];

    await Promise.all(
      themeAssets.map(async (assetPath) => {
        try {
          const data = await fetchRuntimeBinaryAsset(assetPath);
          const relativePath = assetPath.replace(/^sdkjs\/slide\/themes\/?/, '');
          this.writeBinaryFile(`/working/themes/${relativePath}`, data);
        } catch (error) {
          throw new Error(
            `Failed to load OnlyOffice theme asset ${assetPath}: ${
              error instanceof Error ? error.message : String(error)
            }`,
          );
        }
      }),
    );
  }

  private unlinkIfExists(path: string): void {
    try {
      this.x2tModule?.FS.unlink(path);
    } catch {
      // The path may not exist, or may be a directory we intentionally keep.
    }
  }

  private clearDirectoryFiles(dir: string): void {
    const fs = this.x2tModule?.FS;
    if (!fs) return;

    try {
      fs.readdir(dir)
        .filter((file) => file !== '.' && file !== '..')
        .forEach((file) => {
          const path = `${dir}/${file}`;
          try {
            fs.readdir(path);
            this.clearDirectoryFiles(path);
          } catch {
            this.unlinkIfExists(path);
          }
        });
    } catch (error) {
      console.warn(`Failed to clean ${dir}:`, error);
    }
  }

  private clearConversionWorkspace(): void {
    const fs = this.x2tModule?.FS;
    if (!fs) return;

    try {
      fs.readdir('/working')
        .filter((file) => !['.', '..', 'fonts', 'media', 'themes'].includes(file))
        .forEach((file) => this.unlinkIfExists(`/working/${file}`));
    } catch (error) {
      console.warn('Failed to clean /working files:', error);
    }

    this.clearDirectoryFiles('/working/media');
    this.clearDirectoryFiles(CONVERTER_TEMP_DIR);
  }

  /**
   * Get document type
   */
  private getDocumentType(extension: string): DocumentType {
    const docType = DOCUMENT_TYPE_MAP[extension.toLowerCase()];
    if (!docType) {
      throw new Error(`Unsupported file format: ${extension}`);
    }
    return docType;
  }

  /**
   * Sanitize file name
   */
  private sanitizeFileName(input: string): string {
    if (typeof input !== 'string' || !input.trim()) {
      return 'file.bin';
    }

    const parts = input.split('.');
    const ext = parts.pop() || 'bin';
    const name = parts.join('.');

    const illegalChars = /[/?<>\\:*|"]/g;
    // eslint-disable-next-line no-control-regex
    const controlChars = /[\x00-\x1f\x80-\x9f]/g;
    const reservedPattern = /^\.+$/;
    const unsafeChars = /[&'%!"{}[\]]/g;

    let sanitized = name
      .replace(illegalChars, '')
      .replace(controlChars, '')
      .replace(reservedPattern, '')
      .replace(unsafeChars, '');

    sanitized = sanitized.trim() || 'file';
    return `${sanitized.slice(0, 200)}.${ext}`; // Limit length
  }

  /**
   * Execute document conversion
   */
  private executeConversion(paramsPath: string): void {
    if (!this.x2tModule) {
      throw new Error('X2T module not initialized');
    }

    const result = this.x2tModule.ccall('main1', 'number', ['string'], [paramsPath]);
    if (result !== 0) {
      // Read the params XML for debugging
      try {
        const paramsContent = this.x2tModule.FS.readFile(paramsPath, { encoding: 'binary' });
        // Convert binary to string for logging
        if (paramsContent instanceof Uint8Array) {
          const paramsText = new TextDecoder('utf-8').decode(paramsContent);
          console.error('Conversion failed. Parameters XML:', paramsText);
        } else {
          console.error('Conversion failed. Parameters XML:', paramsContent);
        }
      } catch (e) {
        console.error('Conversion failed. Parameters XML:', e);
        // Ignore if we can't read the params file
      }
      throw new Error(`Conversion failed with code: ${result}`);
    }
  }

  private readRequiredOutputFile(outputPath: string): BlobPart {
    const fs = this.x2tModule?.FS;
    if (!fs) {
      throw new Error('X2T module not initialized');
    }

    let result: BlobPart;
    try {
      result = fs.readFile(outputPath, { encoding: 'binary' });
    } catch {
      throw new Error(`Conversion completed without output file: ${outputPath}`);
    }

    const byteLength =
      typeof result === 'string'
        ? result.length
        : result instanceof Blob
          ? result.size
          : result instanceof ArrayBuffer
            ? result.byteLength
            : result.byteLength;
    if (byteLength === 0) {
      throw new Error(`Conversion produced an empty output file: ${outputPath}`);
    }
    return result;
  }

  /**
   * Create conversion parameters XML
   */
  private createConversionParams(fromPath: string, toPath: string, additionalParams = '', isNoBase64 = false): string {
    return `<?xml version="1.0" encoding="utf-8"?>
<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
  <m_sFileFrom>${fromPath}</m_sFileFrom>
  <m_sThemeDir>/working/themes</m_sThemeDir>
  <m_sAllFontsPath>${CONVERTER_ALL_FONTS_PATH}</m_sAllFontsPath>
  <m_sTempDir>${CONVERTER_TEMP_DIR}</m_sTempDir>
  <m_sFileTo>${toPath}</m_sFileTo>
  <m_bIsNoBase64>${isNoBase64 ? 'true' : 'false'}</m_bIsNoBase64>
  ${additionalParams}
</TaskQueueDataConvert>`;
  }

  private createDocumentToBinParams(
    fromPath: string,
    toPath: string,
    sourceExtension: string,
    documentType: DocumentType,
  ): string {
    const normalizedExtension = sourceExtension.toLowerCase();
    const sourceFormat = X2T_SOURCE_FORMAT_BY_EXTENSION[normalizedExtension];
    const targetFormat = X2T_CANVAS_FORMAT_BY_DOCUMENT_TYPE[documentType];
    const formatParams =
      sourceFormat && targetFormat
        ? `<m_nFormatFrom>${sourceFormat}</m_nFormatFrom>
  <m_nFormatTo>${targetFormat}</m_nFormatTo>
  <m_sFontDir>/working/fonts/</m_sFontDir>`
        : '';

    return this.createConversionParams(fromPath, toPath, formatParams);
  }

  private detectNativeBinFormat(bin: Uint8Array): number | undefined {
    if (bin.length < 4) return undefined;
    const signature = String.fromCharCode(bin[0], bin[1], bin[2], bin[3]);
    return X2T_NATIVE_FORMAT_BY_SIGNATURE[signature];
  }

  private detectNativeBinSourceFormat(bin: Uint8Array): number | undefined {
    return this.detectNativeBinFormat(bin) || this.detectNativeBinFormat(this.decodeNativeBase64Container(bin));
  }

  private isNativeNoBase64Container(bin: Uint8Array): boolean {
    if (bin.length < 12) return false;
    const header = String.fromCharCode(...bin.slice(0, Math.min(bin.length, 32)));
    return /^(DOCY|XLSY|PPTY);v\d+;0;/.test(header);
  }

  private decodeNativeBase64Container(bin: Uint8Array): Uint8Array {
    const text = decodeLatin1(bin).trim();
    if (!/^[A-Za-z0-9+/]+={0,2}$/.test(text)) return bin;

    try {
      const decoded = globalThis.atob(text);
      if (!/^(DOCY|XLSY|PPTY);v\d+;/.test(decoded)) return bin;
      return Uint8Array.from(decoded, (char) => char.charCodeAt(0) & 0xff);
    } catch {
      return bin;
    }
  }

  private createBinToDocumentParams(fromPath: string, toPath: string, bin: Uint8Array, targetExt: string): string {
    const sourceFormat = this.detectNativeBinSourceFormat(bin);
    const targetFormat = X2T_TARGET_FORMAT_BY_EXTENSION[targetExt.toLowerCase()];
    const isNativeBin = sourceFormat !== undefined;
    const formatParams = [
      sourceFormat !== undefined ? `<m_nFormatFrom>${sourceFormat}</m_nFormatFrom>` : '',
      targetFormat !== undefined ? `<m_nFormatTo>${targetFormat}</m_nFormatTo>` : '',
      isNativeBin || targetFormat === oAscFileType.PDF || targetFormat === oAscFileType.PDFA
        ? '<m_sFontDir>/working/fonts/</m_sFontDir>'
        : '',
    ]
      .filter(Boolean)
      .join('\n  ');

    return this.createConversionParams(fromPath, toPath, formatParams, this.isNativeNoBase64Container(bin));
  }

  /**
   * Read media files
   */
  private async readMediaFiles(): Promise<DocumentMediaMap> {
    if (!this.x2tModule) return {};

    const media: DocumentMediaMap = {};

    await this.readMediaDirectory('/working/media', 'media', media);

    return media;
  }

  private async readMediaDirectory(dir: string, keyPrefix: string, media: DocumentMediaMap): Promise<void> {
    if (!this.x2tModule) return;
    try {
      const files = this.x2tModule.FS.readdir(dir);
      await Promise.all(
        files
          .filter((file) => file !== '.' && file !== '..')
          .map(async (file) => {
            const path = `${dir}/${file}`;
            const key = `${keyPrefix}/${file}`;
            try {
              const fileData = this.x2tModule!.FS.readFile(path, {
                encoding: 'binary',
              }) as BlobPart;

              const blob = new Blob([fileData]);
              const mediaUrl = createObjectURL(blob);
              media[key] = mediaUrl;
            } catch (error) {
              try {
                await this.readMediaDirectory(path, key, media);
              } catch {
                console.warn(`Failed to read media file ${path}:`, error);
              }
            }
          }),
      );
    } catch (error) {
      console.warn(`Failed to read media directory ${dir}:`, error);
    }
  }

  private getWorkingMediaPath(key: string): string | null {
    const trimmed = safeDecodeURIComponent(key.replace(/^(\.\/|\/)+/, '')).replace(/\\/g, '/');
    const withoutWorkingPrefix = trimmed.startsWith('working/media/')
      ? trimmed.slice('working/media/'.length)
      : trimmed;
    const relativePath = withoutWorkingPrefix.startsWith('media/')
      ? withoutWorkingPrefix.slice('media/'.length)
      : withoutWorkingPrefix;
    const safeParts = relativePath.split('/').filter((part) => part && part !== '.' && part !== '..');
    if (safeParts.length === 0) return null;
    return `/working/media/${safeParts.join('/')}`;
  }

  private async writeMediaFiles(media?: DocumentMediaMap): Promise<void> {
    if (!media || Object.keys(media).length === 0) return;

    for (const [key, url] of Object.entries(media)) {
      const targetPath = this.getWorkingMediaPath(key);
      if (!targetPath) continue;
      if (!url) {
        throw new Error(`Missing media URL for ${key}`);
      }

      let response: Response;
      try {
        response = await fetch(url);
      } catch (error) {
        throw new Error(`Failed to fetch media ${key}: ${error instanceof Error ? error.message : String(error)}`);
      }
      if (!response.ok) {
        throw new Error(`Failed to fetch media ${key}: ${response.status} ${response.statusText}`);
      }

      const buffer = await response.arrayBuffer();
      const normalizedPath = targetPath.replace(/\\/g, '/');
      const dirPath = normalizedPath.slice(0, normalizedPath.lastIndexOf('/'));
      this.ensureDirectory(dirPath);
      this.x2tModule!.FS.writeFile(normalizedPath, new Uint8Array(buffer));
    }
  }

  /**
   * Load SheetJS for the demo CSV input adapter.
   *
   * Office Download As conversions must not use SheetJS or any other secondary
   * converter; they go through native OnlyOffice bytes and x2t-wasm only.
   */
  private async loadXlsxLibrary(): Promise<any> {
    // Check if xlsx is already loaded
    if (typeof window !== 'undefined' && (window as any).XLSX) {
      return (window as any).XLSX;
    }

    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = `${BASE_PATH}libs/sheetjs/xlsx.full.min.js`;
      script.onload = () => {
        if (typeof window !== 'undefined' && (window as any).XLSX) {
          resolve((window as any).XLSX);
        } else {
          reject(new Error('Failed to load xlsx library'));
        }
      };
      script.onerror = () => {
        reject(new Error('Failed to load xlsx library from local file'));
      };
      document.head.appendChild(script);
    });
  }

  /**
   * Convert standalone CSV input to XLSX before opening it in the spreadsheet editor.
   *
   * This is not part of the Office Download As/export pipeline.
   */
  private async convertCsvToXlsx(csvData: Uint8Array, fileName: string): Promise<File> {
    try {
      // Load xlsx library
      const XLSX = await this.loadXlsxLibrary();

      // Remove UTF-8 BOM if present
      let csvText: string;
      if (csvData.length >= 3 && csvData[0] === 0xef && csvData[1] === 0xbb && csvData[2] === 0xbf) {
        csvText = new TextDecoder('utf-8').decode(csvData.slice(3));
      } else {
        // Try UTF-8 first, fallback to other encodings if needed
        try {
          csvText = new TextDecoder('utf-8').decode(csvData);
        } catch {
          csvText = new TextDecoder('latin1').decode(csvData);
        }
      }

      // Parse CSV using SheetJS
      const workbook = XLSX.read(csvText, { type: 'string', raw: false });

      // Convert to XLSX binary format
      const xlsxBuffer = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });

      // Create File object
      const xlsxFileName = fileName.replace(/\.csv$/i, '.xlsx');
      return new File([xlsxBuffer], xlsxFileName, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    } catch (error) {
      throw new Error(
        `Failed to convert CSV to XLSX: ${error instanceof Error ? error.message : 'Unknown error'}. ` +
          'Please convert your CSV file to XLSX format manually and try again.',
      );
    }
  }

  /**
   * Convert document to bin format
   */
  async convertDocument(file: File): Promise<ConversionResult> {
    await this.initialize();

    const fileName = file.name;
    const fileExt = fileName.split('.').pop() || getExtensions(file?.type)[0] || '';
    const documentType = this.getDocumentType(fileExt);

    this.clearConversionWorkspace();
    try {
      // Read file content
      const arrayBuffer = await file.arrayBuffer();
      const data = new Uint8Array(arrayBuffer);

      // CSV is a standalone input adapter. Office Download As exports stay on x2t-wasm.
      if (fileExt.toLowerCase() === 'csv') {
        if (data.length === 0) {
          throw new Error('CSV file is empty');
        }
        console.log('CSV file detected. Converting to XLSX format...');
        console.log('CSV file size:', data.length, 'bytes');

        // Convert CSV to XLSX first
        try {
          const xlsxFile = await this.convertCsvToXlsx(data, fileName);
          console.log('CSV converted to XLSX, now converting with x2t...');

          // Now convert the XLSX file using x2t
          const xlsxArrayBuffer = await xlsxFile.arrayBuffer();
          const xlsxData = new Uint8Array(xlsxArrayBuffer);

          // Use the XLSX file for conversion
          const sanitizedName = this.sanitizeFileName(xlsxFile.name);
          const inputPath = `/working/${sanitizedName}`;
          const outputPath = `${inputPath}.bin`;

          // Write XLSX file to virtual file system
          this.x2tModule!.FS.writeFile(inputPath, xlsxData);

          // Create conversion parameters - no special params needed for XLSX
          const params = this.createDocumentToBinParams(inputPath, outputPath, 'xlsx', documentType);
          this.x2tModule!.FS.writeFile('/working/params.xml', params);

          // Execute conversion
          this.executeConversion('/working/params.xml');

          // Read conversion result
          const result = this.readRequiredOutputFile(outputPath);
          const media = await this.readMediaFiles();

          // Return original CSV fileName, not the XLSX one
          return {
            fileName: this.sanitizeFileName(fileName), // Keep original CSV filename
            type: documentType,
            bin: result,
            media,
          };
        } catch (conversionError: any) {
          // If conversion fails, provide helpful error message
          throw new Error(
            `Failed to convert CSV file: ${conversionError?.message || 'Unknown error'}. ` +
              'Please ensure your CSV file is properly formatted and try again.',
          );
        }
      }

      // For all other file types, use standard conversion
      const sanitizedName = this.sanitizeFileName(fileName);
      const inputPath = `/working/${sanitizedName}`;
      const outputPath = `${inputPath}.bin`;

      // Write file to virtual file system
      this.x2tModule!.FS.writeFile(inputPath, data);

      // Create conversion parameters - no special params needed for non-CSV files
      const params = this.createDocumentToBinParams(inputPath, outputPath, fileExt, documentType);
      this.x2tModule!.FS.writeFile('/working/params.xml', params);

      // Execute conversion
      this.executeConversion('/working/params.xml');

      // Read conversion result
      const result = this.readRequiredOutputFile(outputPath);
      const media = await this.readMediaFiles();

      return {
        fileName: sanitizedName,
        type: documentType,
        bin: result,
        media,
      };
    } catch (error) {
      throw new Error(`Document conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      this.clearConversionWorkspace();
    }
  }

  /**
   * Convert bin format to specified document format.
   */
  async convertBinToDocument(
    bin: Uint8Array,
    originalFileName: string,
    targetExt = 'DOCX',
    media?: DocumentMediaMap,
  ): Promise<BinConversionResult> {
    await this.initialize();

    const requestedTargetExt = targetExt.toLowerCase();
    const sourceBin = NATIVE_BASE64_PASSTHROUGH_TARGETS.has(requestedTargetExt)
      ? bin
      : this.decodeNativeBase64Container(bin);
    const normalizedTargetExt = LEGACY_TARGET_INTERMEDIATE_EXTENSION[requestedTargetExt] || requestedTargetExt;
    const outputExtension = X2T_OUTPUT_EXTENSION_BY_TARGET_EXTENSION[requestedTargetExt] || normalizedTargetExt;
    const sanitizedBase = this.sanitizeFileName(originalFileName).replace(/\.[^/.]+$/, '');
    const binFileName = `${sanitizedBase}.bin`;
    const outputFileName = `${sanitizedBase}.${outputExtension}`;

    this.clearConversionWorkspace();
    try {
      await this.writeMediaFiles(media);

      // Keep the entry point aligned with DocumentServer. Native DOCY/XLSY/PPTY
      // bins go directly to x2t with explicit format ids; x2t owns any internal
      // DOCX/XLSX/PPTX or HTML intermediate steps.
      this.x2tModule!.FS.writeFile(`/working/${binFileName}`, sourceBin);

      const params = this.createBinToDocumentParams(
        `/working/${binFileName}`,
        `/working/${outputFileName}`,
        sourceBin,
        normalizedTargetExt,
      );

      this.x2tModule!.FS.writeFile('/working/params.xml', params);
      this.executeConversion('/working/params.xml');

      // Read generated document
      const result = this.readRequiredOutputFile(`/working/${outputFileName}`);

      return {
        fileName: resolveGeneratedOutputFileName(outputFileName, result, normalizedTargetExt),
        data: result,
      };
    } catch (error) {
      throw new Error(`Bin to document conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      this.clearConversionWorkspace();
    }
  }

  async convertHtmlToDocument(
    htmlData: Uint8Array,
    originalFileName: string,
    targetExt: string,
  ): Promise<BinConversionResult> {
    await this.initialize();

    const normalizedTargetExt = targetExt.toLowerCase();
    const targetFormat = X2T_TARGET_FORMAT_BY_EXTENSION[normalizedTargetExt];
    if (targetFormat === undefined) {
      throw new Error(`Unsupported HTML export target: ${targetExt}`);
    }

    const outputExtension = X2T_OUTPUT_EXTENSION_BY_TARGET_EXTENSION[normalizedTargetExt] || normalizedTargetExt;
    const sanitizedBase = this.sanitizeFileName(originalFileName).replace(/\.[^/.]+$/, '');
    const inputFileName = `${sanitizedBase}.html`;
    const outputFileName = `${sanitizedBase}.${outputExtension}`;

    this.clearConversionWorkspace();
    try {
      this.x2tModule!.FS.writeFile(`/working/${inputFileName}`, htmlData);

      const params = this.createConversionParams(
        `/working/${inputFileName}`,
        `/working/${outputFileName}`,
        `<m_nFormatFrom>${oAscFileType.HTML}</m_nFormatFrom>
  <m_nFormatTo>${targetFormat}</m_nFormatTo>
  <m_sFontDir>/working/fonts/</m_sFontDir>`,
        true,
      );

      this.x2tModule!.FS.writeFile('/working/params.xml', params);
      this.executeConversion('/working/params.xml');

      const result = this.readRequiredOutputFile(`/working/${outputFileName}`);
      return {
        fileName: outputFileName,
        data: result,
      };
    } catch (error) {
      throw new Error(`HTML to document conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      this.clearConversionWorkspace();
    }
  }

  /**
   * Convert OnlyOffice native print renderer bytes into a PDF.
   *
   * The editor print API returns the raw renderer stream that native
   * DocumentServer/DesktopOffice passes to bin2pdf. It is not an Office canvas
   * save-bin and it is not base64 encoded, so it must use m_bIsNoBase64=true.
   */
  async convertPrintDataToPdf(
    printData: Uint8Array,
    originalFileName: string,
    media?: DocumentMediaMap,
  ): Promise<BinConversionResult> {
    await this.initialize();

    const sanitizedBase = this.sanitizeFileName(originalFileName).replace(/\.[^/.]+$/, '');
    const inputFileName = `${sanitizedBase}-print.bin`;
    const outputFileName = `${sanitizedBase}.pdf`;

    this.clearConversionWorkspace();
    try {
      await this.writeMediaFiles(media);
      this.x2tModule!.FS.writeFile(`/working/${inputFileName}`, printData);

      const params = this.createConversionParams(
        `/working/${inputFileName}`,
        `/working/${outputFileName}`,
        '<m_sFontDir>/working/fonts/</m_sFontDir>',
        true,
      );
      this.x2tModule!.FS.writeFile('/working/params.xml', params);
      this.executeConversion('/working/params.xml');

      const result = this.readRequiredOutputFile(`/working/${outputFileName}`);
      return {
        fileName: outputFileName,
        data: result,
      };
    } catch (error) {
      throw new Error(`Print PDF conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      this.clearConversionWorkspace();
    }
  }

  async convertPdfToImage(
    pdfData: Uint8Array,
    originalFileName: string,
    targetExt: string,
    options: { allPages?: boolean } = {},
  ): Promise<BinConversionResult> {
    await this.initialize();

    const normalizedTargetExt = targetExt.toLowerCase() === 'jpeg' ? 'jpg' : targetExt.toLowerCase();
    const targetFormat = X2T_TARGET_FORMAT_BY_EXTENSION[normalizedTargetExt];
    if (targetFormat !== oAscFileType.JPG && targetFormat !== oAscFileType.PNG) {
      throw new Error(`Unsupported PDF image export target: ${targetExt}`);
    }

    const sanitizedBase = this.sanitizeFileName(originalFileName).replace(/\.[^/.]+$/, '');
    const inputFileName = `${sanitizedBase}.pdf`;
    const outputFileName = `${sanitizedBase}.${normalizedTargetExt}`;
    const rasterFormat = targetFormat === oAscFileType.JPG ? 3 : 4;
    const exportAllPages = options.allPages === true;

    this.clearConversionWorkspace();
    try {
      this.x2tModule!.FS.writeFile(`/working/${inputFileName}`, pdfData);

      const params = this.createConversionParams(
        `/working/${inputFileName}`,
        `/working/${outputFileName}`,
        `<m_nFormatFrom>${oAscFileType.PDF}</m_nFormatFrom>
  <m_nFormatTo>${targetFormat}</m_nFormatTo>
  <m_sFontDir>/working/fonts/</m_sFontDir>
  <m_oThumbnail>
    <format>${rasterFormat}</format>
    <aspect>2</aspect>
    <first>${exportAllPages ? 'false' : 'true'}</first>
    <zip>${exportAllPages ? 'true' : 'false'}</zip>
  </m_oThumbnail>`,
        true,
      );
      this.x2tModule!.FS.writeFile('/working/params.xml', params);
      this.executeConversion('/working/params.xml');

      const result = this.readRequiredOutputFile(`/working/${outputFileName}`);
      return {
        fileName: resolveGeneratedOutputFileName(outputFileName, result, normalizedTargetExt, originalFileName),
        data: result,
      };
    } catch (error) {
      throw new Error(`PDF to image conversion failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      this.clearConversionWorkspace();
    }
  }

  /**
   * Convert bin format to specified format and save it locally.
   */
  async convertBinToDocumentAndDownload(
    bin: Uint8Array,
    originalFileName: string,
    targetExt = 'DOCX',
    media?: DocumentMediaMap,
  ): Promise<BinConversionResult> {
    const result = await this.convertBinToDocument(bin, originalFileName, targetExt, media);
    const data = result.data instanceof Uint8Array ? result.data : new Uint8Array(result.data as ArrayBuffer);

    // TODO: Improve print functionality
    await this.saveWithFileSystemAPI(data, result.fileName);
    return result;
  }

  /**
   * Download file
   */
  private async downloadFile(data: Uint8Array, fileName: string): Promise<void> {
    const blob = new Blob([data as BlobPart]);
    const url = createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = fileName;
    link.style.display = 'none';

    document.body.appendChild(link);
    link.click();

    // Clean up resources
    setTimeout(() => {
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    }, 100);
  }

  /**
   * Get MIME type from file extension
   */
  private getMimeTypeFromExtension(extension: string): string {
    const mimeMap: Record<string, string> = {
      // Document types
      docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      doc: 'application/msword',
      odt: 'application/vnd.oasis.opendocument.text',
      rtf: 'application/rtf',
      txt: 'text/plain',
      pdf: 'application/pdf',

      // Spreadsheet types
      xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      xls: 'application/vnd.ms-excel',
      ods: 'application/vnd.oasis.opendocument.spreadsheet',
      csv: 'text/csv',

      // Presentation types
      pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      ppt: 'application/vnd.ms-powerpoint',
      odp: 'application/vnd.oasis.opendocument.presentation',

      // Image types
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      bmp: 'image/bmp',
      webp: 'image/webp',
      svg: 'image/svg+xml',
    };

    return mimeMap[extension.toLowerCase()] || 'application/octet-stream';
  }

  /**
   * Get file type description
   */
  private getFileDescription(extension: string): string {
    const descriptionMap: Record<string, string> = {
      docx: 'Word Document',
      doc: 'Word 97-2003 Document',
      odt: 'OpenDocument Text',
      pdf: 'PDF Document',
      xlsx: 'Excel Workbook',
      xls: 'Excel 97-2003 Workbook',
      ods: 'OpenDocument Spreadsheet',
      pptx: 'PowerPoint Presentation',
      ppt: 'PowerPoint 97-2003 Presentation',
      odp: 'OpenDocument Presentation',
      txt: 'Text Document',
      rtf: 'Rich Text Format',
      csv: 'CSV File',
    };

    return descriptionMap[extension.toLowerCase()] || 'Document';
  }

  /**
   * Save file using modern File System API
   */
  private async saveWithFileSystemAPI(data: Uint8Array, fileName: string, mimeType?: string): Promise<void> {
    if (!(window as any).showSaveFilePicker) {
      await this.downloadFile(data, fileName);
      return;
    }
    try {
      // Get file extension and determine MIME type
      const extension = fileName.split('.').pop()?.toLowerCase() || '';
      const detectedMimeType = mimeType || this.getMimeTypeFromExtension(extension);

      // Show file save dialog
      const fileHandle = await (window as any).showSaveFilePicker({
        suggestedName: fileName,
        types: [
          {
            description: this.getFileDescription(extension),
            accept: {
              [detectedMimeType]: [`.${extension}`],
            },
          },
        ],
      });

      // Create writable stream and write data
      const writable = await fileHandle.createWritable();
      await writable.write(data);
      await writable.close();
      console.log('File saved successfully:', fileName);
    } catch (error) {
      if ((error as Error).name === 'AbortError') {
        console.log('User cancelled the save operation');
        return;
      }
      throw error;
    }
  }

  /**
   * Destroy instance and clean up resources
   */
  destroy(): void {
    this.x2tModule = null;
    this.isReady = false;
    this.initPromise = null;
    console.log('X2T converter destroyed');
  }
}
