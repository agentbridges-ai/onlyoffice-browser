import { X2TConverter } from './document-converter';
import type { BinConversionResult, ConversionResult, DocumentMediaMap, EmscriptenModule } from './document-types';

// Export types
export type {
  ConversionResult,
  BinConversionResult,
  DocumentMediaMap,
  EmscriptenModule,
  DocumentType,
  SaveEvent,
} from './document-types';

// Export constants
export { oAscFileType, c_oAscFileType2 } from './file-types';

// Export utilities
export { getDocumentType, getBasePath, BASE_PATH, DOCUMENT_TYPE_MAP } from './document-utils';

// Singleton instance
const x2tConverter = new X2TConverter();
let conversionQueue: Promise<unknown> = Promise.resolve();

function queueConversion<T>(operation: () => Promise<T>): Promise<T> {
  const next = conversionQueue.then(operation, operation);
  conversionQueue = next.catch(() => {});
  return next;
}

// Export converter methods
export const loadScript = (): Promise<void> => x2tConverter.loadScript();
export const initX2T = (): Promise<EmscriptenModule> => x2tConverter.initialize();
export const convertDocument = (file: File): Promise<ConversionResult> =>
  queueConversion(() => x2tConverter.convertDocument(file));
export const convertBinToDocument = (
  bin: Uint8Array,
  fileName: string,
  targetExt?: string,
  media?: DocumentMediaMap,
): Promise<BinConversionResult> =>
  queueConversion(() => x2tConverter.convertBinToDocument(bin, fileName, targetExt, media));
export const convertPrintDataToPdf = (
  printData: Uint8Array,
  fileName: string,
  media?: DocumentMediaMap,
): Promise<BinConversionResult> =>
  queueConversion(() => x2tConverter.convertPrintDataToPdf(printData, fileName, media));
export const convertBinToDocumentAndDownload = (
  bin: Uint8Array,
  fileName: string,
  targetExt?: string,
  media?: DocumentMediaMap,
): Promise<BinConversionResult> =>
  queueConversion(() => x2tConverter.convertBinToDocumentAndDownload(bin, fileName, targetExt, media));
