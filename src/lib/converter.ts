import {
  CONVERSION_WORKER_PROTOCOL,
  type ConversionWorkerOperation,
  type ConversionWorkerRequest,
  type ConversionWorkerResponse,
  type SerializedBinConversionResult,
  type SerializedConversionResult,
  type SerializedMediaMap,
} from './conversion-worker-protocol';
import type { BinConversionResult, ConversionResult, DocumentMediaMap } from './document-types';

export type {
  ConversionResult,
  BinConversionResult,
  DocumentMediaMap,
  EmscriptenModule,
  DocumentType,
  SaveEvent,
} from './document-types';

export { oAscFileType, c_oAscFileType2 } from './file-types';
export { getDocumentType, getBasePath, BASE_PATH, DOCUMENT_TYPE_MAP } from './document-utils';

const CONVERSION_TIMEOUT_MS = 5 * 60_000;
let conversionQueue: Promise<unknown> = Promise.resolve();
let nextConversionId = 1;

function queueConversion<T>(operation: () => Promise<T>): Promise<T> {
  const next = conversionQueue.then(operation, operation);
  conversionQueue = next.catch(() => undefined);
  return next;
}

function copyBuffer(value: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(value.byteLength);
  copy.set(value);
  return copy.buffer;
}

async function serializeMedia(media?: DocumentMediaMap): Promise<SerializedMediaMap> {
  if (!media) return {};
  const serialized: SerializedMediaMap = {};
  await Promise.all(
    Object.entries(media).map(async ([key, url]) => {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Failed to read editor media ${key}: ${response.status}`);
      serialized[key] = {
        data: await response.arrayBuffer(),
        type: response.headers.get('Content-Type') || 'application/octet-stream',
      };
    }),
  );
  return serialized;
}

function collectOperationTransferables(operation: ConversionWorkerOperation): Transferable[] {
  switch (operation.kind) {
    case 'convert-document':
      return [];
    case 'convert-bin':
      return [operation.bin, ...Object.values(operation.media).map((entry) => entry.data)];
    case 'convert-print':
      return [operation.printData, ...Object.values(operation.media).map((entry) => entry.data)];
    case 'convert-html':
      return [operation.htmlData];
    case 'convert-pdf-image':
      return [operation.pdfData];
  }
}

function runConversionWorker(
  operation: ConversionWorkerOperation,
): Promise<SerializedConversionResult | SerializedBinConversionResult> {
  const id = `conversion-${nextConversionId++}-${crypto.randomUUID?.() || Date.now()}`;
  // x2t.js is the classic Emscripten runtime and is loaded with importScripts
  // inside conversion-worker.ts. Module workers reject importScripts, while
  // Vite emits this dependency as a self-contained IIFE suitable for a
  // classic worker.
  const worker = new Worker(new URL('../conversion-worker.ts', import.meta.url));
  const request: ConversionWorkerRequest = { protocol: CONVERSION_WORKER_PROTOCOL, id, operation };

  return new Promise((resolve, reject) => {
    const timeoutId = window.setTimeout(() => {
      worker.terminate();
      reject(new Error(`Office conversion timed out after ${CONVERSION_TIMEOUT_MS}ms`));
    }, CONVERSION_TIMEOUT_MS);

    const cleanup = () => {
      window.clearTimeout(timeoutId);
      worker.terminate();
    };

    worker.addEventListener('message', (event: MessageEvent<ConversionWorkerResponse>) => {
      const response = event.data;
      if (!response || response.protocol !== CONVERSION_WORKER_PROTOCOL || response.id !== id) return;
      cleanup();
      if (response.ok) {
        resolve(response.result);
        return;
      }
      const error = new Error(response.error.message);
      error.name = response.error.name;
      if (response.error.stack) error.stack = response.error.stack;
      reject(error);
    });
    worker.addEventListener('error', (event) => {
      cleanup();
      reject(new Error(event.message || 'Office conversion worker failed'));
    });
    worker.postMessage(request, collectOperationTransferables(operation));
  });
}

function hydrateDocumentResult(result: SerializedConversionResult): ConversionResult {
  const media: DocumentMediaMap = {};
  for (const [key, entry] of Object.entries(result.media)) {
    media[key] = URL.createObjectURL(new Blob([entry.data], { type: entry.type }));
  }
  return {
    fileName: result.fileName,
    type: result.type,
    bin: new Uint8Array(result.bin),
    media,
  };
}

function hydrateBinResult(result: SerializedBinConversionResult): BinConversionResult {
  return { fileName: result.fileName, data: new Uint8Array(result.data) };
}

async function expectDocumentResult(operation: ConversionWorkerOperation): Promise<ConversionResult> {
  const result = await runConversionWorker(operation);
  if (result.kind !== 'document') throw new Error('Office conversion worker returned an unexpected result');
  return hydrateDocumentResult(result);
}

async function expectBinResult(operation: ConversionWorkerOperation): Promise<BinConversionResult> {
  const result = await runConversionWorker(operation);
  if (result.kind !== 'bin') throw new Error('Office conversion worker returned an unexpected result');
  return hydrateBinResult(result);
}

// Conversion workers initialize x2t lazily for each operation and terminate
// immediately afterwards so READY editors do not retain a WASM heap.
export const loadScript = (): Promise<void> => Promise.resolve();
export const initX2T = (): Promise<void> => Promise.resolve();

export const convertDocument = (file: File): Promise<ConversionResult> =>
  queueConversion(() => expectDocumentResult({ kind: 'convert-document', file }));

export const convertBinToDocument = (
  bin: Uint8Array,
  fileName: string,
  targetExt?: string,
  media?: DocumentMediaMap,
): Promise<BinConversionResult> =>
  queueConversion(async () =>
    expectBinResult({
      kind: 'convert-bin',
      bin: copyBuffer(bin),
      fileName,
      targetExt,
      media: await serializeMedia(media),
    }),
  );

export const convertPrintDataToPdf = (
  printData: Uint8Array,
  fileName: string,
  media?: DocumentMediaMap,
): Promise<BinConversionResult> =>
  queueConversion(async () =>
    expectBinResult({
      kind: 'convert-print',
      printData: copyBuffer(printData),
      fileName,
      media: await serializeMedia(media),
    }),
  );

export const convertHtmlToDocument = (
  htmlData: Uint8Array,
  fileName: string,
  targetExt: string,
): Promise<BinConversionResult> =>
  queueConversion(() => expectBinResult({ kind: 'convert-html', htmlData: copyBuffer(htmlData), fileName, targetExt }));

export const convertPdfToImage = (
  pdfData: Uint8Array,
  fileName: string,
  targetExt: string,
  options?: { allPages?: boolean },
): Promise<BinConversionResult> =>
  queueConversion(() =>
    expectBinResult({
      kind: 'convert-pdf-image',
      pdfData: copyBuffer(pdfData),
      fileName,
      targetExt,
      options,
    }),
  );

export const convertBinToDocumentAndDownload = async (
  bin: Uint8Array,
  fileName: string,
  targetExt?: string,
  media?: DocumentMediaMap,
): Promise<BinConversionResult> => {
  const result = await convertBinToDocument(bin, fileName, targetExt, media);
  const url = URL.createObjectURL(new Blob([result.data]));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = result.fileName;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
  return result;
};
