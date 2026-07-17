/// <reference lib="webworker" />

import { X2TConverter } from './lib/document-converter';
import {
  CONVERSION_WORKER_PROTOCOL,
  type ConversionWorkerRequest,
  type ConversionWorkerResponse,
  type SerializedBinConversionResult,
  type SerializedConversionResult,
  type SerializedMediaMap,
} from './lib/conversion-worker-protocol';
import type { BinConversionResult, ConversionResult, DocumentMediaMap } from './lib/document-types';

const worker = self as DedicatedWorkerGlobalScope;

async function blobPartToArrayBuffer(value: BlobPart): Promise<ArrayBuffer> {
  return new Blob([value]).arrayBuffer();
}

async function serializeMedia(media: DocumentMediaMap): Promise<SerializedMediaMap> {
  const serialized: SerializedMediaMap = {};
  await Promise.all(
    Object.entries(media).map(async ([key, url]) => {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Failed to serialize converted media ${key}: ${response.status}`);
      serialized[key] = {
        data: await response.arrayBuffer(),
        type: response.headers.get('Content-Type') || 'application/octet-stream',
      };
      URL.revokeObjectURL(url);
    }),
  );
  return serialized;
}

function hydrateMedia(media: SerializedMediaMap): { media: DocumentMediaMap; revoke: () => void } {
  const hydrated: DocumentMediaMap = {};
  const urls: string[] = [];
  for (const [key, entry] of Object.entries(media)) {
    const url = URL.createObjectURL(new Blob([entry.data], { type: entry.type }));
    hydrated[key] = url;
    urls.push(url);
  }
  return { media: hydrated, revoke: () => urls.forEach((url) => URL.revokeObjectURL(url)) };
}

async function serializeDocumentResult(result: ConversionResult): Promise<SerializedConversionResult> {
  return {
    kind: 'document',
    fileName: result.fileName,
    type: result.type,
    bin: await blobPartToArrayBuffer(result.bin),
    media: await serializeMedia(result.media),
  };
}

async function serializeBinResult(result: BinConversionResult): Promise<SerializedBinConversionResult> {
  return {
    kind: 'bin',
    fileName: result.fileName,
    data: await blobPartToArrayBuffer(result.data),
  };
}

function collectTransferables(result: SerializedConversionResult | SerializedBinConversionResult): Transferable[] {
  if (result.kind === 'bin') return [result.data];
  return [result.bin, ...Object.values(result.media).map((entry) => entry.data)];
}

async function execute(request: ConversionWorkerRequest) {
  const converter = new X2TConverter();
  const { operation } = request;
  switch (operation.kind) {
    case 'convert-document':
      return serializeDocumentResult(await converter.convertDocument(operation.file));
    case 'convert-bin': {
      const hydrated = hydrateMedia(operation.media);
      try {
        return serializeBinResult(
          await converter.convertBinToDocument(
            new Uint8Array(operation.bin),
            operation.fileName,
            operation.targetExt,
            hydrated.media,
          ),
        );
      } finally {
        hydrated.revoke();
      }
    }
    case 'convert-print': {
      const hydrated = hydrateMedia(operation.media);
      try {
        return serializeBinResult(
          await converter.convertPrintDataToPdf(
            new Uint8Array(operation.printData),
            operation.fileName,
            hydrated.media,
          ),
        );
      } finally {
        hydrated.revoke();
      }
    }
    case 'convert-html':
      return serializeBinResult(
        await converter.convertHtmlToDocument(
          new Uint8Array(operation.htmlData),
          operation.fileName,
          operation.targetExt,
        ),
      );
    case 'convert-pdf-image':
      return serializeBinResult(
        await converter.convertPdfToImage(
          new Uint8Array(operation.pdfData),
          operation.fileName,
          operation.targetExt,
          operation.options,
        ),
      );
  }
}

worker.addEventListener('message', (event: MessageEvent<ConversionWorkerRequest>) => {
  const request = event.data;
  if (!request || request.protocol !== CONVERSION_WORKER_PROTOCOL) return;
  void execute(request)
    .then((result) => {
      const response: ConversionWorkerResponse = {
        protocol: CONVERSION_WORKER_PROTOCOL,
        id: request.id,
        ok: true,
        result,
      };
      worker.postMessage(response, collectTransferables(result));
    })
    .catch((error: unknown) => {
      const normalized = error instanceof Error ? error : new Error(String(error));
      const response: ConversionWorkerResponse = {
        protocol: CONVERSION_WORKER_PROTOCOL,
        id: request.id,
        ok: false,
        error: { name: normalized.name, message: normalized.message, stack: normalized.stack },
      };
      worker.postMessage(response);
    });
});
