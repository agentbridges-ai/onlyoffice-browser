import type { DocumentType } from './document-types';

export const CONVERSION_WORKER_PROTOCOL = 'onlyoffice-browser-conversion-v1' as const;

export type SerializedMediaEntry = {
  data: ArrayBuffer;
  type: string;
};

export type SerializedMediaMap = Record<string, SerializedMediaEntry>;

export type ConversionWorkerOperation =
  | { kind: 'convert-document'; file: File }
  | {
      kind: 'convert-bin';
      bin: ArrayBuffer;
      fileName: string;
      targetExt?: string;
      media: SerializedMediaMap;
    }
  | {
      kind: 'convert-print';
      printData: ArrayBuffer;
      fileName: string;
      media: SerializedMediaMap;
    }
  | { kind: 'convert-html'; htmlData: ArrayBuffer; fileName: string; targetExt: string }
  | {
      kind: 'convert-pdf-image';
      pdfData: ArrayBuffer;
      fileName: string;
      targetExt: string;
      options?: { allPages?: boolean };
    };

export type ConversionWorkerRequest = {
  protocol: typeof CONVERSION_WORKER_PROTOCOL;
  id: string;
  operation: ConversionWorkerOperation;
};

export type SerializedConversionResult = {
  kind: 'document';
  fileName: string;
  type: DocumentType;
  bin: ArrayBuffer;
  media: SerializedMediaMap;
};

export type SerializedBinConversionResult = {
  kind: 'bin';
  fileName: string;
  data: ArrayBuffer;
};

export type ConversionWorkerResponse =
  | {
      protocol: typeof CONVERSION_WORKER_PROTOCOL;
      id: string;
      ok: true;
      result: SerializedConversionResult | SerializedBinConversionResult;
    }
  | {
      protocol: typeof CONVERSION_WORKER_PROTOCOL;
      id: string;
      ok: false;
      error: { name: string; message: string; stack?: string };
    };
