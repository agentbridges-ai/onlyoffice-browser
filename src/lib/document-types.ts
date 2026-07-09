// Type definitions for x2t module

export interface EmscriptenFileSystem {
  mkdir(path: string): void;
  readdir(path: string): string[];
  readFile(path: string, options?: { encoding: 'binary' }): BlobPart;
  writeFile(path: string, data: Uint8Array | string): void;
  unlink(path: string): void;
}

export interface EmscriptenModule {
  FS: EmscriptenFileSystem;
  ccall: (funcName: string, returnType: string, argTypes: string[], args: any[]) => number;
  onRuntimeInitialized: () => void;
}

export interface ConversionResult {
  fileName: string;
  type: DocumentType;
  bin: BlobPart;
  media: DocumentMediaMap;
}

export interface BinConversionResult {
  fileName: string;
  data: BlobPart;
}

export type DocumentMediaMap = Record<string, string>;

export type DocumentType = 'word' | 'cell' | 'slide';

export interface SaveEvent {
  data: {
    data: {
      data: Uint8Array | ArrayBuffer;
    };
    option: {
      outputformat: number;
    };
  };
}

declare global {
  interface Window {
    Module: EmscriptenModule;
    APP?: Record<string, unknown>;
  }
}
