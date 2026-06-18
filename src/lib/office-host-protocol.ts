export const OFFICE_HOST_PROTOCOL = 'onlyoffice-browser-host/v1';

export type OfficeHostSource =
  | {
      kind: 'empty';
      emptyType: 'docx' | 'xlsx' | 'pptx' | 'csv';
    }
  | {
      kind: 'buffer';
      buffer: ArrayBuffer;
      fileName: string;
      mimeType: string;
    };

export interface OfficeHostInitOptions {
  fileName?: string;
  mode?: 'edit' | 'readonly' | 'preview';
  readonly?: boolean;
  lang?: string;
  source: OfficeHostSource;
}

export interface OfficeHostState {
  id: string;
  fileName: string;
  fileType: string;
  mode: 'edit' | 'readonly' | 'preview';
  readonly: boolean;
  status: 'opening' | 'ready' | 'destroyed' | 'error';
  destroyed: boolean;
}

export interface OfficeHostBaseMessage {
  protocol: typeof OFFICE_HOST_PROTOCOL;
  sessionId: string;
  requestId?: string;
}

export type OfficeHostWindowMessage =
  | (OfficeHostBaseMessage & {
      type: 'HOST_READY';
    })
  | (OfficeHostBaseMessage & {
      type: 'HOST_RESET_DONE';
    })
  | (OfficeHostBaseMessage & {
      type: 'CONNECT';
    });

export type OfficeHostParentMessage =
  | (OfficeHostBaseMessage & {
      type: 'INIT';
      options: OfficeHostInitOptions;
    })
  | (OfficeHostBaseMessage & {
      type: 'SAVE';
      targetExt?: string;
    })
  | (OfficeHostBaseMessage & {
      type: 'SET_READONLY';
      readonly: boolean;
    })
  | (OfficeHostBaseMessage & {
      type: 'DESTROY';
    });

export type OfficeHostChildMessage =
  | (OfficeHostBaseMessage & {
      type: 'READY';
      state: OfficeHostState;
    })
  | (OfficeHostBaseMessage & {
      type: 'STATE';
      state: OfficeHostState;
    })
  | (OfficeHostBaseMessage & {
      type: 'SAVE_RESULT';
      buffer: ArrayBuffer;
      fileName: string;
      mimeType: string;
    })
  | (OfficeHostBaseMessage & {
      type: 'ERROR';
      phase: 'handshake' | 'init' | 'save' | 'setReadonly' | 'destroy' | 'runtime';
      message: string;
    })
  | (OfficeHostBaseMessage & {
      type: 'DESTROYED';
    });

export function isOfficeHostMessage(value: unknown, sessionId: string): value is OfficeHostBaseMessage {
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as OfficeHostBaseMessage).protocol === OFFICE_HOST_PROTOCOL &&
    (value as OfficeHostBaseMessage).sessionId === sessionId
  );
}
