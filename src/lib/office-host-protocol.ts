export const OFFICE_HOST_PROTOCOL = 'onlyoffice-browser-host/v1';

export type OfficeHostSourceKind = 'local-file' | 'new-document' | 'buffer' | 'url';
export type OfficeHostSaveBehavior = 'auto' | 'callback' | 'download';
export type OfficeHostInterfaceTheme = 'system' | 'light' | 'dark';

export interface OfficeHostIdentity {
  packageVersion: string;
  hostBuildId: string;
  assetManifestDigest: string;
}

export type OfficeSaveToNewFormatConfirmationOptions = {
  title?: string;
  message?: string;
  dontshow?: boolean;
};

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
      sourceKind: Exclude<OfficeHostSourceKind, 'new-document'>;
    };

export interface OfficeHostInitOptions {
  fileName?: string;
  mode?: 'edit' | 'readonly' | 'preview';
  readonly?: boolean;
  canReturnToPreview?: boolean;
  spellcheck?: boolean;
  interfaceTheme?: OfficeHostInterfaceTheme;
  lang?: string;
  saveBehavior?: OfficeHostSaveBehavior;
  source: OfficeHostSource;
}

export interface OfficeHostState {
  id: string;
  fileName: string;
  fileType: string;
  mode: 'edit' | 'readonly' | 'preview';
  readonly: boolean;
  dirty: boolean;
  sourceKind: OfficeHostSourceKind;
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
      identity: OfficeHostIdentity;
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
      type: 'SAVE_ACK';
      requestId: string;
      ok: boolean;
      message?: string;
    })
  | (OfficeHostBaseMessage & {
      type: 'CONFIRM_SAVE_TO_NEW_FORMAT';
      requestId: string;
      options?: OfficeSaveToNewFormatConfirmationOptions;
    })
  | (OfficeHostBaseMessage & {
      type: 'SET_READONLY';
      readonly: boolean;
    })
  | (OfficeHostBaseMessage & {
      type: 'SET_INTERFACE_THEME';
      interfaceTheme: OfficeHostInterfaceTheme;
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
      type: 'DOWNLOAD_RESULT';
      buffer: ArrayBuffer;
      fileName: string;
      mimeType: string;
    })
  | (OfficeHostBaseMessage & {
      type: 'SAVE_AS_RESULT';
      buffer: ArrayBuffer;
      fileName: string;
      mimeType: string;
    })
  | (OfficeHostBaseMessage & {
      type: 'PRINT_TITLE';
      title: string;
      durationMs: number;
    })
  | (OfficeHostBaseMessage & {
      type: 'CONFIRM_SAVE_TO_NEW_FORMAT_RESULT';
      requestId: string;
      confirmed: boolean;
    })
  | (OfficeHostBaseMessage & {
      type: 'ERROR';
      phase: 'handshake' | 'init' | 'save' | 'confirm' | 'setReadonly' | 'setInterfaceTheme' | 'destroy' | 'runtime';
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
