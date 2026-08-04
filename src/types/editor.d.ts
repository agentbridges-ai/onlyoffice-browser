interface DocEditorConfig {
  height?: string | number;
  type?: 'desktop' | 'mobile' | 'embedded';
  width?: string | number;
  documentType?: string;
  document: {
    title: string;
    url: string;
    fileType: string;
    key?: string;
    permissions: {
      edit: boolean;
      chat: boolean;
      protect: boolean;
      download?: boolean;
    };
  };
  editorConfig: {
    lang: string;
    mode?: 'edit' | 'view';
    coEditing?: {
      mode?: 'fast' | 'strict';
      change?: boolean;
    };
    user?: {
      id: string;
      name: string;
    };
    embedded?: {
      autostart?: 'document' | 'player';
      toolbarDocked?: 'top' | 'bottom';
      embedUrl?: string;
      fullscreenUrl?: string;
      saveUrl?: string;
      shareUrl?: string;
    };
    customization: {
      help: boolean;
      about: boolean;
      hideRightMenu: boolean;
      /** Start with the native OnlyOffice ribbon collapsed. */
      compactToolbar?: boolean;
      /** OnlyOffice interface theme. Use modern values: theme-system, theme-white, or theme-night. */
      uiTheme?: 'theme-system' | 'theme-white' | 'theme-night';
      /** OnlyOffice zoom preset. -2 maps to native fit-to-width. */
      zoom?: number;
      /** Whether spell checking is enabled by default. */
      spellcheck?: boolean;
      /** Whether the native OnlyOffice autosave option is enabled. */
      autosave?: boolean;
      /** Whether the native force-save-on-user-save option is enabled. */
      forcesave?: boolean;
      /** Enable/disable plugins. Set to false to disable plugins */
      plugins?: boolean;
      features: {
        /** Disable built-in new-feature coach marks. */
        featuresTips?: boolean;
        spellcheck: {
          change: boolean;
        };
      };
      anonymous: {
        request: boolean;
        label: string;
      };
    };
    /** Plugin configuration. Can specify a list of plugins to load */
    plugins?: {
      pluginsData?: string[];
      autostart?: string[];
    };
  };
  events: {
    onAppReady: () => void;
    onDocumentReady: () => void;
    onSave: (event: SaveEvent) => void;
    onDocumentStateChange?: (event: DocumentStateChangeEvent) => void;
    onDownloadAs?: (event: DownloadAsEvent) => void;
    onRequestSaveAs?: (event: DownloadAsEvent) => void;
    onRequestEditRights?: () => void;
    writeFile: (event: WriteFileEvent) => void;
    /** Handle external messages from plugins */
    onExternalPluginMessage?: (event: { type: string; data: any; pluginName?: string }) => void;
  };
}

interface SaveEvent {
  data: {
    data: {
      data: ArrayBuffer;
    };
    option: {
      outputformat: number;
    };
  };
}

interface WriteFileEvent {
  data: {
    data: Uint8Array;
    file: string;
    target: {
      frameOrigin: string;
    };
  };
  callback?: (result: { success: boolean; error?: string }) => void;
}

interface DownloadAsEvent {
  data?: {
    url?: string;
    fileType?: string | number;
    title?: string;
  };
}

type DocumentStateChangeEvent =
  | boolean
  | {
      data?: boolean;
    };

interface OnlyOfficeMockServer {
  buildNumber?: number;
  buildVersion?: string;
  getInitialChanges?: () => any[];
  getParticipants: () => {
    index: number;
    list: [
      {
        id: number;
        idOriginal: string;
        username: string;
        indexUser: number;
        connectionId: string;
        isCloseCoAuthoring: boolean;
        view: boolean;
      },
    ];
  };
  getDocumentOpenData?: (documentUrl: string) => Record<string, string>;
  getImageURL?: (name: string) => Promise<string>;
  onAuth?: () => void;
  handleMessage?: (msg: any, respond: (response: any) => void) => boolean;
  onMessage: (msg: any) => void;
  onCorruptionWarning?: (duplicateId: string) => void;
}

interface DocEditor {
  sendCommand?: (params: {
    command: string;
    data: Record<string, unknown> & {
      err_code?: number;
      urls?: Record<string, string>;
      path?: string;
      imgName?: string;
      buf?: ArrayBuffer;
      success?: boolean;
      error?: string;
      enabled?: boolean;
      message?: string;
    };
  }) => void;
  openDocument?: (data: Uint8Array) => void;
  downloadAs?: (data?: string) => void;
  getEditorApi?: () => {
    asc_DownloadAs?: (options?: unknown) => unknown;
    _downloadAs?: (
      actionType?: unknown,
      options?: unknown,
      additionalData?: Record<string, unknown>,
      dataContainer?: Record<string, unknown>,
      downloadType?: unknown,
    ) => boolean | undefined;
    [key: string]: unknown;
  } | null;
  getEditorWindow?: () => Window | null;
  asc_nativeGetFile3?: () =>
    | {
        data?: Uint8Array | ArrayBuffer | ArrayBufferView;
        header?: string;
      }
    | Uint8Array
    | ArrayBuffer
    | ArrayBufferView;
  asc_nativeGetPDF?: (options?: Record<string, unknown>) =>
    | {
        data?: Uint8Array | ArrayBuffer | ArrayBufferView;
      }
    | Uint8Array
    | ArrayBuffer
    | ArrayBufferView
    | null;
  asc_nativeCalculateFile?: (options?: Record<string, unknown> | null) => unknown;
  asc_nativeGetHtml?: (options?: Record<string, unknown> | null) => string;
  zoomFitToWidth?: () => void;
  processRightsChange?: (enabled: boolean, message?: string) => void;
  connectMockServer?: (server: OnlyOfficeMockServer) => void;
  cryptPadMessageToOO?: (msg: any) => void;
  sendMessageToOO?: (msg: any) => void;
  serviceCommand?: (command: string, data: any) => void;
  waitForAppReady?: Promise<void>;
  destroyEditor: () => void;
}

interface DocsAPI {
  DocEditor: new (elementId: string, config: DocEditorConfig) => DocEditor;
}

declare global {
  interface Window {
    DocsAPI?: DocsAPI;
    APP?: Record<string, unknown>;
  }
}
