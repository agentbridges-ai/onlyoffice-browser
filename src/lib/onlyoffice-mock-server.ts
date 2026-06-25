export interface OnlyOfficeParticipant {
  id: number;
  idOriginal: string;
  username: string;
  indexUser: number;
  connectionId: string;
  isCloseCoAuthoring: boolean;
  view: boolean;
}

export interface OnlyOfficeParticipants {
  index: number;
  list: [OnlyOfficeParticipant];
}

export interface OnlyOfficeFromMessage {
  type?: string;
  openCmd?: {
    url?: string;
  };
  changes?: unknown;
  endSaveChanges?: boolean;
  [key: string]: any;
}

export interface OnlyOfficeToMessage {
  type: string;
  [key: string]: any;
}

export type OnlyOfficeMessageResponder = (msg: OnlyOfficeToMessage) => void;

export interface OnlyOfficeMockServer {
  buildNumber?: number;
  buildVersion?: string;
  getInitialChanges?: () => any[];
  getParticipants: () => OnlyOfficeParticipants;
  getImageURL?: (name: string) => Promise<string>;
  onAuth?: () => void;
  handleMessage?: (msg: OnlyOfficeFromMessage, respond: OnlyOfficeMessageResponder) => boolean;
  onMessage: (msg: OnlyOfficeFromMessage) => void;
  onCorruptionWarning?: (duplicateId: string) => void;
}

export interface CreateOnlyOfficeMockServerOptions {
  buildNumber?: number;
  buildVersion?: string;
  media?: Record<string, string>;
  onAuth?: () => void;
  onMessage?: (msg: OnlyOfficeFromMessage) => void;
  onCorruptionWarning?: (duplicateId: string) => void;
}

export const LOCAL_ONLYOFFICE_USER_ID = 'local-browser-user';
export const LOCAL_ONLYOFFICE_USER_NAME = 'Local Browser User';

const LOCAL_PARTICIPANT: OnlyOfficeParticipant = {
  id: 1,
  idOriginal: LOCAL_ONLYOFFICE_USER_ID,
  username: LOCAL_ONLYOFFICE_USER_NAME,
  indexUser: 0,
  connectionId: 'local-browser-session',
  isCloseCoAuthoring: false,
  view: false,
};

function stripLeadingDotsAndSlashes(value: string): string {
  return value.replace(/^(\.\/|\/)+/, '');
}

export function getOnlyOfficeMediaCandidates(name: string): string[] {
  const trimmed = stripLeadingDotsAndSlashes(name);
  const decoded = (() => {
    try {
      return decodeURIComponent(trimmed);
    } catch {
      return trimmed;
    }
  })();

  const candidates = [name, trimmed, decoded];
  for (const value of [trimmed, decoded]) {
    if (value && !value.startsWith('media/')) {
      candidates.push(`media/${value}`);
    }
  }

  return [...new Set(candidates.filter(Boolean))];
}

export function resolveOnlyOfficeMediaUrl(media: Record<string, string> | undefined, name: string): string {
  if (!media) {
    return '';
  }

  for (const candidate of getOnlyOfficeMediaCandidates(name)) {
    const url = media[candidate];
    if (url) {
      return url;
    }
  }

  return '';
}

function countSavedChanges(changes: unknown): number {
  if (!changes) {
    return 0;
  }

  if (Array.isArray(changes)) {
    return changes.length;
  }

  if (typeof changes === 'string') {
    try {
      const parsed = JSON.parse(changes) as unknown;
      return Array.isArray(parsed) ? parsed.length : 0;
    } catch {
      return 0;
    }
  }

  return 0;
}

export function createOnlyOfficeMockServer(options: CreateOnlyOfficeMockServerOptions = {}): OnlyOfficeMockServer {
  const { buildNumber, buildVersion, media, onAuth, onMessage, onCorruptionWarning } = options;
  let changesIndex = 0;
  let syncChangesIndex = 0;
  let lastSaveTime = Date.now();

  const acknowledgeSaveEnd = (respond: OnlyOfficeMessageResponder) => {
    lastSaveTime = Date.now();
    respond({
      type: 'unSaveLock',
      index: changesIndex,
      time: lastSaveTime,
      syncChangesIndex,
    });
  };

  return {
    buildNumber,
    buildVersion,
    getInitialChanges: () => [],
    getParticipants: () => ({
      index: 0,
      list: [{ ...LOCAL_PARTICIPANT }],
    }),
    getImageURL: async (name: string) => resolveOnlyOfficeMediaUrl(media, name),
    onAuth,
    handleMessage: (msg, respond) => {
      if (msg.type === 'isSaveLock') {
        respond({
          type: 'saveLock',
          saveLock: false,
          syncChangesIndex,
        });
        return true;
      }

      if (msg.type === 'saveChanges') {
        changesIndex += countSavedChanges(msg.changes);
        if (msg.endSaveChanges === false) {
          respond({
            type: 'savePartChanges',
            changesIndex,
            syncChangesIndex,
          });
          return true;
        }

        respond({
          type: 'saveChanges',
          changes: [],
          changesIndex,
          syncChangesIndex,
          endSaveChanges: true,
        });
        acknowledgeSaveEnd(respond);
        return true;
      }

      if (msg.type === 'unLockDocument' || msg.type === 'unSaveLock') {
        acknowledgeSaveEnd(respond);
        return true;
      }

      return false;
    },
    onMessage: (msg) => {
      onMessage?.(msg);
    },
    onCorruptionWarning,
  };
}
