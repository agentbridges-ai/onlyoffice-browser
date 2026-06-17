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
  [key: string]: any;
}

export interface OnlyOfficeMockServer {
  buildNumber?: number;
  buildVersion?: string;
  getInitialChanges?: () => any[];
  getParticipants: () => OnlyOfficeParticipants;
  getImageURL?: (name: string) => Promise<string>;
  onAuth?: () => void;
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

const LOCAL_PARTICIPANT: OnlyOfficeParticipant = {
  id: 1,
  idOriginal: 'local-browser-user',
  username: 'Guest',
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

export function createOnlyOfficeMockServer(options: CreateOnlyOfficeMockServerOptions = {}): OnlyOfficeMockServer {
  const { buildNumber, buildVersion, media, onAuth, onMessage, onCorruptionWarning } = options;

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
    onMessage: (msg) => {
      onMessage?.(msg);
    },
    onCorruptionWarning,
  };
}
