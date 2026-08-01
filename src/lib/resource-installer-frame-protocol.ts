import {
  parseRequiredReleaseIdentity,
  type RequiredReleaseIdentity,
  type ResourceErrorCode,
  type ResourceInstallerSnapshot,
  type ResourcePlan,
  type ResourcePlanRequest,
} from './release-resources';

export const RESOURCE_INSTALLER_FRAME_PROTOCOL = 'onlyoffice-browser-resource-installer/v1' as const;
export const RESOURCE_INSTALLER_CANONICAL_ORIGIN = 'https://onlyoffice.getpi.work' as const;
export const RESOURCE_INSTALLER_FRAME_PATH = '/resource-installer.html' as const;
export const RESOURCE_INSTALLER_RPC_TIMEOUT_MS = 30_000;
export const RESOURCE_INSTALLER_CAPABILITY_TTL_MS = 30_000;
export const RESOURCE_INSTALLER_ALLOWED_PARENT_ORIGINS = [
  'https://piwork.getpi.work',
  'https://onlyoffice.getpi.work',
] as const;

const LOCAL_CANONICAL_HOSTNAME = 'onlyoffice.localhost';
const LOCAL_PARENT_HOSTNAMES = new Set(['piwork.localhost', 'onlyoffice.localhost']);
const TOKEN_PATTERN = /^[a-f0-9]{64}$/;
const REQUEST_ID_PATTERN = /^[a-zA-Z0-9._-]{1,128}$/;
const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const PLAN_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,256}$/;
const RESOURCE_PATH_MAX_LENGTH = 2_048;

const RESOURCE_SCOPES = ['recommended', 'document', 'all', 'repair', 'fonts'] as const;
const RESOURCE_PROFILES = ['base', 'word', 'cell', 'slide', 'fonts-basic', 'fonts-office-compat'] as const;
const DOCUMENT_TYPES = ['word', 'cell', 'slide'] as const;
const REPAIR_SCOPES = ['required', 'installed', 'all'] as const;
const READINESS_VALUES = [
  'checking',
  'needs-download',
  'ready',
  'update-available',
  'updating',
  'paused',
  'repair-needed',
  'error',
] as const;
const PHASE_VALUES = ['idle', 'planning', 'downloading', 'verifying', 'activating', 'repairing', 'paused'] as const;
const RESOURCE_ERROR_CODES = [
  'offline',
  'network',
  'timeout',
  'integrity',
  'quota',
  'manifest',
  'incompatible',
  'storage',
  'aborted',
] as const satisfies readonly ResourceErrorCode[];
const FRAME_ERROR_CODES = [
  ...RESOURCE_ERROR_CODES,
  'protocol',
  'identity',
  'capability',
  'unavailable',
  'busy',
] as const;

export type ResourceInstallerAllowedParentOrigin = (typeof RESOURCE_INSTALLER_ALLOWED_PARENT_ORIGINS)[number];
export type ResourceInstallerFrameCommand =
  | 'INIT'
  | 'SNAPSHOT'
  | 'PLAN'
  | 'APPLY'
  | 'CHECK_UPDATES'
  | 'CHECK_HEALTH'
  | 'REPAIR'
  | 'PAUSE'
  | 'RESUME'
  | 'CANCEL';
export type ResourceInstallerFrameErrorCode = (typeof FRAME_ERROR_CODES)[number];

export interface ResourceInstallerFrameIdentity {
  canonicalOrigin: string;
  physicalParentOrigin: string;
  logicalParentOrigin: ResourceInstallerAllowedParentOrigin;
  localTestMode: boolean;
}

export interface ResourceInstallerCapability {
  token: string;
  parentOrigin: string;
  issuedAtMs: number;
  expiresAtMs: number;
}

export interface ResourceInstallerChallengeMessage {
  protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
  type: 'CHALLENGE';
  capability: ResourceInstallerCapability;
}

export interface ResourceInstallerInitMessage {
  protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
  type: 'REQUEST';
  id: string;
  command: 'INIT';
  capability: {
    token: string;
    parentOrigin: string;
  };
  requiredReleaseIdentity: RequiredReleaseIdentity;
  subscribe: boolean;
}

export interface ResourceInstallerReconnectMessage {
  protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
  type: 'RECONNECT';
}

export type ResourceInstallerRequestMessage =
  | {
      protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
      type: 'REQUEST';
      id: string;
      command: 'SNAPSHOT' | 'CHECK_UPDATES' | 'CHECK_HEALTH' | 'PAUSE' | 'RESUME' | 'CANCEL';
    }
  | {
      protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
      type: 'REQUEST';
      id: string;
      command: 'PLAN';
      request: ResourcePlanRequest;
    }
  | {
      protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
      type: 'REQUEST';
      id: string;
      command: 'APPLY';
      plan: ResourcePlan;
    }
  | {
      protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
      type: 'REQUEST';
      id: string;
      command: 'REPAIR';
      scope: 'required' | 'installed' | 'all';
    };

export interface ResourceInstallerReadyMessage {
  protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
  type: 'READY';
  id: string;
  command: 'INIT';
  snapshot: ResourceInstallerSnapshot;
}

export type ResourceInstallerResultMessage =
  | {
      protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
      type: 'RESULT';
      id: string;
      command: 'PLAN';
      plan: ResourcePlan;
    }
  | {
      protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
      type: 'RESULT';
      id: string;
      command: Exclude<ResourceInstallerFrameCommand, 'INIT' | 'PLAN'>;
      snapshot: ResourceInstallerSnapshot;
    };

export interface ResourceInstallerSnapshotEventMessage {
  protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
  type: 'EVENT';
  event: 'SNAPSHOT';
  snapshot: ResourceInstallerSnapshot;
}

export interface ResourceInstallerErrorMessage {
  protocol: typeof RESOURCE_INSTALLER_FRAME_PROTOCOL;
  type: 'ERROR';
  id: string;
  command: ResourceInstallerFrameCommand;
  error: {
    code: ResourceInstallerFrameErrorCode;
    retryable: boolean;
    path?: string;
  };
}

export type ResourceInstallerServerMessage =
  | ResourceInstallerReadyMessage
  | ResourceInstallerResultMessage
  | ResourceInstallerSnapshotEventMessage
  | ResourceInstallerErrorMessage;

type CapabilityGateOptions = {
  parentOrigin: string;
  now?: () => number;
  randomFill?: (bytes: Uint8Array<ArrayBuffer>) => Uint8Array<ArrayBuffer>;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function hasExactKeys(
  value: Record<string, unknown>,
  required: readonly string[],
  optional: readonly string[] = [],
): boolean {
  const keys = Object.keys(value);
  const allowed = new Set([...required, ...optional]);
  return required.every((key) => Object.hasOwn(value, key)) && keys.every((key) => allowed.has(key));
}

function isOneOf<T extends string>(value: unknown, values: readonly T[]): value is T {
  return typeof value === 'string' && values.includes(value as T);
}

function isSafeNonNegativeInteger(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function isFiniteNonNegative(value: unknown): value is number {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0;
}

function parseExactOrigin(value: unknown): URL | null {
  if (typeof value !== 'string') return null;
  try {
    const url = new URL(value);
    if (url.username || url.password || url.pathname !== '/' || url.search || url.hash || url.origin !== value) {
      return null;
    }
    return url;
  } catch {
    return null;
  }
}

function getReferrerOrigin(value: string): string | null {
  try {
    return new URL(value).origin;
  } catch {
    return null;
  }
}

function hasOnlyTestMode(search: string): boolean {
  const params = new URLSearchParams(search);
  return (
    params.get('testMode') === '1' &&
    params.getAll('testMode').length === 1 &&
    [...params.keys()].every((key) => key === 'testMode')
  );
}

export function resolveResourceInstallerFrameIdentity(input: {
  locationOrigin: string;
  documentReferrer: string;
  search: string;
}): ResourceInstallerFrameIdentity | null {
  const canonical = parseExactOrigin(input.locationOrigin);
  const referrerOrigin = getReferrerOrigin(input.documentReferrer);
  const parent = parseExactOrigin(referrerOrigin);
  if (!canonical || !parent) return null;

  if (canonical.origin === RESOURCE_INSTALLER_CANONICAL_ORIGIN) {
    if (input.search !== '' || !RESOURCE_INSTALLER_ALLOWED_PARENT_ORIGINS.includes(parent.origin as never)) {
      return null;
    }
    return {
      canonicalOrigin: canonical.origin,
      physicalParentOrigin: parent.origin,
      logicalParentOrigin: parent.origin as ResourceInstallerAllowedParentOrigin,
      localTestMode: false,
    };
  }

  if (
    !hasOnlyTestMode(input.search) ||
    (canonical.protocol !== 'http:' && canonical.protocol !== 'https:') ||
    canonical.hostname !== LOCAL_CANONICAL_HOSTNAME ||
    parent.protocol !== canonical.protocol ||
    parent.port !== canonical.port ||
    !LOCAL_PARENT_HOSTNAMES.has(parent.hostname)
  ) {
    return null;
  }

  return {
    canonicalOrigin: canonical.origin,
    physicalParentOrigin: parent.origin,
    logicalParentOrigin:
      parent.hostname === 'piwork.localhost' ? 'https://piwork.getpi.work' : 'https://onlyoffice.getpi.work',
    localTestMode: true,
  };
}

export function resolveResourceInstallerClientIdentity(input: {
  parentOrigin: string;
  canonicalOrigin?: string;
  allowLocalTestMode?: boolean;
}): ResourceInstallerFrameIdentity | null {
  const canonical = parseExactOrigin(input.canonicalOrigin ?? RESOURCE_INSTALLER_CANONICAL_ORIGIN);
  const parent = parseExactOrigin(input.parentOrigin);
  if (!canonical || !parent) return null;

  if (canonical.origin === RESOURCE_INSTALLER_CANONICAL_ORIGIN) {
    if (!RESOURCE_INSTALLER_ALLOWED_PARENT_ORIGINS.includes(parent.origin as never)) return null;
    return {
      canonicalOrigin: canonical.origin,
      physicalParentOrigin: parent.origin,
      logicalParentOrigin: parent.origin as ResourceInstallerAllowedParentOrigin,
      localTestMode: false,
    };
  }

  if (
    input.allowLocalTestMode !== true ||
    (canonical.protocol !== 'http:' && canonical.protocol !== 'https:') ||
    canonical.hostname !== LOCAL_CANONICAL_HOSTNAME ||
    parent.protocol !== canonical.protocol ||
    parent.port !== canonical.port ||
    !LOCAL_PARENT_HOSTNAMES.has(parent.hostname)
  ) {
    return null;
  }
  return {
    canonicalOrigin: canonical.origin,
    physicalParentOrigin: parent.origin,
    logicalParentOrigin:
      parent.hostname === 'piwork.localhost' ? 'https://piwork.getpi.work' : 'https://onlyoffice.getpi.work',
    localTestMode: true,
  };
}

export function createResourceInstallerFrameUrl(identity: ResourceInstallerFrameIdentity): URL {
  const url = new URL(RESOURCE_INSTALLER_FRAME_PATH, identity.canonicalOrigin);
  if (identity.localTestMode) url.searchParams.set('testMode', '1');
  return url;
}

function isCapability(value: unknown): value is ResourceInstallerCapability {
  return (
    isRecord(value) &&
    hasExactKeys(value, ['token', 'parentOrigin', 'issuedAtMs', 'expiresAtMs']) &&
    typeof value.token === 'string' &&
    TOKEN_PATTERN.test(value.token) &&
    parseExactOrigin(value.parentOrigin) !== null &&
    isSafeNonNegativeInteger(value.issuedAtMs) &&
    isSafeNonNegativeInteger(value.expiresAtMs) &&
    value.expiresAtMs > value.issuedAtMs &&
    value.expiresAtMs - value.issuedAtMs <= RESOURCE_INSTALLER_CAPABILITY_TTL_MS
  );
}

export function parseResourceInstallerChallengeMessage(
  value: unknown,
  expectedParentOrigin: string,
  nowMs = Date.now(),
): ResourceInstallerChallengeMessage | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['protocol', 'type', 'capability']) ||
    value.protocol !== RESOURCE_INSTALLER_FRAME_PROTOCOL ||
    value.type !== 'CHALLENGE' ||
    !isCapability(value.capability) ||
    value.capability.parentOrigin !== expectedParentOrigin ||
    value.capability.issuedAtMs > nowMs ||
    value.capability.expiresAtMs <= nowMs
  ) {
    return null;
  }
  return value as unknown as ResourceInstallerChallengeMessage;
}

export function isResourceInstallerRequestId(value: unknown): value is string {
  return typeof value === 'string' && REQUEST_ID_PATTERN.test(value);
}

function isResourcePlanRequest(value: unknown): value is ResourcePlanRequest {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['scope'], ['documentType', 'profiles']) ||
    !isOneOf(value.scope, RESOURCE_SCOPES)
  ) {
    return false;
  }
  if (value.documentType !== undefined && !isOneOf(value.documentType, DOCUMENT_TYPES)) {
    return false;
  }
  if (value.scope === 'document' && !isOneOf(value.documentType, DOCUMENT_TYPES)) {
    return false;
  }
  if (value.scope !== 'document' && value.documentType !== undefined) {
    return false;
  }
  return (
    value.profiles === undefined ||
    (Array.isArray(value.profiles) &&
      value.profiles.length <= RESOURCE_PROFILES.length &&
      new Set(value.profiles).size === value.profiles.length &&
      value.profiles.every((profile) => isOneOf(profile, RESOURCE_PROFILES)))
  );
}

export function isResourceInstallerPlan(value: unknown): value is ResourcePlan {
  return (
    isRecord(value) &&
    hasExactKeys(value, ['planId', 'releaseId', 'scope', 'profiles', 'totalBytes', 'downloadBytes', 'reusedBytes']) &&
    typeof value.planId === 'string' &&
    PLAN_ID_PATTERN.test(value.planId) &&
    typeof value.releaseId === 'string' &&
    RELEASE_ID_PATTERN.test(value.releaseId) &&
    isOneOf(value.scope, RESOURCE_SCOPES) &&
    Array.isArray(value.profiles) &&
    value.profiles.length <= RESOURCE_PROFILES.length &&
    new Set(value.profiles).size === value.profiles.length &&
    value.profiles.every((profile) => isOneOf(profile, RESOURCE_PROFILES)) &&
    isSafeNonNegativeInteger(value.totalBytes) &&
    isSafeNonNegativeInteger(value.downloadBytes) &&
    isSafeNonNegativeInteger(value.reusedBytes) &&
    value.downloadBytes + value.reusedBytes === value.totalBytes
  );
}

export function parseResourceInstallerInitMessage(value: unknown): ResourceInstallerInitMessage | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'capability', 'requiredReleaseIdentity', 'subscribe']) ||
    value.protocol !== RESOURCE_INSTALLER_FRAME_PROTOCOL ||
    value.type !== 'REQUEST' ||
    value.command !== 'INIT' ||
    !isResourceInstallerRequestId(value.id) ||
    typeof value.subscribe !== 'boolean' ||
    !isRecord(value.capability) ||
    !hasExactKeys(value.capability, ['token', 'parentOrigin']) ||
    typeof value.capability.token !== 'string' ||
    !TOKEN_PATTERN.test(value.capability.token) ||
    parseExactOrigin(value.capability.parentOrigin) === null ||
    parseRequiredReleaseIdentity(value.requiredReleaseIdentity) === null
  ) {
    return null;
  }
  return {
    ...(value as unknown as ResourceInstallerInitMessage),
    requiredReleaseIdentity: parseRequiredReleaseIdentity(value.requiredReleaseIdentity)!,
  };
}

export function parseResourceInstallerReconnectMessage(value: unknown): ResourceInstallerReconnectMessage | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['protocol', 'type']) ||
    value.protocol !== RESOURCE_INSTALLER_FRAME_PROTOCOL ||
    value.type !== 'RECONNECT'
  ) {
    return null;
  }
  return value as unknown as ResourceInstallerReconnectMessage;
}

export function parseResourceInstallerRequestMessage(value: unknown): ResourceInstallerRequestMessage | null {
  if (
    !isRecord(value) ||
    value.protocol !== RESOURCE_INSTALLER_FRAME_PROTOCOL ||
    value.type !== 'REQUEST' ||
    !isResourceInstallerRequestId(value.id) ||
    typeof value.command !== 'string'
  ) {
    return null;
  }
  if (
    ['SNAPSHOT', 'CHECK_UPDATES', 'CHECK_HEALTH', 'PAUSE', 'RESUME', 'CANCEL'].includes(value.command) &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command'])
  ) {
    return value as unknown as ResourceInstallerRequestMessage;
  }
  if (
    value.command === 'PLAN' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'request']) &&
    isResourcePlanRequest(value.request)
  ) {
    return value as unknown as ResourceInstallerRequestMessage;
  }
  if (
    value.command === 'APPLY' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'plan']) &&
    isResourceInstallerPlan(value.plan)
  ) {
    return value as unknown as ResourceInstallerRequestMessage;
  }
  if (
    value.command === 'REPAIR' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'scope']) &&
    isOneOf(value.scope, REPAIR_SCOPES)
  ) {
    return value as unknown as ResourceInstallerRequestMessage;
  }
  return null;
}

function isNullableReleaseId(value: unknown): value is string | null {
  return value === null || (typeof value === 'string' && RELEASE_ID_PATTERN.test(value));
}

export function isResourceInstallerSnapshot(value: unknown): value is ResourceInstallerSnapshot {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, [
      'installedRelease',
      'targetRelease',
      'availableRelease',
      'availablePackageVersion',
      'readiness',
      'phase',
      'storageMode',
      'currentChunk',
      'currentChunkIndex',
      'currentChunkCount',
      'downloadedBytes',
      'downloadBytes',
      'verifiedBytes',
      'verifyBytes',
      'bytesPerSecond',
      'failedResources',
      'canPause',
      'canResume',
      'canRetry',
      'errorCode',
      'installedProfiles',
    ]) ||
    !isNullableReleaseId(value.installedRelease) ||
    !isNullableReleaseId(value.targetRelease) ||
    !isNullableReleaseId(value.availableRelease) ||
    (value.availablePackageVersion !== null && typeof value.availablePackageVersion !== 'string') ||
    !isOneOf(value.readiness, READINESS_VALUES) ||
    !isOneOf(value.phase, PHASE_VALUES) ||
    (value.storageMode !== 'cache-storage' && value.storageMode !== 'http-cache') ||
    (value.currentChunk !== null &&
      (typeof value.currentChunk !== 'string' ||
        value.currentChunk.length === 0 ||
        value.currentChunk.length > RESOURCE_PATH_MAX_LENGTH)) ||
    !isSafeNonNegativeInteger(value.currentChunkIndex) ||
    !isSafeNonNegativeInteger(value.currentChunkCount) ||
    value.currentChunkIndex > value.currentChunkCount ||
    !isSafeNonNegativeInteger(value.downloadedBytes) ||
    !isSafeNonNegativeInteger(value.downloadBytes) ||
    value.downloadedBytes > value.downloadBytes ||
    !isSafeNonNegativeInteger(value.verifiedBytes) ||
    !isSafeNonNegativeInteger(value.verifyBytes) ||
    value.verifiedBytes > value.verifyBytes ||
    !isFiniteNonNegative(value.bytesPerSecond) ||
    typeof value.canPause !== 'boolean' ||
    typeof value.canResume !== 'boolean' ||
    typeof value.canRetry !== 'boolean' ||
    (value.errorCode !== null && !isOneOf(value.errorCode, RESOURCE_ERROR_CODES)) ||
    !Array.isArray(value.installedProfiles) ||
    value.installedProfiles.length > RESOURCE_PROFILES.length ||
    new Set(value.installedProfiles).size !== value.installedProfiles.length ||
    !value.installedProfiles.every((profile) => isOneOf(profile, RESOURCE_PROFILES)) ||
    !Array.isArray(value.failedResources)
  ) {
    return false;
  }
  return value.failedResources.every(
    (failure) =>
      isRecord(failure) &&
      hasExactKeys(failure, ['path', 'code', 'attempts']) &&
      typeof failure.path === 'string' &&
      failure.path.length > 0 &&
      failure.path.length <= RESOURCE_PATH_MAX_LENGTH &&
      isOneOf(failure.code, RESOURCE_ERROR_CODES) &&
      isSafeNonNegativeInteger(failure.attempts) &&
      failure.attempts > 0,
  );
}

function parseError(value: unknown): ResourceInstallerErrorMessage['error'] | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['code', 'retryable'], ['path']) ||
    !isOneOf(value.code, FRAME_ERROR_CODES) ||
    typeof value.retryable !== 'boolean' ||
    (value.path !== undefined &&
      (typeof value.path !== 'string' || value.path.length === 0 || value.path.length > RESOURCE_PATH_MAX_LENGTH))
  ) {
    return null;
  }
  return value as unknown as ResourceInstallerErrorMessage['error'];
}

export function isResourceInstallerFrameCommand(value: unknown): value is ResourceInstallerFrameCommand {
  return isOneOf(value, [
    'INIT',
    'SNAPSHOT',
    'PLAN',
    'APPLY',
    'CHECK_UPDATES',
    'CHECK_HEALTH',
    'REPAIR',
    'PAUSE',
    'RESUME',
    'CANCEL',
  ]);
}

export function parseResourceInstallerServerMessage(value: unknown): ResourceInstallerServerMessage | null {
  if (!isRecord(value) || value.protocol !== RESOURCE_INSTALLER_FRAME_PROTOCOL || typeof value.type !== 'string') {
    return null;
  }
  if (
    value.type === 'READY' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'snapshot']) &&
    isResourceInstallerRequestId(value.id) &&
    value.command === 'INIT' &&
    isResourceInstallerSnapshot(value.snapshot)
  ) {
    return value as unknown as ResourceInstallerReadyMessage;
  }
  if (
    value.type === 'EVENT' &&
    hasExactKeys(value, ['protocol', 'type', 'event', 'snapshot']) &&
    value.event === 'SNAPSHOT' &&
    isResourceInstallerSnapshot(value.snapshot)
  ) {
    return value as unknown as ResourceInstallerSnapshotEventMessage;
  }
  if (
    value.type === 'ERROR' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'error']) &&
    isResourceInstallerRequestId(value.id) &&
    isResourceInstallerFrameCommand(value.command) &&
    parseError(value.error)
  ) {
    return value as unknown as ResourceInstallerErrorMessage;
  }
  if (
    value.type === 'RESULT' &&
    isResourceInstallerRequestId(value.id) &&
    value.command === 'PLAN' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'plan']) &&
    isResourceInstallerPlan(value.plan)
  ) {
    return value as unknown as ResourceInstallerResultMessage;
  }
  if (
    value.type === 'RESULT' &&
    isResourceInstallerRequestId(value.id) &&
    isResourceInstallerFrameCommand(value.command) &&
    value.command !== 'INIT' &&
    value.command !== 'PLAN' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'command', 'snapshot']) &&
    isResourceInstallerSnapshot(value.snapshot)
  ) {
    return value as unknown as ResourceInstallerResultMessage;
  }
  return null;
}

export class ResourceInstallerCapabilityGate {
  readonly #parentOrigin: string;
  readonly #now: () => number;
  readonly #randomFill: (bytes: Uint8Array<ArrayBuffer>) => Uint8Array<ArrayBuffer>;
  #capability: ResourceInstallerCapability | null = null;
  #consumed = false;

  constructor(options: CapabilityGateOptions) {
    if (!parseExactOrigin(options.parentOrigin)) throw new TypeError('Invalid Resource Installer parent origin');
    this.#parentOrigin = options.parentOrigin;
    this.#now = options.now ?? Date.now;
    this.#randomFill =
      options.randomFill ??
      ((bytes) => {
        crypto.getRandomValues(bytes);
        return bytes;
      });
  }

  issue(): ResourceInstallerCapability {
    if (this.#capability || this.#consumed) throw new Error('Resource Installer capability was already issued');
    const bytes = this.#randomFill(new Uint8Array(new ArrayBuffer(32)));
    if (!(bytes instanceof Uint8Array) || bytes.byteLength !== 32) {
      throw new Error('Resource Installer capability entropy source failed');
    }
    const issuedAtMs = this.#now();
    const token = [...bytes].map((byte) => byte.toString(16).padStart(2, '0')).join('');
    this.#capability = {
      token,
      parentOrigin: this.#parentOrigin,
      issuedAtMs,
      expiresAtMs: issuedAtMs + RESOURCE_INSTALLER_CAPABILITY_TTL_MS,
    };
    return { ...this.#capability };
  }

  consume(
    value: unknown,
    eventOrigin: string,
    nowMs = this.#now(),
  ): { ok: true; message: ResourceInstallerInitMessage } | { ok: false; code: 'invalid' | 'mismatch' | 'expired' } {
    const capability = this.#capability;
    if (this.#consumed || !capability) return { ok: false, code: 'invalid' };
    this.#capability = null;
    this.#consumed = true;
    const message = parseResourceInstallerInitMessage(value);
    if (!message) return { ok: false, code: 'invalid' };
    if (
      eventOrigin !== this.#parentOrigin ||
      message.capability.parentOrigin !== capability.parentOrigin ||
      message.capability.token !== capability.token
    ) {
      return { ok: false, code: 'mismatch' };
    }
    if (capability.expiresAtMs <= nowMs) return { ok: false, code: 'expired' };
    return { ok: true, message };
  }
}

export function resourceInstallerErrorPayload(error: unknown): ResourceInstallerErrorMessage['error'] {
  const record = isRecord(error) ? error : null;
  const code = record && isOneOf(record.code, FRAME_ERROR_CODES) ? record.code : 'unavailable';
  const retryable = !['integrity', 'manifest', 'incompatible', 'protocol', 'identity', 'capability'].includes(code);
  const path =
    record &&
    typeof record.path === 'string' &&
    record.path.length > 0 &&
    record.path.length <= RESOURCE_PATH_MAX_LENGTH
      ? record.path
      : undefined;
  return path ? { code, retryable, path } : { code, retryable };
}
