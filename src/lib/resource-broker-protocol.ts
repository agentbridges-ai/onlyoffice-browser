import {
  OFFICE_EDITOR_ORIGIN_SLOTS,
  type OfficeEditorOriginSlot,
  isProductionOfficeEditorHostname,
} from './office-origin-pool';

export const RESOURCE_BROKER_PROTOCOL = 'onlyoffice-browser-resource-broker/v1' as const;
export const RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES = 256 * 1024;
export const RESOURCE_BROKER_MAX_READ_WINDOW_BYTES = 1024 * 1024;
export const RESOURCE_BROKER_MAX_CAPABILITY_TTL_MS = 60_000;
export const RESOURCE_BROKER_MAX_RESOURCE_PATH_LENGTH = 2_048;

const RELEASE_ID_PATTERN = /^[a-zA-Z0-9._+-]{1,128}$/;
const SESSION_ID_PATTERN = /^[a-zA-Z0-9._-]{1,128}$/;
const REQUEST_ID_PATTERN = /^[a-zA-Z0-9._-]{1,128}$/;
const CAPABILITY_TOKEN_PATTERN = /^[a-f0-9]{64}$/;
const URI_SCHEME_PATTERN = /^[a-zA-Z][a-zA-Z0-9+.-]*:/;
const ENCODED_SEPARATOR_PATTERN = /%(?:2f|5c)/i;

export type ResourceBrokerEditorOrigin = `https://${OfficeEditorOriginSlot}.getpi.work`;

export interface ResourceBrokerIdentity {
  releaseId: string;
  sessionId: string;
}

export type ResourceBrokerByteRangeRequest =
  | {
      kind: 'closed';
      start: number;
      end: number;
    }
  | {
      kind: 'open';
      start: number;
    }
  | {
      kind: 'suffix';
      bytes: number;
    };

export type ResourceBrokerResolvedRange =
  | {
      status: 200;
      start: 0;
      end: number;
      contentLength: number;
      contentRange: null;
    }
  | {
      status: 206;
      start: number;
      end: number;
      contentLength: number;
      contentRange: string;
    }
  | {
      status: 416;
      start: null;
      end: null;
      contentLength: 0;
      contentRange: string;
    };

export interface ResourceBrokerCapabilityMetadata {
  token: string;
  parentOrigin: string;
  editorOrigin: ResourceBrokerEditorOrigin;
  releaseId: string;
  sessionId: string;
  issuedAtMs: number;
  expiresAtMs: number;
}

export interface ResourceBrokerCapabilityClaim {
  token: string;
  parentOrigin: string;
  editorOrigin: ResourceBrokerEditorOrigin;
  releaseId: string;
  sessionId: string;
}

export interface ResourceBrokerChallengeMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'CHALLENGE';
  capability: ResourceBrokerCapabilityMetadata;
}

export interface ResourceBrokerConnectMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'CONNECT';
  capability: ResourceBrokerCapabilityClaim;
}

export interface ResourceBrokerProbeMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'PROBE';
  id: string;
  releaseId: string;
  sessionId: string;
}

export interface ResourceBrokerReadMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'READ';
  id: string;
  releaseId: string;
  sessionId: string;
  path: string;
  range?: ResourceBrokerByteRangeRequest;
  windowBytes?: number;
}

export interface ResourceBrokerPullMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'PULL';
  id: string;
}

export interface ResourceBrokerCancelMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'CANCEL';
  id: string;
}

export interface ResourceBrokerProbeResultMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'RESULT';
  id: string;
  value: {
    releaseId: string;
    ready: true;
    probePath: string;
    probeBytes: number;
    probeSha256: string;
  };
}

export interface ResourceBrokerHeadersMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'HEADERS';
  id: string;
  status: 200 | 206 | 416;
  headers: {
    acceptRanges: 'bytes';
    contentLength: number;
    contentRange: string | null;
    contentType: string;
  };
}

export interface ResourceBrokerChunkMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'CHUNK';
  id: string;
  bytes: ArrayBuffer;
}

export interface ResourceBrokerEndMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'END' | 'CANCELLED';
  id: string;
  bytesSent: number;
}

export interface ResourceBrokerErrorMessage {
  protocol: typeof RESOURCE_BROKER_PROTOCOL;
  type: 'ERROR';
  id: string;
  code: 'missing' | 'release' | 'protocol' | 'timeout' | 'storage' | 'integrity' | 'busy' | 'cancelled';
}

export type ResourceBrokerServerMessage =
  | ResourceBrokerProbeResultMessage
  | ResourceBrokerHeadersMessage
  | ResourceBrokerChunkMessage
  | ResourceBrokerEndMessage
  | ResourceBrokerErrorMessage;

export type ResourceBrokerClientMessage =
  | ResourceBrokerConnectMessage
  | ResourceBrokerProbeMessage
  | ResourceBrokerReadMessage
  | ResourceBrokerPullMessage
  | ResourceBrokerCancelMessage;

export type ResourceBrokerCapabilityFailureCode = 'invalid' | 'unknown' | 'expired' | 'mismatch';

export type ResourceBrokerCapabilityConsumeResult =
  | {
      ok: true;
      capability: ResourceBrokerCapabilityMetadata;
    }
  | {
      ok: false;
      code: ResourceBrokerCapabilityFailureCode;
    };

type ResourceBrokerCapabilityRegistryOptions = {
  now?: () => number;
  randomFill?: (bytes: Uint8Array) => Uint8Array;
};

type ResourceBrokerCapabilityIssueOptions = Omit<
  ResourceBrokerCapabilityMetadata,
  'token' | 'issuedAtMs' | 'expiresAtMs'
> & {
  ttlMs?: number;
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
  const allowed = new Set([...required, ...optional]);
  const keys = Object.keys(value);
  return required.every((key) => Object.hasOwn(value, key)) && keys.every((key) => allowed.has(key));
}

function isSafeIntegerAtLeastZero(value: unknown): value is number {
  return Number.isSafeInteger(value) && Number(value) >= 0;
}

function hasControlCharacter(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (code <= 0x1f || code === 0x7f) return true;
  }
  return false;
}

export function isResourceBrokerReleaseId(value: unknown): value is string {
  return typeof value === 'string' && RELEASE_ID_PATTERN.test(value);
}

export function isResourceBrokerSessionId(value: unknown): value is string {
  return typeof value === 'string' && SESSION_ID_PATTERN.test(value);
}

export function isResourceBrokerRequestId(value: unknown): value is string {
  return typeof value === 'string' && REQUEST_ID_PATTERN.test(value);
}

export function normalizeResourceBrokerEditorOrigin(value: unknown): ResourceBrokerEditorOrigin | null {
  if (typeof value !== 'string') return null;
  try {
    const url = new URL(value);
    if (
      url.protocol !== 'https:' ||
      url.username ||
      url.password ||
      url.port ||
      url.pathname !== '/' ||
      url.search ||
      url.hash ||
      value !== url.origin ||
      !isProductionOfficeEditorHostname(url.hostname)
    ) {
      return null;
    }
    return url.origin as ResourceBrokerEditorOrigin;
  } catch {
    return null;
  }
}

export function isResourceBrokerEditorOrigin(value: unknown): value is ResourceBrokerEditorOrigin {
  return normalizeResourceBrokerEditorOrigin(value) !== null;
}

export function listResourceBrokerEditorOrigins(): readonly ResourceBrokerEditorOrigin[] {
  return OFFICE_EDITOR_ORIGIN_SLOTS.map((slot) => `https://${slot}.getpi.work` as ResourceBrokerEditorOrigin);
}

export function normalizeResourceBrokerParentOrigin(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  try {
    const url = new URL(value);
    if (
      url.protocol !== 'https:' ||
      url.username ||
      url.password ||
      url.pathname !== '/' ||
      url.search ||
      url.hash ||
      value !== url.origin
    ) {
      return null;
    }
    return url.origin;
  } catch {
    return null;
  }
}

/**
 * Converts an editor request pathname to the canonical relative path used by a
 * signed release manifest. Alternate spellings that could bypass manifest
 * membership checks are rejected instead of silently normalised.
 */
export function normalizeResourceBrokerResourcePath(value: unknown): string | null {
  if (
    typeof value !== 'string' ||
    value.length === 0 ||
    value.length > RESOURCE_BROKER_MAX_RESOURCE_PATH_LENGTH ||
    hasControlCharacter(value) ||
    value.includes('\\') ||
    value.includes('?') ||
    value.includes('#') ||
    value.startsWith('//') ||
    URI_SCHEME_PATTERN.test(value)
  ) {
    return null;
  }

  let decoded = value;
  for (let depth = 0; depth < 8; depth += 1) {
    if (ENCODED_SEPARATOR_PATTERN.test(decoded)) return null;
    let next: string;
    try {
      next = decodeURIComponent(decoded);
    } catch {
      return null;
    }
    if (next === decoded) break;
    decoded = next;
    if (depth === 7) return null;
  }

  decoded = decoded.normalize('NFC');
  if (
    hasControlCharacter(decoded) ||
    decoded.includes('\\') ||
    decoded.includes('?') ||
    decoded.includes('#') ||
    decoded.startsWith('//') ||
    URI_SCHEME_PATTERN.test(decoded)
  ) {
    return null;
  }

  if (decoded.startsWith('/')) decoded = decoded.slice(1);
  if (!decoded || decoded.startsWith('/') || decoded.endsWith('/')) return null;

  const segments = decoded.split('/');
  if (
    segments.some(
      (segment) =>
        segment.length === 0 ||
        segment === '.' ||
        segment === '..' ||
        segment.length > RESOURCE_BROKER_MAX_RESOURCE_PATH_LENGTH,
    )
  ) {
    return null;
  }

  return decoded.length <= RESOURCE_BROKER_MAX_RESOURCE_PATH_LENGTH ? decoded : null;
}

/**
 * ONLYOFFICE's slide theme loader requests this one DocumentServer path with
 * a redundant separator. Map only that established spelling to its signed
 * manifest entry; all other ambiguous paths still fail closed above.
 */
export function normalizeOnlyOfficeRuntimeRequestPath(value: unknown): string | null {
  if (value === '/sdkjs/slide/themes//themes.js' || value === 'sdkjs/slide/themes//themes.js') {
    return 'sdkjs/slide/themes/themes.js';
  }
  return normalizeResourceBrokerResourcePath(value);
}

export function parseResourceBrokerRangeHeader(
  value: string | null | undefined,
): ResourceBrokerByteRangeRequest | null | undefined {
  if (value === null || value === undefined) return undefined;
  const match = /^bytes=(\d*)-(\d*)$/.exec(value.trim());
  if (!match || (!match[1] && !match[2])) return null;

  const first = match[1] ? Number(match[1]) : undefined;
  const second = match[2] ? Number(match[2]) : undefined;
  if (
    (first !== undefined && !isSafeIntegerAtLeastZero(first)) ||
    (second !== undefined && !isSafeIntegerAtLeastZero(second))
  ) {
    return null;
  }
  if (first === undefined) return { kind: 'suffix', bytes: second! };
  if (second === undefined) return { kind: 'open', start: first };
  return { kind: 'closed', start: first, end: second };
}

export function resolveResourceBrokerRange(
  request: ResourceBrokerByteRangeRequest | null | undefined,
  totalBytes: number,
): ResourceBrokerResolvedRange {
  if (!isSafeIntegerAtLeastZero(totalBytes)) {
    throw new RangeError('totalBytes must be a non-negative safe integer');
  }
  if (request === undefined) {
    return {
      status: 200,
      start: 0,
      end: totalBytes - 1,
      contentLength: totalBytes,
      contentRange: null,
    };
  }

  const unsatisfied = (): ResourceBrokerResolvedRange => ({
    status: 416,
    start: null,
    end: null,
    contentLength: 0,
    contentRange: `bytes */${totalBytes}`,
  });
  if (request === null || totalBytes === 0) return unsatisfied();

  let start: number;
  let end: number;
  if (request.kind === 'suffix') {
    if (!Number.isSafeInteger(request.bytes) || request.bytes <= 0) return unsatisfied();
    start = Math.max(0, totalBytes - request.bytes);
    end = totalBytes - 1;
  } else {
    if (!isSafeIntegerAtLeastZero(request.start) || request.start >= totalBytes) return unsatisfied();
    start = request.start;
    if (request.kind === 'open') {
      end = totalBytes - 1;
    } else {
      if (!isSafeIntegerAtLeastZero(request.end) || request.end < start) return unsatisfied();
      end = Math.min(request.end, totalBytes - 1);
    }
  }

  return {
    status: 206,
    start,
    end,
    contentLength: end - start + 1,
    contentRange: `bytes ${start}-${end}/${totalBytes}`,
  };
}

export function resolveResourceBrokerRangeHeader(
  value: string | null | undefined,
  totalBytes: number,
): ResourceBrokerResolvedRange {
  return resolveResourceBrokerRange(parseResourceBrokerRangeHeader(value), totalBytes);
}

export function normalizeResourceBrokerReadWindowBytes(value: unknown): number | null {
  const candidate = value === undefined ? RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES : value;
  return Number.isSafeInteger(candidate) &&
    Number(candidate) > 0 &&
    Number(candidate) <= RESOURCE_BROKER_MAX_READ_WINDOW_BYTES
    ? Number(candidate)
    : null;
}

function parseByteRangeRequest(value: unknown): ResourceBrokerByteRangeRequest | null {
  if (!isRecord(value) || typeof value.kind !== 'string') return null;
  if (
    value.kind === 'closed' &&
    hasExactKeys(value, ['kind', 'start', 'end']) &&
    isSafeIntegerAtLeastZero(value.start) &&
    isSafeIntegerAtLeastZero(value.end)
  ) {
    return { kind: 'closed', start: value.start, end: value.end };
  }
  if (value.kind === 'open' && hasExactKeys(value, ['kind', 'start']) && isSafeIntegerAtLeastZero(value.start)) {
    return { kind: 'open', start: value.start };
  }
  if (value.kind === 'suffix' && hasExactKeys(value, ['kind', 'bytes']) && isSafeIntegerAtLeastZero(value.bytes)) {
    return { kind: 'suffix', bytes: value.bytes };
  }
  return null;
}

export function parseResourceBrokerCapabilityMetadata(value: unknown): ResourceBrokerCapabilityMetadata | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, [
      'token',
      'parentOrigin',
      'editorOrigin',
      'releaseId',
      'sessionId',
      'issuedAtMs',
      'expiresAtMs',
    ]) ||
    typeof value.token !== 'string' ||
    !CAPABILITY_TOKEN_PATTERN.test(value.token) ||
    !isResourceBrokerReleaseId(value.releaseId) ||
    !isResourceBrokerSessionId(value.sessionId) ||
    !isSafeIntegerAtLeastZero(value.issuedAtMs) ||
    !isSafeIntegerAtLeastZero(value.expiresAtMs) ||
    value.expiresAtMs <= value.issuedAtMs ||
    value.expiresAtMs - value.issuedAtMs > RESOURCE_BROKER_MAX_CAPABILITY_TTL_MS
  ) {
    return null;
  }
  const parentOrigin = normalizeResourceBrokerParentOrigin(value.parentOrigin);
  const editorOrigin = normalizeResourceBrokerEditorOrigin(value.editorOrigin);
  if (parentOrigin === null || parentOrigin !== value.parentOrigin || editorOrigin === null) return null;
  return {
    token: value.token,
    parentOrigin,
    editorOrigin,
    releaseId: value.releaseId,
    sessionId: value.sessionId,
    issuedAtMs: value.issuedAtMs,
    expiresAtMs: value.expiresAtMs,
  };
}

export function parseResourceBrokerCapabilityClaim(value: unknown): ResourceBrokerCapabilityClaim | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['token', 'parentOrigin', 'editorOrigin', 'releaseId', 'sessionId']) ||
    typeof value.token !== 'string' ||
    !CAPABILITY_TOKEN_PATTERN.test(value.token) ||
    !isResourceBrokerReleaseId(value.releaseId) ||
    !isResourceBrokerSessionId(value.sessionId)
  ) {
    return null;
  }
  const parentOrigin = normalizeResourceBrokerParentOrigin(value.parentOrigin);
  const editorOrigin = normalizeResourceBrokerEditorOrigin(value.editorOrigin);
  if (parentOrigin === null || parentOrigin !== value.parentOrigin || editorOrigin === null) return null;
  return {
    token: value.token,
    parentOrigin,
    editorOrigin,
    releaseId: value.releaseId,
    sessionId: value.sessionId,
  };
}

export function parseResourceBrokerChallengeMessage(value: unknown): ResourceBrokerChallengeMessage | null {
  if (
    !isRecord(value) ||
    !hasExactKeys(value, ['protocol', 'type', 'capability']) ||
    value.protocol !== RESOURCE_BROKER_PROTOCOL ||
    value.type !== 'CHALLENGE'
  ) {
    return null;
  }
  const capability = parseResourceBrokerCapabilityMetadata(value.capability);
  return capability
    ? {
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CHALLENGE',
        capability,
      }
    : null;
}

export function parseResourceBrokerClientMessage(value: unknown): ResourceBrokerClientMessage | null {
  if (!isRecord(value) || value.protocol !== RESOURCE_BROKER_PROTOCOL || typeof value.type !== 'string') {
    return null;
  }
  if (value.type === 'CONNECT' && hasExactKeys(value, ['protocol', 'type', 'capability'])) {
    const capability = parseResourceBrokerCapabilityClaim(value.capability);
    return capability
      ? {
          protocol: RESOURCE_BROKER_PROTOCOL,
          type: 'CONNECT',
          capability,
        }
      : null;
  }
  if (
    value.type === 'PROBE' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'releaseId', 'sessionId']) &&
    isResourceBrokerRequestId(value.id) &&
    isResourceBrokerReleaseId(value.releaseId) &&
    isResourceBrokerSessionId(value.sessionId)
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'PROBE',
      id: value.id,
      releaseId: value.releaseId,
      sessionId: value.sessionId,
    };
  }
  if (
    value.type === 'READ' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'releaseId', 'sessionId', 'path'], ['range', 'windowBytes']) &&
    isResourceBrokerRequestId(value.id) &&
    isResourceBrokerReleaseId(value.releaseId) &&
    isResourceBrokerSessionId(value.sessionId)
  ) {
    const path = normalizeResourceBrokerResourcePath(value.path);
    if (!path || path !== value.path) return null;
    const range = Object.hasOwn(value, 'range') ? parseByteRangeRequest(value.range) : undefined;
    if (Object.hasOwn(value, 'range') && !range) return null;
    let windowBytes: number | undefined;
    if (Object.hasOwn(value, 'windowBytes')) {
      const normalizedWindowBytes = normalizeResourceBrokerReadWindowBytes(value.windowBytes);
      if (normalizedWindowBytes === null) return null;
      windowBytes = normalizedWindowBytes;
    }
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'READ',
      id: value.id,
      releaseId: value.releaseId,
      sessionId: value.sessionId,
      path,
      ...(range ? { range } : {}),
      ...(windowBytes !== undefined ? { windowBytes } : {}),
    };
  }
  if (
    (value.type === 'PULL' || value.type === 'CANCEL') &&
    hasExactKeys(value, ['protocol', 'type', 'id']) &&
    isResourceBrokerRequestId(value.id)
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: value.type,
      id: value.id,
    };
  }
  return null;
}

export function parseResourceBrokerServerMessage(value: unknown): ResourceBrokerServerMessage | null {
  if (
    !isRecord(value) ||
    value.protocol !== RESOURCE_BROKER_PROTOCOL ||
    typeof value.type !== 'string' ||
    !isResourceBrokerRequestId(value.id)
  ) {
    return null;
  }
  if (value.type === 'RESULT' && hasExactKeys(value, ['protocol', 'type', 'id', 'value']) && isRecord(value.value)) {
    const probePath = normalizeResourceBrokerResourcePath(value.value.probePath);
    if (
      !hasExactKeys(value.value, ['releaseId', 'ready', 'probePath', 'probeBytes', 'probeSha256']) ||
      !isResourceBrokerReleaseId(value.value.releaseId) ||
      value.value.ready !== true ||
      probePath === null ||
      probePath !== value.value.probePath ||
      !Number.isSafeInteger(value.value.probeBytes) ||
      Number(value.value.probeBytes) <= 0 ||
      typeof value.value.probeSha256 !== 'string' ||
      !CAPABILITY_TOKEN_PATTERN.test(value.value.probeSha256)
    ) {
      return null;
    }
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'RESULT',
      id: value.id,
      value: {
        releaseId: value.value.releaseId,
        ready: true,
        probePath,
        probeBytes: Number(value.value.probeBytes),
        probeSha256: value.value.probeSha256,
      },
    };
  }
  if (
    value.type === 'HEADERS' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'status', 'headers']) &&
    (value.status === 200 || value.status === 206 || value.status === 416) &&
    isRecord(value.headers) &&
    hasExactKeys(value.headers, ['acceptRanges', 'contentLength', 'contentRange', 'contentType']) &&
    value.headers.acceptRanges === 'bytes' &&
    isSafeIntegerAtLeastZero(value.headers.contentLength) &&
    (value.headers.contentRange === null || typeof value.headers.contentRange === 'string') &&
    typeof value.headers.contentType === 'string' &&
    value.headers.contentType.length > 0
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'HEADERS',
      id: value.id,
      status: value.status,
      headers: {
        acceptRanges: 'bytes',
        contentLength: value.headers.contentLength,
        contentRange: value.headers.contentRange,
        contentType: value.headers.contentType,
      },
    };
  }
  if (
    value.type === 'CHUNK' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'bytes']) &&
    value.bytes instanceof ArrayBuffer &&
    value.bytes.byteLength > 0 &&
    value.bytes.byteLength <= RESOURCE_BROKER_MAX_READ_WINDOW_BYTES
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHUNK',
      id: value.id,
      bytes: value.bytes,
    };
  }
  if (
    (value.type === 'END' || value.type === 'CANCELLED') &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'bytesSent']) &&
    isSafeIntegerAtLeastZero(value.bytesSent)
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: value.type,
      id: value.id,
      bytesSent: value.bytesSent,
    };
  }
  if (
    value.type === 'ERROR' &&
    hasExactKeys(value, ['protocol', 'type', 'id', 'code']) &&
    ['missing', 'release', 'protocol', 'timeout', 'storage', 'integrity', 'busy', 'cancelled'].includes(
      String(value.code),
    )
  ) {
    return {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'ERROR',
      id: value.id,
      code: value.code as ResourceBrokerErrorMessage['code'],
    };
  }
  return null;
}

export class OneTimeResourceBrokerCapabilityRegistry {
  readonly #capabilities = new Map<string, ResourceBrokerCapabilityMetadata>();
  readonly #now: () => number;
  readonly #randomFill: (bytes: Uint8Array) => Uint8Array;

  constructor(options: ResourceBrokerCapabilityRegistryOptions = {}) {
    this.#now = options.now ?? Date.now;
    this.#randomFill =
      options.randomFill ??
      ((bytes) => {
        globalThis.crypto.getRandomValues(bytes as Uint8Array<ArrayBuffer>);
        return bytes;
      });
  }

  get size(): number {
    return this.#capabilities.size;
  }

  issue(options: ResourceBrokerCapabilityIssueOptions): ResourceBrokerCapabilityMetadata {
    const issuedAtMs = this.#now();
    const ttlMs = options.ttlMs ?? 30_000;
    if (
      !isSafeIntegerAtLeastZero(issuedAtMs) ||
      !Number.isSafeInteger(ttlMs) ||
      ttlMs <= 0 ||
      ttlMs > RESOURCE_BROKER_MAX_CAPABILITY_TTL_MS
    ) {
      throw new TypeError('Invalid capability lifetime');
    }
    const randomBytes = new Uint8Array(32);
    this.#randomFill(randomBytes);
    const token = Array.from(randomBytes, (byte) => byte.toString(16).padStart(2, '0')).join('');
    const capability = parseResourceBrokerCapabilityMetadata({
      token,
      parentOrigin: options.parentOrigin,
      editorOrigin: options.editorOrigin,
      releaseId: options.releaseId,
      sessionId: options.sessionId,
      issuedAtMs,
      expiresAtMs: issuedAtMs + ttlMs,
    });
    if (!capability) throw new TypeError('Invalid capability metadata');
    if (this.#capabilities.has(token)) throw new Error('Capability token collision');
    this.#capabilities.set(token, Object.freeze(capability));
    return capability;
  }

  register(value: unknown): ResourceBrokerCapabilityMetadata {
    const capability = parseResourceBrokerCapabilityMetadata(value);
    if (!capability) throw new TypeError('Invalid capability metadata');
    if (this.#capabilities.has(capability.token)) throw new Error('Capability already registered');
    this.#capabilities.set(capability.token, Object.freeze(capability));
    return capability;
  }

  consume(value: unknown, eventOrigin: unknown, nowMs = this.#now()): ResourceBrokerCapabilityConsumeResult {
    const presentedToken =
      isRecord(value) && typeof value.token === 'string' && CAPABILITY_TOKEN_PATTERN.test(value.token)
        ? value.token
        : null;
    const capability = presentedToken ? this.#capabilities.get(presentedToken) : undefined;
    if (presentedToken) this.#capabilities.delete(presentedToken);

    const claim = parseResourceBrokerCapabilityClaim(value);
    if (!claim) return { ok: false, code: 'invalid' };
    if (!capability) return { ok: false, code: 'unknown' };
    if (!isSafeIntegerAtLeastZero(nowMs) || nowMs < capability.issuedAtMs || nowMs >= capability.expiresAtMs) {
      return { ok: false, code: 'expired' };
    }
    if (
      eventOrigin !== capability.editorOrigin ||
      claim.parentOrigin !== capability.parentOrigin ||
      claim.editorOrigin !== capability.editorOrigin ||
      claim.releaseId !== capability.releaseId ||
      claim.sessionId !== capability.sessionId
    ) {
      return { ok: false, code: 'mismatch' };
    }
    return { ok: true, capability };
  }

  revoke(token: string): boolean {
    return this.#capabilities.delete(token);
  }

  clearExpired(nowMs = this.#now()): number {
    let removed = 0;
    for (const [token, capability] of this.#capabilities) {
      if (nowMs >= capability.expiresAtMs) {
        this.#capabilities.delete(token);
        removed += 1;
      }
    }
    return removed;
  }
}
