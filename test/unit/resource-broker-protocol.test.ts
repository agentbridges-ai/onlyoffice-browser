import { describe, expect, it } from 'vitest';
import {
  OneTimeResourceBrokerCapabilityRegistry,
  RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES,
  RESOURCE_BROKER_MAX_READ_WINDOW_BYTES,
  RESOURCE_BROKER_PROTOCOL,
  isResourceBrokerEditorOrigin,
  listResourceBrokerEditorOrigins,
  normalizeResourceBrokerEditorOrigin,
  normalizeResourceBrokerReadWindowBytes,
  normalizeResourceBrokerResourcePath,
  normalizeOnlyOfficeRuntimeRequestPath,
  parseResourceBrokerChallengeMessage,
  parseResourceBrokerClientMessage,
  parseResourceBrokerRangeHeader,
  parseResourceBrokerServerMessage,
  resolveResourceBrokerRange,
  resolveResourceBrokerRangeHeader,
  type ResourceBrokerCapabilityClaim,
  type ResourceBrokerCapabilityMetadata,
} from '../../src/lib/resource-broker-protocol';

const editorOrigin = 'https://aries.getpi.work' as const;
const parentOrigin = 'https://piwork.getpi.work';
const releaseId = 'onlyoffice-browser-0.6.0+release.1';
const sessionId = 'office-session_1';
const token = 'ab'.repeat(32);

function capability(overrides: Partial<ResourceBrokerCapabilityMetadata> = {}): ResourceBrokerCapabilityMetadata {
  return {
    token,
    parentOrigin,
    editorOrigin,
    releaseId,
    sessionId,
    issuedAtMs: 1_000,
    expiresAtMs: 31_000,
    ...overrides,
  };
}

function claim(
  metadata: ResourceBrokerCapabilityMetadata,
  overrides: Partial<ResourceBrokerCapabilityClaim> = {},
): ResourceBrokerCapabilityClaim {
  return {
    token: metadata.token,
    parentOrigin: metadata.parentOrigin,
    editorOrigin: metadata.editorOrigin,
    releaseId: metadata.releaseId,
    sessionId: metadata.sessionId,
    ...overrides,
  };
}

describe('resource broker editor origins', () => {
  it('allows exactly the twelve production constellation origins', () => {
    expect(listResourceBrokerEditorOrigins()).toEqual([
      'https://aries.getpi.work',
      'https://taurus.getpi.work',
      'https://gemini.getpi.work',
      'https://cancer.getpi.work',
      'https://leo.getpi.work',
      'https://virgo.getpi.work',
      'https://libra.getpi.work',
      'https://scorpio.getpi.work',
      'https://sagittarius.getpi.work',
      'https://capricorn.getpi.work',
      'https://aquarius.getpi.work',
      'https://pisces.getpi.work',
    ]);
    for (const origin of listResourceBrokerEditorOrigins()) {
      expect(isResourceBrokerEditorOrigin(origin)).toBe(true);
      expect(normalizeResourceBrokerEditorOrigin(origin)).toBe(origin);
    }
  });

  it.each([
    'http://aries.getpi.work',
    'https://onlyoffice.getpi.work',
    'https://office-editor-evil.getpi.work',
    'https://aries.getpi.work.evil.example',
    'https://sub.aries.getpi.work',
    'https://aries.getpi.work:444',
    'https://aries.getpi.work/',
    'https://aries.getpi.work/path',
    'https://user@aries.getpi.work',
    'HTTPS://ARIES.GETPI.WORK',
    'not-an-origin',
  ])('rejects an uncontrolled or non-canonical origin: %s', (origin) => {
    expect(isResourceBrokerEditorOrigin(origin)).toBe(false);
  });

  it('strictly accepts bounded server response messages', () => {
    const probeResult = {
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'RESULT',
      id: 'probe-1',
      value: {
        releaseId,
        ready: true,
        probePath: 'sdkjs/word/sdk-all.js',
        probeBytes: 1_024,
        probeSha256: 'a'.repeat(64),
      },
    } as const;
    expect(parseResourceBrokerServerMessage(probeResult)).toEqual(probeResult);
    expect(
      parseResourceBrokerServerMessage({
        ...probeResult,
        value: { ...probeResult.value, probePath: '../office-host.html' },
      }),
    ).toBeNull();
    expect(
      parseResourceBrokerServerMessage({
        ...probeResult,
        value: { ...probeResult.value, probeBytes: 0 },
      }),
    ).toBeNull();
    expect(
      parseResourceBrokerServerMessage({
        ...probeResult,
        value: { ...probeResult.value, probeSha256: 'A'.repeat(64) },
      }),
    ).toBeNull();
    expect(
      parseResourceBrokerServerMessage({
        ...probeResult,
        value: { ...probeResult.value, extra: true },
      }),
    ).toBeNull();
    expect(
      parseResourceBrokerServerMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'HEADERS',
        id: 'read-1',
        status: 206,
        headers: {
          acceptRanges: 'bytes',
          contentLength: 4,
          contentRange: 'bytes 2-5/10',
          contentType: 'application/wasm',
        },
      }),
    ).toMatchObject({ type: 'HEADERS', status: 206 });
    expect(
      parseResourceBrokerServerMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CHUNK',
        id: 'read-1',
        bytes: new ArrayBuffer(256 * 1024),
      }),
    ).toMatchObject({ type: 'CHUNK' });
    expect(
      parseResourceBrokerServerMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CHUNK',
        id: 'read-1',
        bytes: new ArrayBuffer(RESOURCE_BROKER_MAX_READ_WINDOW_BYTES + 1),
      }),
    ).toBeNull();
    expect(
      parseResourceBrokerServerMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'ERROR',
        id: 'read-1',
        code: 'arbitrary-network-fallback',
      }),
    ).toBeNull();
  });
});

describe('resource broker path canonicalisation', () => {
  it.each([
    ['assets/sdk-all.js', 'assets/sdk-all.js'],
    ['/fonts/zh-CN/等线.ttf', 'fonts/zh-CN/等线.ttf'],
    ['fonts/Aptos%20Display.ttf', 'fonts/Aptos Display.ttf'],
    ['dictionaries/en_US/en_US.dic', 'dictionaries/en_US/en_US.dic'],
  ])('normalises a safe manifest path %s', (input, expected) => {
    expect(normalizeResourceBrokerResourcePath(input)).toBe(expected);
  });

  it.each([
    '',
    '/',
    '//evil.example/file',
    'https://evil.example/file',
    'data:text/plain,office',
    '../secret',
    'assets/../secret',
    'assets/./sdk.js',
    'assets//sdk.js',
    'assets/sdk.js/',
    '/%2e%2e/secret',
    '/%252e%252e/secret',
    '/assets%2fsdk.js',
    '/assets%252fsdk.js',
    '/assets%5csdk.js',
    '/assets%255csdk.js',
    '/https%3A%2F%2Fevil.example/file',
    'assets\\sdk.js',
    'assets/sdk.js?release=other',
    'assets/sdk.js#other',
    'assets/%00sdk.js',
    'assets/%',
    `assets/${'a'.repeat(2_049)}`,
  ])('rejects traversal, an arbitrary URL, or an ambiguous path: %s', (input) => {
    expect(normalizeResourceBrokerResourcePath(input)).toBeNull();
  });

  it('maps only the native ONLYOFFICE slide theme double separator to its signed manifest path', () => {
    expect(normalizeOnlyOfficeRuntimeRequestPath('/sdkjs/slide/themes//themes.js')).toBe(
      'sdkjs/slide/themes/themes.js',
    );
    expect(normalizeOnlyOfficeRuntimeRequestPath('/sdkjs/slide/themes//other.js')).toBeNull();
    expect(normalizeOnlyOfficeRuntimeRequestPath('/assets//sdk.js')).toBeNull();
  });
});

describe('resource broker Range semantics', () => {
  it('returns a complete 200 response when Range is absent', () => {
    expect(resolveResourceBrokerRangeHeader(null, 1_000)).toEqual({
      status: 200,
      start: 0,
      end: 999,
      contentLength: 1_000,
      contentRange: null,
    });
    expect(resolveResourceBrokerRangeHeader(undefined, 0)).toEqual({
      status: 200,
      start: 0,
      end: -1,
      contentLength: 0,
      contentRange: null,
    });
  });

  it.each([
    ['bytes=10-19', { start: 10, end: 19, contentLength: 10, contentRange: 'bytes 10-19/1000' }],
    ['bytes=990-', { start: 990, end: 999, contentLength: 10, contentRange: 'bytes 990-999/1000' }],
    ['bytes=-25', { start: 975, end: 999, contentLength: 25, contentRange: 'bytes 975-999/1000' }],
    ['bytes=-2000', { start: 0, end: 999, contentLength: 1_000, contentRange: 'bytes 0-999/1000' }],
    ['bytes=995-2000', { start: 995, end: 999, contentLength: 5, contentRange: 'bytes 995-999/1000' }],
  ])('resolves a single closed, open, or suffix range: %s', (header, expected) => {
    expect(resolveResourceBrokerRangeHeader(header, 1_000)).toMatchObject({
      status: 206,
      ...expected,
    });
  });

  it.each([
    'bytes=1000-1001',
    'bytes=20-10',
    'bytes=-0',
    'bytes=',
    'bytes=0-1,4-5',
    'items=0-1',
    'bytes=-',
    'bytes=9007199254740992-',
  ])('returns 416 for an invalid or unsatisfied range: %s', (header) => {
    expect(resolveResourceBrokerRangeHeader(header, 1_000)).toEqual({
      status: 416,
      start: null,
      end: null,
      contentLength: 0,
      contentRange: 'bytes */1000',
    });
  });

  it('keeps parsing separate from resource-size resolution', () => {
    expect(parseResourceBrokerRangeHeader('bytes=5-')).toEqual({ kind: 'open', start: 5 });
    expect(parseResourceBrokerRangeHeader('bytes=-7')).toEqual({ kind: 'suffix', bytes: 7 });
    expect(parseResourceBrokerRangeHeader('bytes=5-9')).toEqual({
      kind: 'closed',
      start: 5,
      end: 9,
    });
    expect(parseResourceBrokerRangeHeader('bytes=1-2,4-5')).toBeNull();
    expect(parseResourceBrokerRangeHeader(null)).toBeUndefined();
    expect(() => resolveResourceBrokerRange(undefined, -1)).toThrow(RangeError);
  });
});

describe('resource broker message discrimination', () => {
  const baseRead = {
    protocol: RESOURCE_BROKER_PROTOCOL,
    type: 'READ',
    id: 'read-1',
    releaseId,
    sessionId,
    path: 'wasm/x2t/x2t.wasm',
  } as const;

  it('accepts only the exact protocol version and known message shape', () => {
    expect(
      parseResourceBrokerClientMessage({
        ...baseRead,
        range: { kind: 'open', start: 1024 },
        windowBytes: 512 * 1024,
      }),
    ).toEqual({
      ...baseRead,
      range: { kind: 'open', start: 1024 },
      windowBytes: 512 * 1024,
    });
    expect(
      parseResourceBrokerClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PULL',
        id: 'read-1',
      }),
    ).toEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'PULL',
      id: 'read-1',
    });
    expect(
      parseResourceBrokerClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'PROBE',
        id: 'probe-1',
        releaseId,
        sessionId,
      }),
    ).not.toBeNull();
  });

  it.each([
    { ...baseRead, protocol: 'onlyoffice-browser-resource-broker/v2' },
    { ...baseRead, protocol: 1 },
    { ...baseRead, type: 'DELETE' },
    { ...baseRead, arbitraryUrl: 'https://evil.example' },
    { ...baseRead, path: '/wasm/x2t/x2t.wasm' },
    { ...baseRead, path: '../x2t.wasm' },
    { ...baseRead, id: 'bad request' },
    { ...baseRead, range: { kind: 'open', start: 0, end: 1 } },
    { ...baseRead, range: { kind: 'suffix', bytes: -1 } },
    { ...baseRead, windowBytes: RESOURCE_BROKER_MAX_READ_WINDOW_BYTES + 1 },
  ])('rejects a wrong-version, extended, or unsafe request %#', (message) => {
    expect(parseResourceBrokerClientMessage(message)).toBeNull();
  });

  it('strictly validates challenge and connect capability metadata', () => {
    const metadata = capability();
    expect(
      parseResourceBrokerChallengeMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CHALLENGE',
        capability: metadata,
      }),
    ).toEqual({
      protocol: RESOURCE_BROKER_PROTOCOL,
      type: 'CHALLENGE',
      capability: metadata,
    });
    expect(
      parseResourceBrokerClientMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CONNECT',
        capability: claim(metadata),
      }),
    ).not.toBeNull();
    expect(
      parseResourceBrokerChallengeMessage({
        protocol: RESOURCE_BROKER_PROTOCOL,
        type: 'CHALLENGE',
        capability: { ...metadata, unexpected: true },
      }),
    ).toBeNull();
  });

  it('enforces a hard upper bound for each pull window', () => {
    expect(normalizeResourceBrokerReadWindowBytes(undefined)).toBe(RESOURCE_BROKER_DEFAULT_READ_WINDOW_BYTES);
    expect(normalizeResourceBrokerReadWindowBytes(1)).toBe(1);
    expect(normalizeResourceBrokerReadWindowBytes(RESOURCE_BROKER_MAX_READ_WINDOW_BYTES)).toBe(
      RESOURCE_BROKER_MAX_READ_WINDOW_BYTES,
    );
    expect(normalizeResourceBrokerReadWindowBytes(0)).toBeNull();
    expect(normalizeResourceBrokerReadWindowBytes(1.5)).toBeNull();
    expect(normalizeResourceBrokerReadWindowBytes(RESOURCE_BROKER_MAX_READ_WINDOW_BYTES + 1)).toBeNull();
  });
});

describe('one-time resource broker capabilities', () => {
  const randomFill = (bytes: Uint8Array) => {
    bytes.fill(0xab);
    return bytes;
  };

  it('binds the token to event, parent, editor, release, and session identity and rejects replay', () => {
    const registry = new OneTimeResourceBrokerCapabilityRegistry({
      now: () => 1_000,
      randomFill,
    });
    const metadata = registry.issue({
      parentOrigin,
      editorOrigin,
      releaseId,
      sessionId,
      ttlMs: 10_000,
    });
    expect(metadata).toEqual(capability({ expiresAtMs: 11_000 }));
    expect(registry.size).toBe(1);

    expect(registry.consume(claim(metadata), editorOrigin, 2_000)).toEqual({
      ok: true,
      capability: metadata,
    });
    expect(registry.size).toBe(0);
    expect(registry.consume(claim(metadata), editorOrigin, 2_001)).toEqual({
      ok: false,
      code: 'unknown',
    });
  });

  it('burns a presented token after an identity mismatch or malformed claim', () => {
    const mismatchRegistry = new OneTimeResourceBrokerCapabilityRegistry({ randomFill });
    const mismatchCapability = mismatchRegistry.register(capability());
    expect(
      mismatchRegistry.consume(
        claim(mismatchCapability, { releaseId: 'onlyoffice-browser-other' }),
        editorOrigin,
        2_000,
      ),
    ).toEqual({ ok: false, code: 'mismatch' });
    expect(mismatchRegistry.consume(claim(mismatchCapability), editorOrigin, 2_000)).toEqual({
      ok: false,
      code: 'unknown',
    });

    const malformedRegistry = new OneTimeResourceBrokerCapabilityRegistry({ randomFill });
    const malformedCapability = malformedRegistry.register(capability({ token: 'cd'.repeat(32) }));
    expect(malformedRegistry.consume({ ...claim(malformedCapability), unexpected: true }, editorOrigin, 2_000)).toEqual(
      { ok: false, code: 'invalid' },
    );
    expect(malformedRegistry.size).toBe(0);
  });

  it('rejects expired capabilities and clears stale registrations', () => {
    const registry = new OneTimeResourceBrokerCapabilityRegistry({ randomFill });
    const expired = registry.register(capability());
    expect(registry.consume(claim(expired), editorOrigin, expired.expiresAtMs)).toEqual({
      ok: false,
      code: 'expired',
    });

    registry.register(capability({ token: 'cd'.repeat(32), expiresAtMs: 20_000 }));
    registry.register(capability({ token: 'ef'.repeat(32), expiresAtMs: 40_000 }));
    expect(registry.clearExpired(25_000)).toBe(1);
    expect(registry.size).toBe(1);
  });

  it('refuses weak metadata, uncontrolled origins, and excessive lifetimes', () => {
    const registry = new OneTimeResourceBrokerCapabilityRegistry({
      now: () => 1_000,
      randomFill,
    });
    expect(() =>
      registry.issue({
        parentOrigin,
        editorOrigin: 'https://office-editor-evil.getpi.work' as typeof editorOrigin,
        releaseId,
        sessionId,
      }),
    ).toThrow(TypeError);
    expect(() =>
      registry.issue({
        parentOrigin,
        editorOrigin,
        releaseId,
        sessionId,
        ttlMs: 60_001,
      }),
    ).toThrow(TypeError);
    expect(() => registry.register({ ...capability(), token: 'predictable' })).toThrow(TypeError);
  });
});
