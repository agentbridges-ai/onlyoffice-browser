import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import worker, {
  type LocalMatrixCounter,
  type ReleaseManifest,
  type WorkerEnv,
  applyReleaseMime,
  canonicalReleasePathname,
  isIsolatedEditorHost,
  isAssetRevision,
  isImmutableReleaseReceiptKey,
  isOnlyOfficeHost,
  releaseIdFromReferrer,
  resolveContentSegmentRequest,
  resolveEditorAssetRoute,
  resolveObjectKey,
  resolveReleaseContentObject,
  resolveReleaseContentObjectRequest,
  resolveReleasePackageRequest,
  resolveReleaseRequest,
  resolveRequestHostname,
  resolveRuntimeHost,
  shouldDisableResponseTransform,
  shouldShareAsset,
  shouldPopulateEdgeCache,
} from '../../cloudflare/worker';
import { isOfficeEditorOriginSlot, isReusableOfficeEditorHostname } from '../../src/lib/office-origin-pool';

const digest = (value: string) => value.repeat(64);

function v5Manifest(
  releaseId = 'v5-release',
  options: { wholeSha256?: string; bytes?: number; fastcdc?: boolean } = {},
): ReleaseManifest {
  const wholeSha256 = options.wholeSha256 || digest('a');
  const bytes = options.bytes || 6;
  const packageBytes = bytes + 16;
  const packageSha256 = digest('e');
  return {
    version: 5,
    releaseId,
    contentProtocol: {
      version: 1,
      digest: 'sha256',
      cacheKeyFormat: 'canonical-sha256-v1',
      storageSetSha256: digest('f'),
      fastcdcPolicyId: 'fastcdc-v2020-min64k-avg256k-max1m-norm1-seed0',
    },
    package: {
      format: 'onlyoffice-pack-v1',
      path: 'office-resources.oobpack',
      bytes: packageBytes,
      sha256: packageSha256,
      segments: [{ id: packageSha256, offset: 0, bytes: packageBytes, sha256: packageSha256 }],
    },
    assets: [
      {
        path: 'sdkjs/word/sdk-all.js',
        sha256: wholeSha256,
        mime: 'text/javascript; charset=utf-8',
        bytes,
        profile: 'word',
        chunk: 'word-001',
        packageOffset: 16,
        representations: {
          whole: { sha256: wholeSha256, bytes },
          ...(options.fastcdc
            ? {
                fastcdc: {
                  algorithm: 'fastcdc-v2020',
                  minBytes: 65_536,
                  averageBytes: 262_144,
                  maxBytes: 1_048_576,
                  normalization: 1,
                  seed: 0,
                  chunks: [
                    { offset: 0, bytes: 2, sha256: digest('b') },
                    { offset: 2, bytes: bytes - 2, sha256: digest('c') },
                  ],
                },
              }
            : {}),
        },
      },
    ],
  };
}

describe('Cloudflare OnlyOffice runtime routing', () => {
  it('recognizes only canonical lowercase constellation slot values', () => {
    expect(isOfficeEditorOriginSlot('aries')).toBe(true);
    expect(isOfficeEditorOriginSlot('ARIES')).toBe(false);
    expect(isOfficeEditorOriginSlot('orion')).toBe(false);
  });

  it('retains runtime storage only for the fixed production and local origin pool', () => {
    expect(isReusableOfficeEditorHostname('aries.getpi.work')).toBe(true);
    expect(isReusableOfficeEditorHostname('pisces.localhost')).toBe(true);
    expect(isReusableOfficeEditorHostname('host-gemini.office.localhost')).toBe(true);
    expect(isReusableOfficeEditorHostname('office-editor-legacy.getpi.work')).toBe(false);
    expect(isReusableOfficeEditorHostname('orion.getpi.work')).toBe(false);
  });

  it('accepts the canonical, fixed constellation, and legacy wildcard editor hosts', () => {
    expect(isOnlyOfficeHost('onlyoffice.getpi.work')).toBe(true);
    expect(isOnlyOfficeHost('aries.getpi.work')).toBe(true);
    expect(isOnlyOfficeHost('pisces.getpi.work')).toBe(true);
    expect(isOnlyOfficeHost('office-editor-a.getpi.work')).toBe(true);
    expect(isOnlyOfficeHost('unrelated.getpi.work')).toBe(false);
    expect(isOnlyOfficeHost('office-a.dev.getpi.work')).toBe(false);
    expect(isOnlyOfficeHost('onlyoffice.getpi.work.example.com')).toBe(false);
    expect(isIsolatedEditorHost('gemini.getpi.work')).toBe(true);
    expect(isIsolatedEditorHost('office-editor-a.getpi.work')).toBe(true);
    expect(isIsolatedEditorHost('onlyoffice.getpi.work')).toBe(false);
  });

  it('maps loopback matrix origins onto the production host classes only when enabled', () => {
    expect(resolveRuntimeHost('onlyoffice.localhost', true)).toEqual({
      logicalHostname: 'onlyoffice.getpi.work',
      canonicalHostname: 'assets.office.localhost',
    });
    expect(resolveRuntimeHost('assets.office.localhost', true)).toEqual({
      logicalHostname: 'onlyoffice.getpi.work',
      canonicalHostname: 'assets.office.localhost',
    });
    expect(resolveRuntimeHost('aries.localhost', true)).toEqual({
      logicalHostname: 'aries.getpi.work',
      canonicalHostname: 'assets.office.localhost',
    });
    expect(resolveRuntimeHost('host-gemini.office.localhost', true)).toEqual({
      logicalHostname: 'gemini.getpi.work',
      canonicalHostname: 'assets.office.localhost',
    });
    expect(resolveRuntimeHost('office-editor-matrix.localhost', true)).toEqual({
      logicalHostname: 'office-editor-matrix.getpi.work',
      canonicalHostname: 'assets.office.localhost',
    });
    expect(resolveRuntimeHost('host-office-editor-matrix.office.localhost', true)).toEqual({
      logicalHostname: 'office-editor-matrix.getpi.work',
      canonicalHostname: 'assets.office.localhost',
    });
    expect(resolveRuntimeHost('office-editor-matrix.localhost', false)).toEqual({
      logicalHostname: 'office-editor-matrix.localhost',
      canonicalHostname: 'onlyoffice.getpi.work',
    });
  });

  it('uses the matrix gateway host only in local mode and fails closed after Wrangler host rewriting', () => {
    expect(
      resolveRequestHostname(
        new Request('https://onlyoffice.getpi.work/', {
          headers: { 'X-OnlyOffice-Matrix-Host': 'Aries.Localhost:8787' },
        }),
        true,
      ),
    ).toBe('aries.localhost');
    expect(
      resolveRequestHostname(
        new Request('https://onlyoffice.getpi.work/', {
          headers: { 'X-OnlyOffice-Matrix-Host': 'aries.localhost:8787' },
        }),
        false,
      ),
    ).toBe('onlyoffice.getpi.work');
    expect(resolveRequestHostname(new Request('https://onlyoffice.getpi.work/'), true)).toBeNull();
    expect(resolveRequestHostname(new Request('https://aries.localhost/'), true)).toBe('aries.localhost');
    expect(
      resolveRequestHostname(
        new Request('https://onlyoffice.getpi.work/', {
          headers: { 'X-OnlyOffice-Matrix-Host': 'aries.localhost/../../evil' },
        }),
        true,
      ),
    ).toBeNull();
  });

  it('keeps only origin-bound boot files, bounded shell chunks, and workers on each editor origin', () => {
    expect(shouldShareAsset('/office-host.html', 'document')).toBe(false);
    expect(shouldShareAsset('/editor-shell-prime.html', 'iframe')).toBe(false);
    expect(shouldShareAsset('/assets/officeHost-a.js', 'script')).toBe(false);
    expect(shouldShareAsset('/assets/base-a.css', 'style')).toBe(false);
    expect(shouldShareAsset('/assets/nested/officeHost-a.js', 'script')).toBe(true);
    expect(shouldShareAsset('/wasm/x2t/worker.js', 'worker')).toBe(false);
    expect(shouldShareAsset('/wasm/x2t/conversion-worker-a.js', null)).toBe(false);
    expect(shouldShareAsset('/wasm/x2t/startup-heartbeat-worker-a.js', 'empty')).toBe(false);
    expect(shouldShareAsset('/wasm/x2t/x2t.js', 'script')).toBe(false);
    expect(shouldShareAsset('/wasm/x2t/x2t.wasm', 'empty')).toBe(true);
    expect(shouldShareAsset('/document_editor_service_worker.js', 'serviceworker')).toBe(false);
    expect(shouldShareAsset('/document_editor_service_worker.js', null)).toBe(false);
    expect(shouldShareAsset('/sw.js', 'script')).toBe(false);
    expect(shouldShareAsset('/sdkjs/word/sdk-all.js', 'script')).toBe(true);
    expect(shouldShareAsset('/fonts/000.ttf', 'empty')).toBe(true);
  });

  it('accepts only immutable SHA-256 segment paths', () => {
    const sha256 = 'a'.repeat(64);
    expect(resolveContentSegmentRequest(`/segments/sha256/${sha256}`)).toEqual({ sha256 });
    expect(resolveContentSegmentRequest('/segments/sha256/not-a-digest')).toBeNull();
  });

  it('accepts only release-bound immutable v5 content object paths', () => {
    const sha256 = digest('a');
    expect(resolveReleaseContentObjectRequest(`/objects/v0.6.0-build.1+canary/sha256/${sha256}`)).toEqual({
      releaseId: 'v0.6.0-build.1+canary',
      sha256,
    });
    expect(resolveReleaseContentObjectRequest(`/objects/v0.6.0%2B1/sha256/${sha256}`)).toEqual({
      releaseId: 'v0.6.0+1',
      sha256,
    });
    expect(resolveReleaseContentObjectRequest(`/objects/v0.6.0/sha256/${digest('A')}`)).toBeNull();
    expect(resolveReleaseContentObjectRequest('/objects/v0.6.0/sha256/not-a-digest')).toBeNull();
    expect(resolveReleaseContentObjectRequest(`/objects/%2e%2e%2fsecret/sha256/${sha256}`)).toBeNull();
  });

  it('binds whole and FastCDC objects to the exact compatible v5 release manifest', () => {
    const manifest = v5Manifest('v5-a', { fastcdc: true });
    expect(resolveReleaseContentObject(manifest, { releaseId: 'v5-a', sha256: digest('a') })).toEqual({
      sha256: digest('a'),
      bytes: 6,
    });
    expect(resolveReleaseContentObject(manifest, { releaseId: 'v5-a', sha256: digest('b') })).toEqual({
      sha256: digest('b'),
      bytes: 2,
    });
    expect(resolveReleaseContentObject(manifest, { releaseId: 'v5-a', sha256: digest('c') })).toEqual({
      sha256: digest('c'),
      bytes: 4,
    });
    expect(resolveReleaseContentObject(manifest, { releaseId: 'v5-a', sha256: digest('e') })).toEqual({
      sha256: digest('e'),
      bytes: 22,
    });
    expect(resolveReleaseContentObject(manifest, { releaseId: 'v5-b', sha256: digest('a') })).toBeNull();
    expect(resolveReleaseContentObject(manifest, { releaseId: 'v5-a', sha256: digest('d') })).toBeNull();
  });

  it('fails closed for incompatible or malformed v5 content object sets', () => {
    const wrongProtocol = v5Manifest();
    wrongProtocol.contentProtocol!.fastcdcPolicyId = 'fastcdc-v2016';
    expect(resolveReleaseContentObject(wrongProtocol, { releaseId: 'v5-release', sha256: digest('a') })).toBeNull();

    const gap = v5Manifest('v5-gap', { fastcdc: true });
    gap.assets[0].representations!.fastcdc!.chunks![1].offset = 3;
    expect(resolveReleaseContentObject(gap, { releaseId: 'v5-gap', sha256: digest('b') })).toBeNull();

    const conflictingDigest = v5Manifest('v5-conflict', { fastcdc: true });
    conflictingDigest.assets[0].representations!.fastcdc!.chunks![0].sha256 = digest('a');
    expect(
      resolveReleaseContentObject(conflictingDigest, { releaseId: 'v5-conflict', sha256: digest('a') }),
    ).toBeNull();

    const legacy: ReleaseManifest = {
      version: 4,
      releaseId: 'v4-release',
      assets: [],
      package: {
        format: 'onlyoffice-pack-v1',
        path: 'office-resources.oobpack',
        bytes: 1,
        sha256: digest('e'),
      },
    };
    expect(resolveReleaseContentObject(legacy, { releaseId: 'v4-release', sha256: digest('e') })).toBeNull();
  });

  it('maps the public root and rejects traversal keys', () => {
    expect(resolveObjectKey('/')).toBe('index.html');
    expect(resolveObjectKey('/sdkjs/word/sdk-all.js')).toBe('sdkjs/word/sdk-all.js');
    expect(resolveObjectKey('/sdkjs/slide/themes//themes.js')).toBe('sdkjs/slide/themes/themes.js');
    expect(resolveObjectKey('/sdkjs/%2e%2e/secret')).toBeNull();
    expect(resolveObjectKey('/%252e%252e/%252e%252e/objects/other/sha256/abc')).toBeNull();
    expect(resolveObjectKey('/%25252e%25252e/blobs/sha256/abc')).toBeNull();
    expect(resolveObjectKey('/sdkjs/%255csecret')).toBeNull();
    expect(resolveObjectKey('/%E0%A4%A')).toBeNull();
  });

  it('maps immutable release URLs without allowing traversal', () => {
    expect(resolveReleaseRequest('/r/v0.4.0-abcd/sdkjs/word/word.js')).toEqual({
      releaseId: 'v0.4.0-abcd',
      path: 'sdkjs/word/word.js',
    });
    expect(resolveReleaseRequest('/r/v0.4.0%2B1/office-host.html')).toEqual({
      releaseId: 'v0.4.0+1',
      path: 'office-host.html',
    });
    expect(resolveReleaseRequest('/r/v0.4.0-abcd/sdkjs/%2e%2e/secret')).toBeNull();
    expect(releaseIdFromReferrer('https://office-editor-a.getpi.work/r/v0.4.0-abcd/office-host.html')).toBe(
      'v0.4.0-abcd',
    );
    expect(releaseIdFromReferrer('not a url')).toBeNull();
    expect(resolveReleasePackageRequest('/p/v0.5.0-abcd/office-resources.oobpack')).toEqual({
      releaseId: 'v0.5.0-abcd',
    });
    expect(resolveReleasePackageRequest('/p/v0.5.0%2B1/office-resources.oobpack')).toEqual({
      releaseId: 'v0.5.0+1',
    });
    expect(resolveReleasePackageRequest('/p/../../secret/office-resources.oobpack')).toBeNull();
  });

  it('classifies explicit release assets by their inner path without nesting the release route', () => {
    const explicitHost = resolveEditorAssetRoute('/r/v0.4.0-abcd/office-host.html');
    expect(explicitHost).toEqual({
      releaseId: 'v0.4.0-abcd',
      pathname: '/office-host.html',
    });
    expect(shouldShareAsset(explicitHost.pathname, 'document')).toBe(false);
    const explicitSdk = resolveEditorAssetRoute('/r/v0.4.0-abcd/sdkjs/word/word.js');
    expect(explicitSdk).toEqual({
      releaseId: 'v0.4.0-abcd',
      pathname: '/sdkjs/word/word.js',
    });
    expect(shouldShareAsset(explicitSdk.pathname, 'script')).toBe(true);
    expect(resolveEditorAssetRoute('/office-host.html')).toEqual({
      releaseId: null,
      pathname: '/office-host.html',
    });
    expect(canonicalReleasePathname('/r/v0.4.0-abcd/sdkjs/word/word.js', 'v0.4.0-next')).toBe(
      '/r/v0.4.0-abcd/sdkjs/word/word.js',
    );
    expect(canonicalReleasePathname('/sdkjs/word/word.js', 'v0.4.0-next')).toBe('/r/v0.4.0-next/sdkjs/word/word.js');
    expect(canonicalReleasePathname('/sdkjs/slide/themes//themes.js', 'v0.4.0-next')).toBe(
      '/r/v0.4.0-next/sdkjs/slide/themes/themes.js',
    );
    expect(canonicalReleasePathname('/%252e%252e/%252e%252e/objects/other/sha256/abc', 'v0.4.0-next')).toBeNull();
  });

  it('prevents automatic analytics injection only in isolated Office HTML', () => {
    expect(shouldDisableResponseTransform('office-host.html', true)).toBe(true);
    expect(shouldDisableResponseTransform('web-apps/apps/documenteditor/main/index.html', true)).toBe(true);
    expect(shouldDisableResponseTransform('assets/officeHost.js', true)).toBe(false);
    expect(shouldDisableResponseTransform('index.html', false)).toBe(false);
  });

  it('accepts content revisions as immutable canonical cache keys', () => {
    expect(isAssetRevision('f59fbffe31d7f98f')).toBe(true);
    expect(isAssetRevision('f59fbffe31d7f98')).toBe(false);
    expect(isAssetRevision('release-v1')).toBe(false);
    expect(isAssetRevision(null)).toBe(false);
  });

  it('accepts only strictly content-addressed promotion receipt keys as immutable', () => {
    const candidate = 'a'.repeat(40);
    const sha256 = 'b'.repeat(64);
    expect(isImmutableReleaseReceiptKey(`promotions/v0.6.0+receipt/${candidate}-${sha256}.json`)).toBe(true);
    expect(isImmutableReleaseReceiptKey(`rollbacks/v0.6.0+receipt/${candidate}-${sha256}.json`)).toBe(true);
    expect(isImmutableReleaseReceiptKey(`promotions/v0.6.0+receipt/${candidate}-${sha256}.JSON`)).toBe(false);
    expect(isImmutableReleaseReceiptKey(`promotions/v0.6.0+receipt/${candidate.slice(1)}-${sha256}.json`)).toBe(false);
    expect(isImmutableReleaseReceiptKey(`promotions/v0.6.0+receipt/${candidate}-${sha256.slice(1)}.json`)).toBe(false);
    expect(isImmutableReleaseReceiptKey(`promotions/v0.6.0+receipt/${candidate}-${sha256}.json/extra`)).toBe(false);
    expect(isImmutableReleaseReceiptKey(`promotions/${candidate}-${sha256}.json`)).toBe(false);
  });

  it('keeps large Office payloads out of the Worker edge Cache API', () => {
    expect(
      shouldPopulateEdgeCache({ method: 'GET', immutable: true, hasRange: false, publicSize: 8 * 1024 * 1024 }),
    ).toBe(true);
    expect(
      shouldPopulateEdgeCache({ method: 'GET', immutable: true, hasRange: false, publicSize: 8 * 1024 * 1024 + 1 }),
    ).toBe(false);
    expect(
      shouldPopulateEdgeCache({ method: 'GET', immutable: true, hasRange: false, unboundedContentSegment: true }),
    ).toBe(false);
    expect(shouldPopulateEdgeCache({ method: 'GET', immutable: true, hasRange: true, publicSize: 1024 })).toBe(false);
  });

  it('serves content-addressed release blobs with their manifest MIME type', () => {
    const headers = new Headers({ 'Content-Type': 'application/octet-stream' });
    applyReleaseMime(headers, 'text/html; charset=utf-8');
    expect(headers.get('Content-Type')).toBe('text/html; charset=utf-8');
  });
});

describe('Cloudflare v5 release content object responses', () => {
  const releaseId = 'v5-worker-route';
  const sha256 = digest('7');
  const objectBytes = new TextEncoder().encode('release-bound-content-object');
  let storedResponses: Map<string, Response>;

  function objectMetadata(bytes: Uint8Array, includeBody: boolean) {
    return {
      ...(includeBody ? { body: new Response(bytes.slice().buffer).body! } : {}),
      httpEtag: `"${sha256}"`,
      size: bytes.byteLength,
      writeHttpMetadata(headers: Headers) {
        headers.set('Content-Type', 'application/octet-stream');
      },
    };
  }

  function testEnvironment(extraManifests: ReleaseManifest[] = []) {
    const manifest = v5Manifest(releaseId, { wholeSha256: sha256, bytes: objectBytes.byteLength });
    const packageSegment = new Uint8Array(manifest.package!.bytes);
    packageSegment.fill(0x5a);
    const objects = new Map<string, Uint8Array>([
      [`releases/${releaseId}/manifest.json`, new TextEncoder().encode(JSON.stringify(manifest))],
      [`blobs/sha256/${sha256}`, objectBytes],
      [`segments/sha256/${manifest.package!.sha256}`, packageSegment],
    ]);
    for (const extra of extraManifests) {
      objects.set(`releases/${extra.releaseId}/manifest.json`, new TextEncoder().encode(JSON.stringify(extra)));
    }
    const calls = {
      heads: [] as string[],
      gets: [] as Array<{ key: string; range: string | null }>,
      puts: [] as string[],
    };
    const env: WorkerEnv = {
      ASSET_VERSION: 'test-shell',
      LOCAL_MATRIX_MODE: '1',
      ASSETS: {
        async head(key) {
          calls.heads.push(key);
          const bytes = objects.get(key);
          return bytes ? objectMetadata(bytes, false) : null;
        },
        async get(key, options) {
          const bytes = objects.get(key);
          calls.gets.push({ key, range: options?.range?.get('range') || null });
          if (!bytes) return null;
          const range = options?.range?.get('range');
          if (!range) return objectMetadata(bytes, true);
          const match = /^bytes=(\d+)-(\d+)$/.exec(range);
          if (!match) return null;
          return objectMetadata(bytes.slice(Number(match[1]), Number(match[2]) + 1), true);
        },
        async put(key, value) {
          calls.puts.push(key);
          const bytes =
            typeof value === 'string'
              ? new TextEncoder().encode(value)
              : value instanceof ArrayBuffer
                ? new Uint8Array(value)
                : ArrayBuffer.isView(value)
                  ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength)
                  : value instanceof ReadableStream
                    ? new Uint8Array(await new Response(value).arrayBuffer())
                    : new Uint8Array();
          objects.set(key, bytes);
        },
      },
    };
    return { env, calls, objects };
  }

  async function fetchWorker(request: Request, env: WorkerEnv) {
    if (
      env.LOCAL_MATRIX_MODE === '1' &&
      !request.headers.has('x-onlyoffice-matrix-host') &&
      !new URL(request.url).hostname.endsWith('.localhost')
    ) {
      const headers = new Headers(request.headers);
      headers.set('X-OnlyOffice-Matrix-Host', 'onlyoffice.localhost');
      request = new Request(request, { headers });
    }
    const pending: Promise<unknown>[] = [];
    const response = await worker.fetch(request, env, {
      waitUntil(promise) {
        pending.push(promise);
      },
    });
    await Promise.all(pending);
    return response;
  }

  beforeEach(() => {
    storedResponses = new Map();
    vi.stubGlobal('caches', {
      default: {
        async match(request: Request) {
          return storedResponses.get(request.url)?.clone();
        },
        async put(request: Request, response: Response) {
          storedResponses.set(request.url, response.clone());
        },
      },
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('serves ordinary production assets from an immutable release root and exposes the exact Worker version', async () => {
    const { env, calls, objects } = testEnvironment();
    objects.set(
      'channels/stable-v5.json',
      new TextEncoder().encode(
        JSON.stringify({
          version: 1,
          releaseId,
          manifestUrl: `/releases/${releaseId}/manifest.json`,
          manifestSha256: digest('d'),
        }),
      ),
    );
    delete env.LOCAL_MATRIX_MODE;
    env.CF_VERSION_METADATA = {
      id: 'dc8dcd28-271b-4367-9840-6c244f84cb40',
      tag: 'candidate',
      timestamp: '2026-08-04T00:00:00.000Z',
    };

    const response = await fetchWorker(new Request('https://onlyoffice.getpi.work/sdkjs/word/sdk-all.js'), env);

    expect(response.status).toBe(200);
    const blobKey = calls.gets.find(({ key }) => key.startsWith('blobs/sha256/'))?.key;
    expect(blobKey).toBeTruthy();
    expect(Array.from(new Uint8Array(await response.arrayBuffer()))).toEqual(Array.from(objects.get(blobKey!)!));
    expect(calls.gets.map(({ key }) => key)).toEqual(
      expect.arrayContaining(['channels/stable-v5.json', `releases/${releaseId}/manifest.json`, blobKey]),
    );
    expect(response.headers.get('x-onlyoffice-worker-version')).toBe(env.CF_VERSION_METADATA.id);
    expect(response.headers.get('access-control-expose-headers')).toContain('X-OnlyOffice-Worker-Version');
    expect(response.headers.get('cache-control')).toBe('public, max-age=0, must-revalidate');
  });

  it('serves manifest-bound objects with immutable CORS, 200, 206, HEAD, and 416 semantics', async () => {
    const { env } = testEnvironment();
    const url = `https://onlyoffice.getpi.work/objects/${releaseId}/sha256/${sha256}`;

    const complete = await fetchWorker(new Request(url), env);
    expect(complete.status).toBe(200);
    expect(Array.from(new Uint8Array(await complete.arrayBuffer()))).toEqual(Array.from(objectBytes));
    expect(complete.headers.get('content-length')).toBe(String(objectBytes.byteLength));
    expect(complete.headers.get('content-type')).toBe('application/octet-stream');
    expect(complete.headers.get('accept-ranges')).toBe('bytes');
    expect(complete.headers.get('x-content-sha256')).toBe(sha256);
    expect(complete.headers.get('x-onlyoffice-asset-version')).toBe(releaseId);
    expect(complete.headers.get('cache-control')).toBe('public, max-age=31536000, immutable');
    expect(complete.headers.get('access-control-allow-origin')).toBe('*');
    expect(complete.headers.get('access-control-expose-headers')).toContain('Content-Range');
    expect(complete.headers.get('cross-origin-resource-policy')).toBe('cross-origin');

    const range = await fetchWorker(new Request(url, { headers: { Range: 'bytes=3-8' } }), env);
    expect(range.status).toBe(206);
    expect(Array.from(new Uint8Array(await range.arrayBuffer()))).toEqual(Array.from(objectBytes.slice(3, 9)));
    expect(range.headers.get('content-length')).toBe('6');
    expect(range.headers.get('content-range')).toBe(`bytes 3-8/${objectBytes.byteLength}`);

    const head = await fetchWorker(new Request(url, { method: 'HEAD' }), env);
    expect(head.status).toBe(200);
    expect(await head.text()).toBe('');
    expect(head.headers.get('content-length')).toBe(String(objectBytes.byteLength));
    expect(head.headers.get('accept-ranges')).toBe('bytes');

    const unsatisfiable = await fetchWorker(
      new Request(url, { headers: { Range: `bytes=${objectBytes.byteLength}-` } }),
      env,
    );
    expect(unsatisfiable.status).toBe(416);
    expect(unsatisfiable.headers.get('content-range')).toBe(`bytes */${objectBytes.byteLength}`);
    expect(unsatisfiable.headers.get('accept-ranges')).toBe('bytes');
    expect(unsatisfiable.headers.get('x-content-sha256')).toBe(sha256);

    const counters = await fetchWorker(new Request('https://onlyoffice.getpi.work/__matrix__/object-counters'), env);
    expect(await counters.json()).toMatchObject({
      [`${releaseId}:${sha256}`]: {
        workerRequests: 4,
        cacheHits: 0,
        r2Heads: 4,
        r2Gets: 2,
        declaredBytes: objectBytes.byteLength + 6,
        actualBytes: objectBytes.byteLength + 6,
        r2Bytes: objectBytes.byteLength + 6,
        completed: 4,
        aborted: 0,
        failed: 0,
        statuses: { '200': 2, '206': 1, '416': 1 },
      },
    });

    const routeCounters = await fetchWorker(
      new Request('https://onlyoffice.getpi.work/__matrix__/route-counters'),
      env,
    );
    expect(await routeCounters.json()).toMatchObject({
      [`GET /objects/${releaseId}/sha256/${sha256}`]: {
        workerRequests: 3,
        r2Gets: 2,
        actualBytes: objectBytes.byteLength + 6,
        statuses: { '200': 1, '206': 1, '416': 1 },
      },
      [`HEAD /objects/${releaseId}/sha256/${sha256}`]: {
        workerRequests: 1,
        r2Heads: 1,
        r2Gets: 0,
        actualBytes: 0,
        statuses: { '200': 1 },
      },
    });
  });

  it('serves only content-addressed promotion receipts with immutable caching', async () => {
    const { env, objects } = testEnvironment();
    const candidate = 'a'.repeat(40);
    const sha256 = 'b'.repeat(64);
    const validKey = `promotions/v0.6.0+receipt/${candidate}-${sha256}.json`;
    const invalidKey = `promotions/v0.6.0+receipt/${candidate.slice(1)}-${sha256}.json`;
    objects.set(validKey, new TextEncoder().encode('{"receipt":true}\n'));
    objects.set(invalidKey, new TextEncoder().encode('{"receipt":false}\n'));

    const valid = await fetchWorker(new Request(`https://onlyoffice.getpi.work/${validKey}`), env);
    expect(valid.status).toBe(200);
    expect(valid.headers.get('cache-control')).toBe('public, max-age=31536000, immutable');

    const invalid = await fetchWorker(new Request(`https://onlyoffice.getpi.work/${invalidKey}`), env);
    expect(invalid.status).toBe(200);
    expect(invalid.headers.get('cache-control')).toBe('public, max-age=0, must-revalidate');
  });

  it('keeps the production FixedLengthStream pipeline complete while counting R2 bytes', async () => {
    class TestFixedLengthStream extends TransformStream<Uint8Array, Uint8Array> {
      constructor(expectedBytes: number) {
        let receivedBytes = 0;
        super({
          transform(chunk, controller) {
            receivedBytes += chunk.byteLength;
            if (receivedBytes > expectedBytes) throw new Error('too many bytes');
            controller.enqueue(chunk);
          },
          flush() {
            if (receivedBytes !== expectedBytes) throw new Error('too few bytes');
          },
        });
      }
    }
    vi.stubGlobal('FixedLengthStream', TestFixedLengthStream);
    const { env } = testEnvironment();
    const url = `https://onlyoffice.getpi.work/objects/${releaseId}/sha256/${sha256}`;
    const beforeCounters = await fetchWorker(
      new Request('https://onlyoffice.getpi.work/__matrix__/object-counters'),
      env,
    );
    const before = ((await beforeCounters.json()) as Record<string, LocalMatrixCounter>)[`${releaseId}:${sha256}`] || {
      actualBytes: 0,
      r2Bytes: 0,
      completed: 0,
      aborted: 0,
      failed: 0,
    };
    const response = await fetchWorker(new Request(url), env);
    expect(response.status).toBe(200);
    expect(Array.from(new Uint8Array(await response.arrayBuffer()))).toEqual(Array.from(objectBytes));

    const counters = await fetchWorker(new Request('https://onlyoffice.getpi.work/__matrix__/object-counters'), env);
    const after = ((await counters.json()) as Record<string, LocalMatrixCounter>)[`${releaseId}:${sha256}`];
    expect(after.completed - before.completed).toBe(1);
    expect(after.aborted - before.aborted).toBe(0);
    expect(after.failed - before.failed).toBe(0);
    expect(after.actualBytes - before.actualBytes).toBe(objectBytes.byteLength);
    expect(after.r2Bytes - before.r2Bytes).toBe(objectBytes.byteLength);
  });

  it('switches the local v5 pointer only after token, identity, and manifest digest verification', async () => {
    const nextRelease = 'v5-next-release';
    const nextManifest = v5Manifest(nextRelease, { wholeSha256: digest('9'), bytes: 9 });
    const { env, calls, objects } = testEnvironment([nextManifest]);
    env.LOCAL_MATRIX_CONTROL_TOKEN = 'matrix-control-secret';
    const manifestBytes = objects.get(`releases/${nextRelease}/manifest.json`)!;
    const manifestSha256 = Array.from(
      new Uint8Array(await crypto.subtle.digest('SHA-256', manifestBytes.slice().buffer as ArrayBuffer)),
      (value) => value.toString(16).padStart(2, '0'),
    ).join('');
    const pointer = {
      version: 1,
      releaseId: nextRelease,
      manifestUrl: `/releases/${nextRelease}/manifest.json`,
      manifestSha256,
    };
    const endpoint = 'https://onlyoffice.getpi.work/__matrix__/stable-v5';

    const unauthorized = await fetchWorker(
      new Request(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(pointer),
      }),
      env,
    );
    expect(unauthorized.status).toBe(401);
    expect(calls.puts).toEqual([]);

    const switched = await fetchWorker(
      new Request(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-OnlyOffice-Matrix-Control-Token': 'matrix-control-secret',
        },
        body: JSON.stringify(pointer),
      }),
      env,
    );
    expect(switched.status).toBe(200);
    expect(await switched.json()).toEqual({ ok: true, releaseId: nextRelease });
    expect(calls.puts).toEqual(['channels/stable-v5.json']);
    expect(JSON.parse(new TextDecoder().decode(objects.get('channels/stable-v5.json')))).toEqual(pointer);

    const mismatched = await fetchWorker(
      new Request(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-OnlyOffice-Matrix-Control-Token': 'matrix-control-secret',
        },
        body: JSON.stringify({ ...pointer, manifestSha256: digest('0') }),
      }),
      env,
    );
    expect(mismatched.status).toBe(409);
    expect(calls.puts).toHaveLength(1);
  });

  it('does not read R2 for a digest absent from the requested release or for v4 manifests', async () => {
    const otherRelease = 'v5-without-object';
    const otherManifest = v5Manifest(otherRelease, { wholeSha256: digest('8'), bytes: 8 });
    const { env, calls } = testEnvironment([otherManifest]);

    const crossRelease = await fetchWorker(
      new Request(`https://onlyoffice.getpi.work/objects/${otherRelease}/sha256/${sha256}`),
      env,
    );
    expect(crossRelease.status).toBe(404);
    expect(await crossRelease.text()).toBe('Release content object not found');
    expect(calls.heads).toEqual([]);
    expect(calls.gets.some((call) => call.key === `blobs/sha256/${sha256}`)).toBe(false);

    const v4Release = 'v4-with-segment';
    const v4Manifest: ReleaseManifest = {
      version: 4,
      releaseId: v4Release,
      assets: [],
      package: {
        format: 'onlyoffice-pack-v1',
        path: 'office-resources.oobpack',
        bytes: objectBytes.byteLength,
        sha256,
        segments: [{ id: sha256, offset: 0, bytes: objectBytes.byteLength, sha256 }],
      },
    };
    const second = testEnvironment([v4Manifest]);
    const legacy = await fetchWorker(
      new Request(`https://onlyoffice.getpi.work/objects/${v4Release}/sha256/${sha256}`),
      second.env,
    );
    expect(legacy.status).toBe(404);
    expect(second.calls.heads).toEqual([]);
    expect(second.calls.gets.some((call) => call.key === `blobs/sha256/${sha256}`)).toBe(false);
  });

  it('redirects editor-origin object requests to the same release-bound canonical path', async () => {
    const { env } = testEnvironment();
    const pathname = `/objects/${releaseId}/sha256/${sha256}`;
    const response = await fetchWorker(new Request(`https://aries.localhost${pathname}`), env);
    expect(response.status).toBe(307);
    expect(response.headers.get('location')).toBe(`https://assets.office.localhost${pathname}`);
    expect(response.headers.get('access-control-allow-origin')).toBe('*');
    expect(response.headers.get('cross-origin-resource-policy')).toBe('cross-origin');
  });

  it('redirects editor-origin release metadata without nesting it below the editor release path', async () => {
    const { env } = testEnvironment();
    const pathname = `/releases/${releaseId}/manifest.json`;
    const response = await fetchWorker(new Request(`https://aries.localhost${pathname}`), env);
    expect(response.status).toBe(307);
    expect(response.headers.get('location')).toBe(`https://assets.office.localhost${pathname}`);
  });

  it('rejects multiply encoded traversal before constructing a canonical release redirect', async () => {
    const { env } = testEnvironment();
    const response = await fetchWorker(
      new Request('https://aries.localhost/%252e%252e/%252e%252e/objects/other/sha256/' + sha256, {
        headers: {
          Referer: `https://aries.getpi.work/r/${releaseId}/office-host.html`,
        },
      }),
      env,
    );
    expect(response.status).toBe(400);
    expect(response.headers.get('location')).toBeNull();
  });

  it('serves a manifest-bound package segment through the same release object API', async () => {
    const { env, calls } = testEnvironment();
    const packageSha256 = digest('e');
    const response = await fetchWorker(
      new Request(`https://onlyoffice.getpi.work/objects/${releaseId}/sha256/${packageSha256}`),
      env,
    );

    expect(response.status).toBe(200);
    expect(response.headers.get('x-content-sha256')).toBe(packageSha256);
    expect(response.headers.get('content-length')).toBe(String(objectBytes.byteLength + 16));
    expect(calls.gets).toContainEqual({
      key: `segments/sha256/${packageSha256}`,
      range: `bytes=0-${objectBytes.byteLength + 15}`,
    });
    expect(calls.gets.some((call) => call.key === `blobs/sha256/${packageSha256}`)).toBe(false);
  });

  it('sets distinct frame security policies for the installer, broker, and editor primer', async () => {
    const fixture = testEnvironment();
    fixture.env.LOCAL_MATRIX_PORT = '8787';
    const installerBytes = new TextEncoder().encode('<!doctype html><title>installer</title>');
    const brokerBytes = new TextEncoder().encode('<!doctype html><title>broker</title>');
    const officeHostSha256 = digest('7');
    const officeHostBytes = new TextEncoder().encode('<!doctype html><title>host</title>');
    const initialManifest = v5Manifest(releaseId);
    initialManifest.assets.push({
      path: 'office-host.html',
      sha256: officeHostSha256,
      mime: 'text/html; charset=utf-8',
      bytes: officeHostBytes.byteLength,
      profile: 'base',
      chunk: 'base-001',
      packageOffset: 0,
      representations: {
        whole: { sha256: officeHostSha256, bytes: officeHostBytes.byteLength },
      },
    });
    fixture.objects.set('resource-installer.html', installerBytes);
    fixture.objects.set('resource-broker.html', brokerBytes);
    fixture.objects.set(
      `releases/${releaseId}/manifest.json`,
      new TextEncoder().encode(JSON.stringify(initialManifest)),
    );
    fixture.objects.set(`blobs/sha256/${officeHostSha256}`, officeHostBytes);

    const installer = await fetchWorker(
      new Request('https://onlyoffice.localhost/resource-installer.html'),
      fixture.env,
    );
    const installerCsp = installer.headers.get('content-security-policy') || '';
    expect(installerCsp).toContain('frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work');
    expect(installerCsp).not.toContain(
      'frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work https://aries.getpi.work',
    );
    expect(installerCsp).toContain('frame-src https://aries.getpi.work');
    expect(installerCsp).toContain('https://host-gemini.office.localhost:8787');

    const broker = await fetchWorker(new Request('https://onlyoffice.localhost/resource-broker.html'), fixture.env);
    const brokerCsp = broker.headers.get('content-security-policy') || '';
    expect(brokerCsp).toContain(
      'frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work https://aries.getpi.work',
    );
    expect(brokerCsp).toContain('https://host-aries.office.localhost:8787');
    expect(brokerCsp).not.toContain('frame-src https://aries.getpi.work');

    const officeHost = await fetchWorker(
      new Request(`https://aries.localhost/r/${releaseId}/office-host.html`),
      fixture.env,
    );
    const officeHostCsp = officeHost.headers.get('content-security-policy') || '';
    expect(officeHost.headers.get('x-onlyoffice-matrix-host')).toBe('aries.localhost');
    expect(officeHost.headers.get('x-onlyoffice-matrix-isolated')).toBe('1');
    expect(officeHostCsp).toContain(
      'frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work https://piwork.localhost:8787 https://onlyoffice.localhost:8787',
    );
    expect(officeHostCsp).toContain("base-uri 'none'");
    expect(officeHostCsp).toContain("object-src 'none'");
    expect(officeHostCsp).not.toContain('https://evil.getpi.work');

    const primerSha256 = digest('9');
    const primerBytes = new Uint8Array(6).fill(0x09);
    const manifest = v5Manifest(releaseId);
    manifest.assets.push({
      path: 'editor-shell-prime.html',
      sha256: primerSha256,
      mime: 'text/html; charset=utf-8',
      bytes: primerBytes.byteLength,
      profile: 'base',
      chunk: 'base-001',
      packageOffset: 0,
      representations: {
        whole: { sha256: primerSha256, bytes: primerBytes.byteLength },
      },
    });
    fixture.objects.set(`releases/${releaseId}/manifest.json`, new TextEncoder().encode(JSON.stringify(manifest)));
    fixture.objects.set(`blobs/sha256/${primerSha256}`, primerBytes);
    storedResponses.clear();
    const primer = await fetchWorker(
      new Request(`https://aries.localhost/r/${releaseId}/editor-shell-prime.html`),
      fixture.env,
    );
    expect(primer.status).toBe(200);
    expect(primer.headers.get('location')).toBeNull();
    const primerCsp = primer.headers.get('content-security-policy') || '';
    expect(primerCsp).toContain('frame-src https://onlyoffice.getpi.work');
    expect(primerCsp).toContain('frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work');
    expect(primerCsp).not.toContain(
      'frame-ancestors https://piwork.getpi.work https://onlyoffice.getpi.work https://aries.getpi.work',
    );
  });
});
