/**
 * Tests for the fetch routing rules in src/service-worker.js.
 *
 * The service worker is bundled with Workbox and can't be imported directly,
 * so we replicate the routing conditions here as a living specification.
 * If sw.js changes, update both files together.
 *
 * The rules guard against Office resources leaking into generic Workbox caches
 * and document URLs being cached as stale editor input. Isolated editor origins
 * use the earlier streaming canonical-broker route instead.
 */

import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';

const FONT_REGEX = /\.(ttf|tte|ttc|otf|otc|woff2?|eot)(\?.*)?$/;
const ONLYOFFICE_RUNTIME_ASSET_REGEX = /(^|\/)(web-apps|sdkjs|wasm\/x2t)\//;
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
const CONTENT_SEGMENT_PATH_REGEX = /^\/segments\/sha256\/[a-f0-9]{64}$/;
const CONTENT_OBJECT_PATH_REGEX = /^\/objects\/[^/]+\/sha256\/[a-f0-9]{64}$/;
const CONTENT_BLOB_PATH_REGEX = /^\/blobs\/sha256\/[a-f0-9]{64}$/;
const ONLYOFFICE_NAVIGATION_PATHS = new Set(['/office-host.html', '/reset.html']);

const ORIGIN = 'http://localhost:5173';

function swShouldHandle(method: string, urlStr: string, mode = 'same-origin'): boolean {
  const url = new URL(urlStr);
  if (url.origin !== ORIGIN) return false;
  if (url.pathname.startsWith(PRINT_PDF_ROUTE_PREFIX)) return method === 'GET' || method === 'HEAD';
  if (method !== 'GET') return false;
  if (mode === 'navigate' && !ONLYOFFICE_NAVIGATION_PATHS.has(url.pathname)) return false;
  if (
    url.pathname.startsWith('/api/') ||
    url.pathname.startsWith('/internal/') ||
    url.pathname.startsWith('/ws/') ||
    url.pathname.startsWith('/@vite/') ||
    url.pathname.startsWith('/@react-refresh') ||
    url.pathname.startsWith('/@id/') ||
    url.pathname.startsWith('/@fs/') ||
    url.pathname.startsWith('/node_modules/') ||
    url.pathname.startsWith('/src/')
  ) {
    return false;
  }
  if (url.searchParams.has('file') || url.searchParams.has('src')) return false;
  if (url.pathname.startsWith('/fonts/') || FONT_REGEX.test(url.pathname)) return false;
  if (CONTENT_SEGMENT_PATH_REGEX.test(url.pathname)) return false;
  if (CONTENT_OBJECT_PATH_REGEX.test(url.pathname)) return false;
  if (CONTENT_BLOB_PATH_REGEX.test(url.pathname)) return false;
  if (url.pathname.startsWith('/channels/') || url.pathname.startsWith('/releases/')) return false;
  if (/^\/r\/[^/]+\/.+/.test(url.pathname)) return false;
  if (ONLYOFFICE_RUNTIME_ASSET_REGEX.test(url.pathname)) return false;
  if (url.pathname === '/onlyoffice-runtime-assets.json') return false;
  return true;
}

function swStaticStrategy(urlStr: string): 'none' | 'network-first' | 'stale-while-revalidate' {
  const url = new URL(urlStr);
  if (ONLYOFFICE_RUNTIME_ASSET_REGEX.test(url.pathname)) return 'none';
  const isHtml = url.pathname.endsWith('.html') || url.pathname === '/' || url.pathname.endsWith('/');
  return isHtml || ONLYOFFICE_RUNTIME_ASSET_REGEX.test(url.pathname) ? 'network-first' : 'stale-while-revalidate';
}

function parseRangeHeader(rangeHeader: string, byteLength: number): { start: number; end: number } | null {
  const match = /^bytes=(\d*)-(\d*)$/.exec(rangeHeader || '');
  if (!match || byteLength <= 0) return null;

  let start: number;
  let end: number;
  if (match[1] === '' && match[2] === '') return null;
  if (match[1] === '') {
    const suffixLength = Number(match[2]);
    if (!Number.isFinite(suffixLength) || suffixLength <= 0) return null;
    start = Math.max(byteLength - suffixLength, 0);
    end = byteLength - 1;
  } else {
    start = Number(match[1]);
    end = match[2] === '' ? byteLength - 1 : Number(match[2]);
  }

  if (!Number.isFinite(start) || !Number.isFinite(end) || start < 0 || end < start || start >= byteLength) {
    return null;
  }

  return {
    start,
    end: Math.min(end, byteLength - 1),
  };
}

describe('SW fetch routing', () => {
  it('does not await a navigation resultingClientId before returning its Broker response', () => {
    const sw = fs.readFileSync(path.join(process.cwd(), 'src/service-worker.js'), 'utf8');
    expect(sw).toContain('const sourceClient = event.clientId ? await self.clients.get(event.clientId) : null;');
    expect(sw).not.toContain('await self.clients.get(clientId)');
    expect(sw).toContain("event.request.mode === 'navigate'");
    expect(sw).toContain("event.request.destination === 'iframe'");
  });

  it('exposes the OnlyOffice service worker at the root path expected by editor frames', () => {
    const bridgePath = path.join(process.cwd(), 'public/document_editor_service_worker.js');
    const bridge = fs.readFileSync(bridgePath, 'utf8');

    expect(bridge).toContain("importScripts('/sw.js')");
  });

  it('precaches the canonical PWA shell without fixed offline editor slots', () => {
    const sw = fs.readFileSync(path.join(process.cwd(), 'src/service-worker.js'), 'utf8');

    expect(sw).toContain("const PWA_APP_NAVIGATION_PATHS = new Set(['/', '/index.html', '/resource-broker.html'])");
    expect(sw).toContain('const pwaShellManifest = self.__WB_MANIFEST || []');
    expect(sw).toContain('precacheAndRoute(pwaShellManifest)');
    expect(sw).toContain("matchPrecache('/index.html')");
    expect(sw).toContain("url.pathname === '/resource-broker.html'");
    expect(sw).toContain("matchPrecache('/resource-broker.html')");
    expect(sw).toContain("url.pathname === '/channels/stable-v5.json'");
    expect(sw).toContain("JSON.stringify({ code: 'offline' })");
    expect(sw).toContain("'cache-control': 'no-store'");
    expect(sw).toContain('isCanonicalPwaHost');
    expect(sw).not.toContain('FIXED_OFFLINE_SLOT');
    expect(sw).not.toContain('onlyoffice-slot-prewarm-v1');
    expect(sw).toContain('if (isIsolatedEditorHost)');
    expect(sw).toContain('clientsClaim()');
    expect(sw).not.toContain("self.addEventListener('install'");
  });

  it('activates a waiting shell update only after every tab is safe to restore', () => {
    const source = fs.readFileSync(path.join(process.cwd(), 'src/index.ts'), 'utf8');

    expect(source).toContain("workbox.addEventListener('waiting'");
    expect(source).toContain('PREPARE_UPDATE');
    expect(source).toContain('allTabsSafeForUpdate()');
    expect(source).toContain('await waitingWorkbox.messageSkipWaiting()');
    expect(source).toContain("workbox.addEventListener('controlling'");
    expect(source).toContain('if (reloading || hasUnsafeWork()) return');
    expect(source).toContain('onlyoffice-browser:pwa-reload:');
    expect(source).toContain('location.reload()');
    expect(source).toContain('tabs.some((tab) => tab.dirty || !tab.handle)');
    expect(source).toContain('10 * 60 * 1000');
  });

  it('serves isolated-host Office resources only through the canonical broker', () => {
    const sw = fs.readFileSync(path.join(process.cwd(), 'src/service-worker.js'), 'utf8');

    expect(sw).toContain("const CANONICAL_OFFICE_HOST = 'onlyoffice.getpi.work'");
    expect(sw).toContain("const LOCAL_CANONICAL_OFFICE_HOST = 'onlyoffice.localhost'");
    expect(sw).toContain('isProductionOfficeEditorHostname(self.location.hostname)');
    expect(sw).toContain('LEGACY_EDITOR_HOST_PATTERN.test(self.location.hostname)');
    expect(sw).toContain('new EditorResourceBrokerClient({');
    expect(sw).toContain('parseEditorOfficeHostClientIdentity');
    expect(sw).toContain('new EditorClientIdentityRegistry()');
    expect(sw).toContain('editorClientIdentities.resolve({');
    expect(sw).toContain('resultingClientId: event.resultingClientId || undefined');
    expect(sw).toContain('includeUncontrolled: true');
    expect(sw).toContain('editorClientIdentities?.canBindHost(sourceClientId, bind, liveHosts)');
    expect(sw).toContain('editorClientIdentities?.canBindPrime(bind, liveHosts)');
    expect(sw).toContain('editorClientIdentities.bindHost(sourceClientId, bind)');
    expect(sw).toContain('const bindTask = editorBrokerBindQueue.then(async () => {');
    expect(sw).toContain('editorBrokerBindQueue = bindTask.catch(() => {');
    expect(sw).toContain('releaseIdFromOfficeHostUrl(url)');
    expect(sw).toContain('sourceIdentity.releaseId === bind.releaseId');
    expect(sw).toContain('sourceIdentity.sessionId === bind.sessionId');
    expect(sw).toContain('editorResourceBroker.handleBindMessage(bind, ports, commitHostIdentity)');
    expect(sw).toContain('if (hostAuthorized) editorClientIdentities.bindHost(sourceClientId, bind)');
    expect(sw).toContain('parseEditorResourceBrokerUnbindMessage(event.data)');
    expect(sw).toContain('editorClientIdentities?.unbindHost(sourceClientId, unbind)');
    expect(sw).toContain('if (trusted) editorResourceBroker?.disconnect()');
    expect(sw).toContain("type: 'ONLYOFFICE_BROKER_NEEDED'");
    expect(sw).toContain('candidate?.releaseId === identity.releaseId');
    expect(sw).toContain('candidate.sessionId === identity.sessionId');
    expect(sw).toContain('isEditorOfficeRuntimeClientUrl');
    expect(sw).not.toContain('AbortSignal.timeout(EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS)');
    expect(sw).not.toContain('EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS - (Date.now() - startedAt)');
    expect(sw).toContain('editorResourceBroker.fetchAsset(request, path, {');
    expect(sw).toContain("Object.defineProperty(self, '__ONLYOFFICE_BROKER_ACCEPTANCE_METRICS__'");
    expect(sw).toContain('canonical: canonicalResourceBroker?.metrics ?? null');
    expect(sw).toContain('editor: editorResourceBroker?.metrics ?? null');
    expect(sw).toContain('connectionTimeoutMs: EDITOR_RESOURCE_BROKER_REQUEST_TIMEOUT_MS');
    expect(sw).toContain("error.stage === 'connection'");
    expect(sw).toContain("error.code === 'connection' || error.code === 'replaced'");
    expect(sw).toContain('for (let attempt = 0; attempt < 2; attempt += 1)');
    expect(sw).toContain("registerRoute(editorBrokerRouteMatcher, editorBrokerRouteHandler, 'GET')");
    expect(sw).toContain("registerRoute(editorBrokerRouteMatcher, editorBrokerRouteHandler, 'HEAD')");
    expect(sw).toContain('matchEditorShell(request, releaseAsset.releaseId, caches)');
    expect(sw).toContain('editorBootstrapAssetPaths.has(`/${releaseAsset.path}`)');
    expect(sw).toContain('return fetchEditorBrokerAsset(event, request, url)');
    expect(sw).toContain('status: 503');
    expect(sw).toContain('!isIsolatedEditorHost &&');
    expect(sw).not.toContain('loadCurrentRelease');
    expect(sw).not.toContain('fetchOfficePackAsset');
    expect(sw).not.toContain('loadOfficePackSegment');
    expect(sw).not.toContain('fetchSharedAsset');
    expect(sw).not.toContain("cache: 'force-cache'");
    expect(sw).not.toContain('return fetch(request)');
    expect(sw).toContain('if (isCanonicalPwaHost)');
    expect(sw).toContain('new NetworkFirst');
    expect(sw).toContain('new StaleWhileRevalidate');
    expect(sw).toContain('url.searchParams.has(SHARED_ASSET_VERSION_QUERY)');
    expect(sw).toContain("event.data?.type === 'SET_FONT_ALLOWLIST'");
    expect(sw).toContain('Retain the v2 message contract for older clients');
    expect(sw).not.toContain('downloadedFontPaths = new Set(');
    expect(sw).not.toContain('buildAllFontsMetadataFallbackBootstrap');
    expect(sw).toContain('isContentSegmentRequest(url)');
    expect(sw).toContain("event.data?.type === 'ONLYOFFICE_VERIFY_EDITOR_SHELL'");
    expect(sw).toContain('verifyEditorShell({');
    expect(sw).not.toContain("responseHeaders.set('cache-control', 'no-store')");
    expect(sw).not.toContain("return new Response('Font is not installed'");
    expect(sw).not.toContain('selectFallbackFont');
  });

  it('keeps immutable content endpoints outside the generic Workbox caches', () => {
    const digest = 'a'.repeat(64);
    expect(swShouldHandle('GET', `${ORIGIN}/segments/sha256/${digest}`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/objects/release-a/sha256/${digest}`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/blobs/sha256/${digest}`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/channels/stable-v5.json`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/releases/release-a/manifest.json`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/r/release-a/sdkjs/word/word.js`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/sdkjs/word/word.js`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/wasm/x2t/x2t.wasm`)).toBe(false);
    expect(swShouldHandle('GET', `${ORIGIN}/segments/sha256/not-a-digest`)).toBe(true);

    const sw = fs.readFileSync(path.join(process.cwd(), 'src/service-worker.js'), 'utf8');
    expect(sw).toContain('isCanonicalManagedResourceRequest(url)');
    expect(sw).toContain('CONTENT_OBJECT_PATH_REGEX.test(url.pathname)');
    expect(sw).toContain("url.pathname.startsWith('/channels/')");
    expect(sw).toContain("url.pathname.startsWith('/releases/')");
  });

  it('provides root OnlyOffice desktop-mode discovery manifests', () => {
    const plugins = JSON.parse(fs.readFileSync(path.join(process.cwd(), 'public/plugins.json'), 'utf8'));
    const themes = JSON.parse(fs.readFileSync(path.join(process.cwd(), 'public/themes.json'), 'utf8'));

    expect(plugins).toEqual({ pluginsData: [] });
    expect(themes).toEqual({ themes: [] });
  });

  describe('non-GET requests are not handled', () => {
    it.each(['POST', 'PUT', 'DELETE', 'PATCH'])('%s', (method) => {
      expect(swShouldHandle(method, `${ORIGIN}/index.html`)).toBe(false);
    });
  });

  describe('cross-origin requests are not handled', () => {
    it('skips external document URL', () => {
      expect(swShouldHandle('GET', 'https://example.com/doc.docx')).toBe(false);
    });

    it('skips CDN asset', () => {
      expect(swShouldHandle('GET', 'https://cdn.example.com/font.ttf')).toBe(false);
    });
  });

  describe('document query params bypass the SW cache', () => {
    it('skips ?src= URLs', () => {
      expect(swShouldHandle('GET', `${ORIGIN}/?src=https://example.com/doc.docx`)).toBe(false);
    });

    it('skips ?file= URLs', () => {
      expect(swShouldHandle('GET', `${ORIGIN}/?file=report.xlsx`)).toBe(false);
    });

    it('skips URL with both src and other params', () => {
      expect(swShouldHandle('GET', `${ORIGIN}/?src=doc.docx&readonly=true`)).toBe(false);
    });
  });

  describe('temporary print PDFs are handled by the SW cache', () => {
    it('serves same-origin print PDFs before generic static routing', () => {
      expect(swShouldHandle('GET', `${ORIGIN}/__onlyoffice-browser-print__/print-1.pdf`)).toBe(true);
    });

    it('handles HEAD probes for Chrome PDF viewer print PDFs', () => {
      expect(swShouldHandle('HEAD', `${ORIGIN}/__onlyoffice-browser-print__/print-1.pdf`)).toBe(true);
      expect(swShouldHandle('HEAD', `${ORIGIN}/office-host.html`)).toBe(false);
    });

    it.each([
      ['bytes=0-99', 1000, { start: 0, end: 99 }],
      ['bytes=200-', 1000, { start: 200, end: 999 }],
      ['bytes=-250', 1000, { start: 750, end: 999 }],
      ['bytes=950-1200', 1000, { start: 950, end: 999 }],
    ])('parses PDF range request %s', (range, byteLength, expected) => {
      expect(parseRangeHeader(range, byteLength)).toEqual(expected);
    });

    it.each(['bytes=100-99', 'bytes=1000-', 'items=0-10', 'bytes=-0', 'bytes=-'])(
      'rejects invalid range %s',
      (range) => {
        expect(parseRangeHeader(range, 1000)).toBeNull();
      },
    );
  });

  describe('host navigation and app routes are not intercepted', () => {
    it.each([
      [`${ORIGIN}/`, 'root navigation'],
      [`${ORIGIN}/index.html`, 'demo index navigation'],
      [`${ORIGIN}/api/me`, 'API'],
      [`${ORIGIN}/internal/user-space-transfer/session/file`, 'internal API'],
      [`${ORIGIN}/ws/browser/session`, 'websocket route'],
      [`${ORIGIN}/@vite/client`, 'Vite client'],
      [`${ORIGIN}/@react-refresh`, 'Vite React refresh runtime'],
      [`${ORIGIN}/@id/react`, 'Vite module id'],
      [`${ORIGIN}/@fs/Users/xy/Documents/Nexolyra/web/src/main.tsx`, 'Vite fs module'],
      [`${ORIGIN}/node_modules/.vite/deps/@agentbridges-ai_onlyoffice-browser.js`, 'Vite optimized dependency'],
      [`${ORIGIN}/src/main.tsx`, 'Vite source module'],
    ])('%s (%s)', (url) => {
      expect(swShouldHandle('GET', url, 'navigate')).toBe(false);
    });

    it.each([`${ORIGIN}/office-host.html`, `${ORIGIN}/reset.html`])('allows Office host navigation %s', (url) => {
      expect(swShouldHandle('GET', url, 'navigate')).toBe(true);
    });

    it.each([
      `${ORIGIN}/@vite/client`,
      `${ORIGIN}/@react-refresh`,
      `${ORIGIN}/@id/react`,
      `${ORIGIN}/@fs/Users/xy/Documents/Nexolyra/web/src/main.tsx`,
      `${ORIGIN}/node_modules/.vite/deps/@agentbridges-ai_onlyoffice-browser.js`,
      `${ORIGIN}/src/main.tsx`,
    ])('does not intercept Vite development subresource %s', (url) => {
      expect(swShouldHandle('GET', url)).toBe(false);
    });
  });

  describe('font files stay outside the generic Workbox caches', () => {
    // Isolated editor origins use the earlier canonical-broker route. The
    // generic runtime/static caches must never create a second font copy.
    it.each([
      ['/web-apps/apps/common/main/resources/font/ASC.ttf', '.ttf (OnlyOffice internal font)'],
      ['/fonts/NotoSansTC-VF.ttf', '.ttf (CJK fallback font)'],
      ['/fonts/LiberationSans-Bold.woff2', '.woff2'],
      ['/fonts/arial.woff', '.woff'],
      ['/fonts/symbol.otf', '.otf'],
      ['/fonts/msyh.ttc', '.ttc font collection'],
      ['/fonts/cambria.otc', '.otc font collection'],
      ['/fonts/embedded.tte', '.tte embedded TrueType'],
      ['/fonts/legacy.eot', '.eot'],
      ['/fonts/000', 'official generated font without extension'],
      ['/fonts/font.ttf?v=123', '.ttf with query string'],
    ])('%s (%s)', (pathname) => {
      expect(swShouldHandle('GET', `${ORIGIN}${pathname}`)).toBe(false);
    });
  });

  describe('font regex matches extensions correctly', () => {
    it.each(['.ttf', '.tte', '.ttc', '.otf', '.otc', '.woff', '.woff2', '.eot'])('matches %s', (ext) => {
      expect(FONT_REGEX.test(`/fonts/file${ext}`)).toBe(true);
    });

    it('does not match .ttfx', () => {
      expect(FONT_REGEX.test('/fonts/file.ttfx')).toBe(false);
    });

    it('does not match .js or .css', () => {
      expect(FONT_REGEX.test('/sdk-all.js')).toBe(false);
      expect(FONT_REGEX.test('/styles.css')).toBe(false);
    });

    it('matches font extensions embedded in longer paths', () => {
      expect(FONT_REGEX.test('/web-apps/apps/common/main/resources/font/ASC.ttf')).toBe(true);
    });
  });

  describe('same-origin static assets are handled', () => {
    it.each([
      `${ORIGIN}/office-host.html`,
      `${ORIGIN}/reset.html`,
      `${ORIGIN}/styles/base.css`,
      `${ORIGIN}/plugins.json`,
      `${ORIGIN}/themes.json`,
    ])('%s', (url) => {
      expect(swShouldHandle('GET', url)).toBe(true);
    });

    it.each([`${ORIGIN}/web-apps/apps/api/documents/api.js`, `${ORIGIN}/public/sdkjs/slide/sdk-all.js`])(
      'keeps Office runtime outside Workbox: %s',
      (url) => {
        expect(swShouldHandle('GET', url)).toBe(false);
      },
    );
  });

  describe('OnlyOffice runtime assets stay outside generic Workbox strategies', () => {
    it.each([
      `${ORIGIN}/web-apps/apps/api/documents/api.js`,
      `${ORIGIN}/sdkjs/word/sdk-all.js`,
      `${ORIGIN}/wasm/x2t/x2t.wasm`,
      `${ORIGIN}/document/web-apps/apps/documenteditor/main/app.js`,
    ])('%s', (url) => {
      expect(swStaticStrategy(url)).toBe('none');
    });

    it('keeps ordinary static assets on stale-while-revalidate', () => {
      expect(swStaticStrategy(`${ORIGIN}/styles/base.css`)).toBe('stale-while-revalidate');
    });
  });
});
