/**
 * Tests for the fetch routing rules in public/sw.js.
 *
 * sw.js is a non-module service worker file that can't be imported directly,
 * so we replicate the routing conditions here as a living specification.
 * If sw.js changes, update both files together.
 *
 * The rules guard against two classes of bug found in this project:
 *   - Font files intercepted by SW → added latency → Chrome "Slow Network"
 *     intervention → OnlyOffice fallback font crash (units_per_EM)
 *   - Document URLs cached by SW → stale content served to editor
 */

import { describe, expect, it } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';

const FONT_REGEX = /\.(ttf|tte|ttc|otf|otc|woff2?|eot)(\?.*)?$/;
const ONLYOFFICE_RUNTIME_ASSET_REGEX = /(^|\/)(web-apps|sdkjs|wasm\/x2t)\//;
const PRINT_PDF_ROUTE_PREFIX = '/__onlyoffice-browser-print__/';
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
  return true;
}

function swStaticStrategy(urlStr: string): 'network-first' | 'stale-while-revalidate' {
  const url = new URL(urlStr);
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
  it('exposes the OnlyOffice service worker at the root path expected by editor frames', () => {
    const bridgePath = path.join(process.cwd(), 'public/document_editor_service_worker.js');
    const bridge = fs.readFileSync(bridgePath, 'utf8');

    expect(bridge).toContain("importScripts('/sw.js')");
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

    it.each(['bytes=100-99', 'bytes=1000-', 'items=0-10', 'bytes=-0', 'bytes=-'])('rejects invalid range %s', (range) => {
      expect(parseRangeHeader(range, 1000)).toBeNull();
    });
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

  describe('font files are not intercepted (crash prevention)', () => {
    // Intercepting font files adds SW latency which triggers Chrome's
    // "Slow Network" font-loading intervention. OnlyOffice can then crash with
    // "Cannot read properties of undefined (reading 'units_per_EM')"
    // in the fallback font code path of slide/word/cell sdk-all.js.
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
      `${ORIGIN}/web-apps/apps/api/documents/api.js`,
      `${ORIGIN}/public/sdkjs/slide/sdk-all.js`,
      `${ORIGIN}/styles/base.css`,
      `${ORIGIN}/plugins.json`,
      `${ORIGIN}/themes.json`,
    ])('%s', (url) => {
      expect(swShouldHandle('GET', url)).toBe(true);
    });
  });

  describe('OnlyOffice runtime assets use network-first freshness', () => {
    it.each([
      `${ORIGIN}/web-apps/apps/api/documents/api.js`,
      `${ORIGIN}/sdkjs/word/sdk-all.js`,
      `${ORIGIN}/wasm/x2t/x2t.wasm`,
      `${ORIGIN}/document/web-apps/apps/documenteditor/main/app.js`,
    ])('%s', (url) => {
      expect(swStaticStrategy(url)).toBe('network-first');
    });

    it('keeps ordinary static assets on stale-while-revalidate', () => {
      expect(swStaticStrategy(`${ORIGIN}/styles/base.css`)).toBe('stale-while-revalidate');
    });
  });
});
