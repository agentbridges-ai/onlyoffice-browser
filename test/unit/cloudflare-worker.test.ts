import { describe, expect, it } from 'vitest';
import {
  applyReleaseMime,
  canonicalReleasePathname,
  isIsolatedEditorHost,
  isAssetRevision,
  isOnlyOfficeHost,
  resolveEditorAssetRoute,
  resolveObjectKey,
  resolveReleaseRequest,
  releaseIdFromReferrer,
  resolveRuntimeHost,
  shouldDisableResponseTransform,
  shouldShareAsset,
} from '../../cloudflare/worker';

describe('Cloudflare OnlyOffice runtime routing', () => {
  it('accepts the canonical and unbounded per-editor wildcard hosts', () => {
    expect(isOnlyOfficeHost('onlyoffice.getpi.work')).toBe(true);
    expect(isOnlyOfficeHost('office-editor-a.getpi.work')).toBe(true);
    expect(isOnlyOfficeHost('unrelated.getpi.work')).toBe(false);
    expect(isOnlyOfficeHost('office-a.dev.getpi.work')).toBe(false);
    expect(isOnlyOfficeHost('onlyoffice.getpi.work.example.com')).toBe(false);
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

  it('keeps only origin-bound boot files and workers on each editor origin', () => {
    expect(shouldShareAsset('/office-host.html', 'document')).toBe(false);
    expect(shouldShareAsset('/assets/officeHost-a.js', 'script')).toBe(false);
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

  it('maps the public root and rejects traversal keys', () => {
    expect(resolveObjectKey('/')).toBe('index.html');
    expect(resolveObjectKey('/sdkjs/word/sdk-all.js')).toBe('sdkjs/word/sdk-all.js');
    expect(resolveObjectKey('/sdkjs/slide/themes//themes.js')).toBe('sdkjs/slide/themes/themes.js');
    expect(resolveObjectKey('/sdkjs/%2e%2e/secret')).toBeNull();
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

  it('serves content-addressed release blobs with their manifest MIME type', () => {
    const headers = new Headers({ 'Content-Type': 'application/octet-stream' });
    applyReleaseMime(headers, 'text/html; charset=utf-8');
    expect(headers.get('Content-Type')).toBe('text/html; charset=utf-8');
  });
});
