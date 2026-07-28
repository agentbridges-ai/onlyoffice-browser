import { describe, expect, it } from 'vitest';
import {
  isIsolatedEditorHost,
  isFixedOfflineEditorHost,
  isOnlyOfficeHost,
  resolveObjectKey,
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
    expect(isFixedOfflineEditorHost('office-misaka.getpi.work')).toBe(true);
    expect(isFixedOfflineEditorHost('office-pectics.getpi.work')).toBe(true);
    expect(isFixedOfflineEditorHost('office-editor-a.getpi.work')).toBe(false);
  });

  it('keeps only origin-bound boot files and workers on each editor origin', () => {
    expect(shouldShareAsset('/office-host.html', 'document')).toBe(false);
    expect(shouldShareAsset('/assets/officeHost-a.js', 'script')).toBe(false);
    expect(shouldShareAsset('/wasm/x2t/worker.js', 'worker')).toBe(false);
    expect(shouldShareAsset('/document_editor_service_worker.js', 'serviceworker')).toBe(false);
    expect(shouldShareAsset('/sdkjs/word/sdk-all.js', 'script')).toBe(true);
    expect(shouldShareAsset('/fonts/000.ttf', 'empty')).toBe(true);
  });

  it('maps the public root and rejects traversal keys', () => {
    expect(resolveObjectKey('/')).toBe('index.html');
    expect(resolveObjectKey('/sdkjs/word/sdk-all.js')).toBe('sdkjs/word/sdk-all.js');
    expect(resolveObjectKey('/sdkjs/%2e%2e/secret')).toBeNull();
    expect(resolveObjectKey('/%E0%A4%A')).toBeNull();
  });
});
