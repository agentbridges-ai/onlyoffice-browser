import { describe, expect, it, vi } from 'vitest';
import { clearLegacyDemoHostState, resolveDemoHostUrl } from '../../src/lib/demo-host-url';
import type { OfficeEditorOriginSlot } from '../../src/lib/office-origin-pool';

const context = (sessionId: string, hostSlot: OfficeEditorOriginSlot = 'aries') => ({
  sessionId,
  hostSlot,
  fileName: 'New_Document.docx',
  fileType: 'docx',
  mode: 'edit' as const,
});

function resolve(url: URL, sessionId: string, hostSlot: OfficeEditorOriginSlot = 'aries'): string {
  const resolver = resolveDemoHostUrl(url);
  return String(typeof resolver === 'function' ? resolver(context(sessionId, hostSlot)) : resolver);
}

describe('demo host URL routing', () => {
  it('maps production demo editors onto the fixed constellation origin pool', () => {
    const page = new URL('https://onlyoffice.getpi.work/');

    expect(new URL(resolve(page, 'office-editor-a', 'aries')).origin).toBe('https://aries.getpi.work');
    expect(new URL(resolve(page, 'office-editor-b', 'taurus')).origin).toBe('https://taurus.getpi.work');
  });

  it('migrates the legacy fixed R2 host to production wildcard origins', () => {
    const page = new URL(
      'https://onlyoffice.getpi.work/?hostUrl=https%3A%2F%2Fpub-7144f7712bc5465b880405ca7741e61f.r2.dev%2Foffice-host.html',
    );

    expect(new URL(resolve(page, 'office-editor-c', 'gemini')).origin).toBe('https://gemini.getpi.work');
  });

  it('supports an explicit session template for non-production deployments', () => {
    const page = new URL(
      'https://demo.example.com/?hostUrl=https%3A%2F%2Foffice-%7BsessionId%7D.example.com%2Foffice-host.html',
    );

    expect(resolve(page, 'editor-1')).toBe('https://office-editor-1.example.com/office-host.html');
  });

  it('supports an explicit fixed-pool slot template for non-production deployments', () => {
    const page = new URL(
      'https://demo.example.com/?hostUrl=https%3A%2F%2F%7BhostSlot%7D.office.example.com%2Foffice-host.html',
    );

    expect(resolve(page, 'editor-1', 'libra')).toBe('https://libra.office.example.com/office-host.html');
  });

  it('removes stale production query and storage state', () => {
    const replaceState = vi.spyOn(window.history, 'replaceState').mockImplementation(() => {});
    const storage = { removeItem: vi.fn() } as unknown as Storage;
    const location = {
      hostname: 'onlyoffice.getpi.work',
      href: 'https://onlyoffice.getpi.work/?hostUrl=https%3A%2F%2Fpub-7144f7712bc5465b880405ca7741e61f.r2.dev%2Foffice-host.html',
    } as Location;

    clearLegacyDemoHostState(location, storage);

    expect(storage.removeItem).toHaveBeenCalledWith('onlyoffice-browser:last-host-url');
    expect(replaceState).toHaveBeenCalledWith(window.history.state, '', new URL('https://onlyoffice.getpi.work/'));
  });
});
