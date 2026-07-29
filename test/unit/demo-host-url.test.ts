import { describe, expect, it, vi } from 'vitest';
import { clearLegacyDemoHostState, resolveDemoHostUrl } from '../../src/lib/demo-host-url';

const context = (sessionId: string) => ({
  sessionId,
  fileName: 'New_Document.docx',
  fileType: 'docx',
  mode: 'edit' as const,
});

function resolve(url: URL, sessionId: string): string {
  const resolver = resolveDemoHostUrl(url);
  return String(typeof resolver === 'function' ? resolver(context(sessionId)) : resolver);
}

describe('demo host URL routing', () => {
  it('gives every production demo editor a distinct first-level host origin', () => {
    const page = new URL('https://onlyoffice.getpi.work/');

    expect(new URL(resolve(page, 'office-editor-a')).origin).toBe(
      'https://office-editor-a.getpi.work',
    );
    expect(new URL(resolve(page, 'office-editor-b')).origin).toBe(
      'https://office-editor-b.getpi.work',
    );
  });

  it('migrates the legacy fixed R2 host to production wildcard origins', () => {
    const page = new URL(
      'https://onlyoffice.getpi.work/?hostUrl=https%3A%2F%2Fpub-7144f7712bc5465b880405ca7741e61f.r2.dev%2Foffice-host.html',
    );

    expect(new URL(resolve(page, 'office-editor-c')).origin).toBe(
      'https://office-editor-c.getpi.work',
    );
  });

  it('supports an explicit session template for non-production deployments', () => {
    const page = new URL(
      'https://demo.example.com/?hostUrl=https%3A%2F%2Foffice-%7BsessionId%7D.example.com%2Foffice-host.html',
    );

    expect(resolve(page, 'editor-1')).toBe(
      'https://office-editor-1.example.com/office-host.html',
    );
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
    expect(replaceState).toHaveBeenCalledWith(
      window.history.state,
      '',
      new URL('https://onlyoffice.getpi.work/'),
    );
  });
});
