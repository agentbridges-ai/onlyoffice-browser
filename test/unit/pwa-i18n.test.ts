import { describe, expect, it } from 'vitest';
import { OFFICE_LOCALE_STORAGE_KEY, officeCopy, resolveOfficeLocale } from '../../src/pwa/i18n';

describe('standalone PWA i18n', () => {
  it('prefers a persisted supported locale', () => {
    expect(resolveOfficeLocale('en-US', ['zh-CN'])).toBe('en-US');
    expect(resolveOfficeLocale('zh-CN', ['en-US'])).toBe('zh-CN');
    expect(OFFICE_LOCALE_STORAGE_KEY).toBe('onlyoffice-browser.locale');
  });

  it('falls back through browser languages and then English', () => {
    expect(resolveOfficeLocale(null, ['fr-FR', 'zh-Hans'])).toBe('zh-CN');
    expect(resolveOfficeLocale(null, ['fr-FR', 'en-GB'])).toBe('en-US');
    expect(resolveOfficeLocale(null, ['fr-FR'])).toBe('en-US');
  });

  it('keeps both catalogs complete and removes the old privacy subtitle', () => {
    expect(Object.keys(officeCopy['zh-CN']).sort()).toEqual(Object.keys(officeCopy['en-US']).sort());
    expect(JSON.stringify(officeCopy)).not.toContain('文件只在此设备中处理');
    expect(JSON.stringify(officeCopy)).not.toContain('Files stay on this device');
    expect(officeCopy['zh-CN'].openDocuments).toBe('已打开的文档');
    expect(officeCopy['en-US'].openDocuments).toBe('Open documents');
  });
});
