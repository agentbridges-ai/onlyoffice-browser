/**
 * Internationalization configuration
 */

/**
 * Language codes enum
 * Internal language codes (simplified): 'zh' | 'en'
 * OnlyOffice language codes (BCP 47 standard): 'zh-CN' | 'en'
 */
export enum LanguageCode {
  /** Simplified Chinese (internal) */
  ZH = 'zh',
  /** English (internal) */
  EN = 'en',
}

/**
 * OnlyOffice language codes (BCP 47 standard)
 */
export enum OnlyOfficeLanguageCode {
  /** Simplified Chinese (Mainland China) - BCP 47 standard */
  ZH_CN = 'zh-CN',
  /** English */
  EN = 'en',
}

export type Language = LanguageCode.ZH | LanguageCode.EN;

export interface I18nMessages {
  fileSavedSuccess: string;
  documentLoaded: string;
  failedToLoadEditor: string;
}

const messages: Record<Language, I18nMessages> = {
  [LanguageCode.ZH]: {
    fileSavedSuccess: '文件保存成功：',
    documentLoaded: '文档加载完成：',
    failedToLoadEditor: '无法加载编辑器组件。请确保已正确安装 OnlyOffice API。',
  },
  [LanguageCode.EN]: {
    fileSavedSuccess: 'File saved successfully: ',
    documentLoaded: 'Document loaded: ',
    failedToLoadEditor: 'Failed to load editor component. Please ensure OnlyOffice API is properly installed.',
  },
};

class I18n {
  private currentLanguage: Language = LanguageCode.EN;

  private getCookie(name: string): string | null {
    if (typeof document === 'undefined') return null;
    const cookies = document.cookie.split(';').map((value) => value.trim());
    const match = cookies.find((value) => value.startsWith(`${encodeURIComponent(name)}=`));
    return match ? decodeURIComponent(match.split('=').slice(1).join('=')) : null;
  }

  private getUrlParameter(name: string): string | null {
    if (typeof window === 'undefined') return null;
    return new URLSearchParams(window.location.search).get(name);
  }

  private getLocalStorageItem(name: string): string | null {
    try {
      return window.localStorage.getItem(name);
    } catch {
      return null;
    }
  }

  private setLocalStorageItem(name: string, value: string): void {
    try {
      window.localStorage.setItem(name, value);
    } catch {
      // Ignore storage errors in restricted browser contexts.
    }
  }

  /**
   * Normalize language code to LanguageCode enum
   * Supports: 'zh', 'zh-CN', 'zh_CN', 'en', 'en-US', etc.
   */
  private normalizeLanguage(lang: string | null): Language | null {
    if (!lang) return null;
    const normalized = lang.toLowerCase().split(/[-_]/)[0];
    if (normalized === 'zh') return LanguageCode.ZH;
    if (normalized === 'en') return LanguageCode.EN;
    return null;
  }

  constructor() {
    // Priority: path prefix -> URL locale -> cookie -> localStorage -> navigator.language -> 'en'
    let detectedLang: Language | null = null;

    // 0. Highest priority: sub-directory path prefix (e.g. /zh-cn/...)
    if (typeof window !== 'undefined' && window.location.pathname.startsWith('/zh-cn/')) {
      detectedLang = LanguageCode.ZH;
    }

    // 1. Try to get from URL parameter 'locale'
    if (!detectedLang) {
      const urlLocale = this.getUrlParameter('locale');
      detectedLang = this.normalizeLanguage(urlLocale);
    }

    // 2. If not found in URL, try cookies (locale field)
    if (!detectedLang) {
      const cookieLang = this.getCookie('locale');
      detectedLang = this.normalizeLanguage(cookieLang);
    }

    // 3. If not found in cookies, try localStorage
    if (!detectedLang) {
      const savedLang = this.getLocalStorageItem('document-lang') as Language;
      if (savedLang && (savedLang === LanguageCode.ZH || savedLang === LanguageCode.EN)) {
        detectedLang = savedLang;
      }
    }

    // 4. If not found in localStorage, try navigator.language
    if (!detectedLang) {
      const browserLang =
        // eslint-disable-next-line n/no-unsupported-features/node-builtins
        typeof navigator !== 'undefined' && navigator.language
          ? // eslint-disable-next-line n/no-unsupported-features/node-builtins
            navigator.language
          : LanguageCode.EN;
      detectedLang = this.normalizeLanguage(browserLang);
    }

    // 5. Default to 'en' if nothing found
    this.currentLanguage = detectedLang || LanguageCode.EN;
  }

  /**
   * Get current language
   */
  getLanguage(): Language {
    return this.currentLanguage;
  }

  /**
   * Set language
   */
  setLanguage(lang: Language): void {
    if (lang === LanguageCode.ZH || lang === LanguageCode.EN) {
      this.currentLanguage = lang;
      this.setLocalStorageItem('document-lang', lang);
      // Trigger language change event
      // eslint-disable-next-line n/no-unsupported-features/node-builtins
      window.dispatchEvent(new CustomEvent('languagechange', { detail: { language: lang } }));
    }
  }

  /**
   * Get translated text
   */
  t(key: keyof I18nMessages): string {
    return messages[this.currentLanguage][key] || messages[LanguageCode.EN][key] || key;
  }

  /**
   * Get all messages
   */
  getMessages(): I18nMessages {
    return messages[this.currentLanguage];
  }

  /**
   * Get OnlyOffice language code (BCP 47 standard)
   * OnlyOffice uses BCP 47 standard language codes
   * - English: 'en'
   * - Simplified Chinese (Mainland China): 'zh-CN'
   */
  getOnlyOfficeLang(): string {
    // Mapping from internal language code to OnlyOffice BCP 47 standard code
    const langMap: Record<Language, OnlyOfficeLanguageCode> = {
      [LanguageCode.ZH]: OnlyOfficeLanguageCode.ZH_CN,
      [LanguageCode.EN]: OnlyOfficeLanguageCode.EN,
    };
    return langMap[this.currentLanguage] || OnlyOfficeLanguageCode.EN;
  }
}

// Export singleton
export const i18n = new I18n();

// Export convenience functions
export const t = (key: keyof I18nMessages): string => i18n.t(key);
export const getLanguage = (): Language => i18n.getLanguage();
export const setLanguage = (lang: Language): void => i18n.setLanguage(lang);
export const getOnlyOfficeLang = (): string => i18n.getOnlyOfficeLang();
