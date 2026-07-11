/**
 * Client-side language switching for English/Arabic bilingual content.
 * The choice is a browser-local preference (there is no per-user profile
 * store available here), seeded from the site's configured default language.
 */

export type Language = 'en' | 'ar';

const LANG_STORAGE_KEY: string = 'intranet.lang.v1';

function readStoredLanguage(): Language | undefined {
  try {
    const raw: string | null = window.localStorage.getItem(LANG_STORAGE_KEY);
    return raw === 'ar' || raw === 'en' ? raw : undefined;
  } catch {
    return undefined;
  }
}

/** Resolves the active language: explicit user choice, else the site default, else English. */
export function getCurrentLanguage(defaultLanguage?: Language): Language {
  return readStoredLanguage() || defaultLanguage || 'en';
}

/**
 * Seeds the browser's language preference from the site default on first visit
 * (no-op if the visitor already made an explicit choice), and applies the
 * resulting `dir`/`lang` to the document. Called once per page load by the
 * header/footer extension so every web part's services resolve the same
 * language via `getCurrentLanguage()` without each needing the site settings.
 */
export function initializeLanguage(defaultLanguage: Language): Language {
  const resolved: Language = getCurrentLanguage(defaultLanguage);
  if (!readStoredLanguage()) {
    try {
      window.localStorage.setItem(LANG_STORAGE_KEY, resolved);
    } catch {
      /* localStorage unavailable — direction is still applied for this page view */
    }
  }
  applyDocumentDirection(resolved);
  return resolved;
}

/** Sets `document.documentElement`'s `lang`/`dir` to match the given language. Safe to call outside a DOM (no-ops). */
export function applyDocumentDirection(language: Language): void {
  try {
    document.documentElement.setAttribute('lang', language === 'ar' ? 'ar' : 'en');
    document.documentElement.setAttribute('dir', language === 'ar' ? 'rtl' : 'ltr');
  } catch {
    /* no DOM available — ignore */
  }
}

/** Persists the chosen language and reloads the page so every extension/web part re-renders in it. */
export function setLanguage(language: Language): void {
  try {
    window.localStorage.setItem(LANG_STORAGE_KEY, language);
  } catch {
    /* localStorage unavailable — the toggle still works for this page view via applyDocumentDirection */
  }
  applyDocumentDirection(language);
  window.location.reload();
}

export function isRtl(language: Language): boolean {
  return language === 'ar';
}

/** Returns the Arabic value when the active language is Arabic and it has content, else the English value. */
export function pickLocalized(en: string, ar: string | undefined, language: Language): string {
  if (language === 'ar' && ar && ar.trim()) {
    return ar;
  }
  return en || '';
}
