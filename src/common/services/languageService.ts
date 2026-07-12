/**
 * Client-side language switching for English/Arabic bilingual content.
 *
 * Each visitor's active language is resolved once per page load by the
 * header/footer extension and cached in browser localStorage, so every web
 * part's service can read it synchronously via `getCurrentLanguage()`. The
 * first-visit default is personalised to the user: it follows their own
 * Microsoft 365 / SharePoint UI language, then the site-wide default, then
 * English. An explicit toggle always wins from then on.
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

/**
 * Maps a culture name (e.g. an M365/SharePoint UI culture like "ar-SA" or a
 * browser language like "en-US") to one of our two supported languages, or
 * undefined when it is neither Arabic nor English.
 */
export function languageFromCulture(culture: string | undefined): Language | undefined {
  const c: string = (culture || '').toLowerCase();
  if (c.indexOf('ar') === 0) {
    return 'ar';
  }
  if (c.indexOf('en') === 0) {
    return 'en';
  }
  return undefined;
}

/** Resolves the active language: explicit stored choice, else the given default, else English. */
export function getCurrentLanguage(defaultLanguage?: Language): Language {
  return readStoredLanguage() || defaultLanguage || 'en';
}

/**
 * Seeds the browser's language preference on first visit and applies the
 * resulting `dir`/`lang` to the document. Called once per page load by the
 * header/footer extension so every web part's services resolve the same
 * language via `getCurrentLanguage()` without each needing the site settings.
 *
 * First-visit precedence (once the visitor toggles, that stored choice wins
 * over everything below on subsequent visits):
 *   1. the user's own M365 / SharePoint UI language (`userUiCulture`), so an
 *      Arabic-language user lands on Arabic automatically — per user, not per
 *      site, and roaming across their devices via their M365 profile;
 *   2. the site-wide default language configured in the Admin screen;
 *   3. English.
 */
export function initializeLanguage(defaultLanguage: Language, userUiCulture?: string): Language {
  const stored: Language | undefined = readStoredLanguage();
  const resolved: Language = stored || languageFromCulture(userUiCulture) || defaultLanguage || 'en';
  if (!stored) {
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
