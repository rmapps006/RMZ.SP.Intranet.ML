import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/security';
import { PermissionKind } from '@pnp/sp/security';

/** Title of the key/value list that stores the central intranet configuration. */
export const SETTINGS_LIST: string = 'Intranet Settings';
/** Title of the list that stores the top navigation links. */
export const NAV_LIST: string = 'Intranet Navigation';
/** Title of the single settings item that holds the JSON configuration blob. */
export const SETTINGS_KEY: string = 'config';
/** Internal field name on the settings list that stores the JSON blob. */
export const SETTINGS_VALUE_FIELD: string = 'SettingValue';

/** Internal field names on the navigation list. */
export const NAV_URL_FIELD: string = 'NavUrl';
export const NAV_ORDER_FIELD: string = 'SortOrder';
export const NAV_NEWTAB_FIELD: string = 'OpenInNewTab';

/** A single top-navigation link, edited from the Admin screen. */
export interface INavLink {
  label: string;
  url: string;
  newTab: boolean;
}

/**
 * A company email-domain filter for the Employee Directory. The label is the
 * chip text; the domain (e.g. "ora-uae.com") scopes which users are shown and
 * is matched against each user's mail / userPrincipalName.
 */
export interface IDirectoryDomain {
  label: string;
  domain: string;
}

/**
 * A department entry shown in the Departments directory grid on the main site.
 * Each links to that department's site.
 */
export interface IDepartmentEntry {
  label: string;
  description: string;
  url: string;
  icon?: string; // Fluent UI (Office) icon name; falls back to initials
  accent?: string; // tile icon background
}

/**
 * A single quick-access link tile, edited from the Admin screen and rendered by
 * the Quick Links web part and the Productivity Strip cards.
 */
export interface IQuickLinkSetting {
  label: string;
  img?: string; // image/logo URL or data URI; takes priority over icon and abbr
  icon?: string; // Fluent UI (Office) icon name; takes priority over abbr
  abbr?: string; // short text shown when no icon/img is set
  url: string;
  bg: string; // tile circle background
  color: string; // tile glyph/text colour
  newTab?: boolean;
}

/**
 * The complete set of centrally-managed intranet settings. Everything here is
 * editable from the Admin screen and consumed live by the header/footer
 * extension and the home-page web parts.
 */
export interface IIntranetSettings {
  // Branding — client name and logo, applied across the portal and department sites
  clientName: string; // brand/wordmark shown in footer, header alt and admin/content eyebrows
  logoUrl: string; // header logo image URL; blank = bundled default logo
  logoSubLabel: string; // small text beside the header logo; blank = hidden
  fontHead: string; // heading font family; blank = bundled ORA font (Raleway)
  fontBody: string; // body font family; blank = bundled ORA font (Public Sans)
  fontStylesheetUrl: string; // optional stylesheet URL (e.g. Google Fonts) to load a custom font
  // Footer
  footerVisible: boolean;
  footerCopyright: string;
  footerTagline: string;
  // Header / native chrome
  searchPlaceholder: string;
  hideSiteHeader: boolean;
  hideSiteNav: boolean;
  hideCommandBar: boolean;
  // Microsoft 365 / SharePoint suite bar (the tenant top bar) — blank = tenant default
  suiteBarColor: string;
  suiteBarTextColor: string;
  // Embedded portal
  embedHideUrl: boolean;
  // Web part defaults
  showViewAll: boolean;
  accentColor: string;
  accentColorDark: string; // secondary/gradient accent; blank = auto-derived from accentColor
  openLinksInNewTab: boolean; // when true, widget links open in a new browser tab
  detailPageUrl: string; // page hosting the Detail View web part (cards link here with ?type=&recId=)
  newsMaxItems: number;
  eventsMaxItems: number;
  directoryPageSize: number;
  // Employee Directory — which companies/users to show and how to filter
  directoryDomains: IDirectoryDomain[];
  directoryExclude: string[];
  // Centrally-managed link tiles
  quickLinks: IQuickLinkSetting[];
  // Department directory (grid of department sites)
  departments: IDepartmentEntry[];
}

/** Defaults used when the settings list / item is missing or a key is absent. */
export const DEFAULT_SETTINGS: IIntranetSettings = {
  clientName: 'ORA UAE',
  logoUrl: '',
  logoSubLabel: 'UAE',
  fontHead: '',
  fontBody: '',
  fontStylesheetUrl: '',
  footerVisible: true,
  footerCopyright: '', // blank → the footer derives "© {year} {clientName}. All rights reserved."
  footerTagline: '',
  searchPlaceholder: 'Search intranet…',
  hideSiteHeader: true,
  hideSiteNav: false,
  hideCommandBar: true,
  suiteBarColor: '',
  suiteBarTextColor: '',
  embedHideUrl: false,
  showViewAll: true,
  accentColor: '#b88a4a',
  accentColorDark: '',
  openLinksInNewTab: false,
  detailPageUrl: '',
  newsMaxItems: 6,
  eventsMaxItems: 4,
  directoryPageSize: 12,
  directoryDomains: [],
  directoryExclude: ['Service Account', 'Services Account', 'svc-', 'no-reply', 'noreply'],
  quickLinks: [],
  departments: []
};

/** Default navigation links seeded by provisioning and used as a fallback. */
export const DEFAULT_NAV: INavLink[] = [
  { label: 'Home', url: '/', newTab: false },
  { label: 'Departments', url: '/SitePages/Departments.aspx', newTab: false },
  { label: 'Policies', url: '/SitePages/Policies.aspx', newTab: false },
  { label: 'Documents', url: '/SitePages/Documents.aspx', newTab: false },
  { label: 'News', url: '/SitePages/News.aspx', newTab: false },
  { label: 'Events', url: '/SitePages/Events.aspx', newTab: false },
  { label: 'Onboarding', url: '/SitePages/Onboarding.aspx', newTab: false },
  { label: 'Approvals', url: '/SitePages/Approvals.aspx', newTab: false }
];

const CACHE_TTL_MS: number = 120000; // 2 minutes — long enough to avoid per-page reads.
const SETTINGS_CACHE_KEY: string = 'ora.settings.v1';
const NAV_CACHE_KEY: string = 'ora.nav.v1';

interface ICacheEnvelope<T> {
  ts: number;
  data: T;
}

function readCache<T>(key: string): T | undefined {
  try {
    const raw: string | null = window.sessionStorage.getItem(key);
    if (!raw) {
      return undefined;
    }
    const env: ICacheEnvelope<T> = JSON.parse(raw);
    if (Date.now() - env.ts > CACHE_TTL_MS) {
      return undefined;
    }
    return env.data;
  } catch {
    return undefined;
  }
}

function writeCache<T>(key: string, data: T): void {
  try {
    const env: ICacheEnvelope<T> = { ts: Date.now(), data };
    window.sessionStorage.setItem(key, JSON.stringify(env));
  } catch {
    /* sessionStorage may be unavailable — ignore. */
  }
}

/**
 * Synchronous best-effort read of the cached settings. The header/footer
 * extension warms this cache on every page, so web parts (and the shared
 * SectionHeader) can honour central defaults without each making a list call.
 * Falls back to the defaults when the cache is cold or unavailable.
 */
export function getCachedSettings(): IIntranetSettings {
  const cached: IIntranetSettings | undefined = readCache<IIntranetSettings>(SETTINGS_CACHE_KEY);
  return cached ? { ...DEFAULT_SETTINGS, ...cached } : { ...DEFAULT_SETTINGS };
}

function clearCache(): void {
  try {
    window.sessionStorage.removeItem(SETTINGS_CACHE_KEY);
    window.sessionStorage.removeItem(NAV_CACHE_KEY);
  } catch {
    /* ignore */
  }
}

/**
 * Reads and writes the central intranet settings. Context-agnostic: it accepts
 * a PnP `SPFI` instance, so both the application customizer (header/footer) and
 * the web parts can share it. Reads are cached in sessionStorage so a normal
 * page load does not pay for extra list calls.
 */
export class SettingsService {
  private readonly sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /** Returns the merged settings (defaults overlaid with stored values). */
  public async getSettings(useCache: boolean = true): Promise<IIntranetSettings> {
    if (useCache) {
      const cached: IIntranetSettings | undefined = readCache<IIntranetSettings>(SETTINGS_CACHE_KEY);
      if (cached) {
        return { ...DEFAULT_SETTINGS, ...cached };
      }
    }
    // Retry once on a transient failure (occasional 404/throttle right after the
    // page loads) before falling back, so widgets reading settings — e.g. the
    // Departments grid — don't intermittently render empty.
    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const items: Record<string, string>[] = await this.sp.web.lists
          .getByTitle(SETTINGS_LIST)
          .items.filter(`Title eq '${SETTINGS_KEY}'`)
          .top(1)
          .select('Title', SETTINGS_VALUE_FIELD)();
        const raw: string = items && items.length > 0 ? items[0][SETTINGS_VALUE_FIELD] || '' : '';
        const parsed: Partial<IIntranetSettings> = raw ? JSON.parse(raw) : {};
        const merged: IIntranetSettings = { ...DEFAULT_SETTINGS, ...parsed };
        writeCache(SETTINGS_CACHE_KEY, merged);
        return merged;
      } catch {
        if (attempt === 0) {
          await new Promise<void>((resolve) => window.setTimeout(resolve, 600));
          continue;
        }
        // List missing or access denied — fall back to defaults so the site still renders.
        return { ...DEFAULT_SETTINGS };
      }
    }
    return { ...DEFAULT_SETTINGS };
  }

  /** Returns the configured navigation links, ordered, or the defaults. */
  public async getNavLinks(useCache: boolean = true): Promise<INavLink[]> {
    if (useCache) {
      const cached: INavLink[] | undefined = readCache<INavLink[]>(NAV_CACHE_KEY);
      if (cached && cached.length > 0) {
        return cached;
      }
    }
    try {
      const rows: Record<string, unknown>[] = await this.sp.web.lists
        .getByTitle(NAV_LIST)
        .items.select('Title', NAV_URL_FIELD, NAV_ORDER_FIELD, NAV_NEWTAB_FIELD)
        .orderBy(NAV_ORDER_FIELD, true)
        .top(50)();

      const links: INavLink[] = (rows || [])
        .filter((r) => !!r.Title && !!r[NAV_URL_FIELD])
        .map((r) => ({
          label: String(r.Title),
          url: String(r[NAV_URL_FIELD]),
          newTab: r[NAV_NEWTAB_FIELD] === true
        }));

      if (links.length === 0) {
        return DEFAULT_NAV;
      }
      writeCache(NAV_CACHE_KEY, links);
      return links;
    } catch {
      return DEFAULT_NAV;
    }
  }

  /** Persists the settings JSON blob (Admin only). Clears the read cache. */
  public async saveSettings(settings: IIntranetSettings): Promise<void> {
    const list = this.sp.web.lists.getByTitle(SETTINGS_LIST);
    const value: string = JSON.stringify(settings);
    const existing: { Id: number }[] = await list.items
      .filter(`Title eq '${SETTINGS_KEY}'`)
      .top(1)
      .select('Id')();
    if (existing && existing.length > 0) {
      await list.items.getById(existing[0].Id).update({ [SETTINGS_VALUE_FIELD]: value });
    } else {
      await list.items.add({ Title: SETTINGS_KEY, [SETTINGS_VALUE_FIELD]: value });
    }
    clearCache();
  }

  /** Replaces all navigation rows with the supplied set (Admin only). */
  public async saveNavLinks(links: INavLink[]): Promise<void> {
    const list = this.sp.web.lists.getByTitle(NAV_LIST);
    const existing: { Id: number }[] = await list.items.select('Id').top(200)();
    for (const row of existing) {
      await list.items.getById(row.Id).delete();
    }
    let order: number = 1;
    for (const link of links) {
      if (!link.label || !link.url) {
        continue;
      }
      await list.items.add({
        Title: link.label,
        [NAV_URL_FIELD]: link.url,
        [NAV_ORDER_FIELD]: order,
        [NAV_NEWTAB_FIELD]: link.newTab === true
      });
      order += 1;
    }
    clearCache();
  }

  /** True when the current user can manage the web (Site Owner / admin). */
  public async canManage(): Promise<boolean> {
    try {
      const perms = await this.sp.web.getCurrentUserEffectivePermissions();
      return this.sp.web.hasPermissions(perms, PermissionKind.ManageWeb);
    } catch {
      return false;
    }
  }
}
