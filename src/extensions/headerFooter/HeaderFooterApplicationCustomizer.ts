import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { spfi, SPFI, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';

import * as strings from 'HeaderFooterApplicationCustomizerStrings';

import { GraphUserService, ICurrentUser } from './services/GraphUserService';
import { IntranetHeader, INavItem } from './components/IntranetHeader';
import { IntranetFooter } from './components/IntranetFooter';
import { ensureBrandFonts } from './components/fonts';
import { SettingsService, IIntranetSettings, DEFAULT_SETTINGS, INavLink } from '../../common/services/SettingsService';
import { DepartmentSettingsService } from '../../common/services/DepartmentSettingsService';
import { getInitials } from '../../common/util/format';
import { Language, initializeLanguage, setLanguage, pickLocalized } from '../../common/services/languageService';

const LOG_SOURCE: string = 'HeaderFooterApplicationCustomizer';

// Bundled logo asset (emitted to the solution CDN at build time).
const LOGO_URL: string = require('./assets/logo.png');

/**
 * Returns a darker shade of a hex colour (for the secondary/gradient accent when
 * the admin hasn't set an explicit dark accent). Accepts #rgb or #rrggbb; returns
 * the input unchanged if it can't be parsed.
 */
function darken(hex: string, amount: number): string {
  const m: RegExpExecArray | null = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec((hex || '').trim());
  if (!m) {
    return hex;
  }
  let h: string = m[1];
  if (h.length === 3) {
    h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  }
  const factor: number = Math.max(0, Math.min(1, 1 - amount));
  const ch = (start: number): string => {
    const v: number = Math.round(parseInt(h.substr(start, 2), 16) * factor);
    return Math.max(0, Math.min(255, v)).toString(16).padStart(2, '0');
  };
  return `#${ch(0)}${ch(2)}${ch(4)}`;
}

const DEFAULT_NAV: INavItem[] = [
  { label: 'Home', url: '/' },
  { label: 'Departments', url: '/SitePages/Departments.aspx' },
  { label: 'Policies', url: '/SitePages/Policies.aspx' },
  { label: 'Documents', url: '/SitePages/Documents.aspx' },
  { label: 'News', url: '/SitePages/News.aspx' },
  { label: 'Events', url: '/SitePages/Events.aspx' },
  { label: 'Onboarding', url: '/SitePages/Onboarding.aspx' },
  { label: 'Approvals', url: '/SitePages/Approvals.aspx' }
];

export interface IHeaderFooterApplicationCustomizerProperties {
  /** Navigation links. Configurable per environment via component properties. */
  navItems?: INavItem[];
  /** Placeholder text for the header search box. */
  searchPlaceholder?: string;
  /** Footer copyright line. */
  copyright?: string;
  /** Footer right-hand tagline. */
  tagline?: string;
  /** Hide the native SharePoint site header (logo/title + site nav) below the Intranet header. Default true. */
  hideSiteHeader?: boolean;
  /** Hide only the native site navigation (keep the SharePoint site logo/title). Default false. */
  hideSiteNav?: boolean;
  /** Hide the SharePoint command bar (New / Page details …). Hides edit controls — default false. */
  hideCommandBar?: boolean;
}

export default class HeaderFooterApplicationCustomizer extends BaseApplicationCustomizer<IHeaderFooterApplicationCustomizerProperties> {
  private _sp: SPFI;
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _user: ICurrentUser = { displayName: '', initials: '?' };
  private _logoAlt: string = strings.LogoAlt;
  private _settings: IIntranetSettings = DEFAULT_SETTINGS;
  private _navLinks: INavLink[] = [];
  private _language: Language = 'en';

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initializing Intranet header/footer');

    this._sp = spfi().using(SPFx(this.context));

    ensureBrandFonts();
    // First paint: apply chrome from the static component properties to avoid a
    // flash of native chrome, then re-apply once the live settings have loaded.
    this._applyChromeOverrides();

    // On a department site the navigation/header/footer/theme are shared from the
    // Main Site. Read the local department settings first (this also warms the
    // department cache for the department web parts) to discover the Main Site URL.
    const deptService: DepartmentSettingsService = new DepartmentSettingsService(this._sp);
    let mainSiteUrl: string = '';
    try {
      const deptSettings = await deptService.getSettings();
      mainSiteUrl = (deptSettings.mainSiteUrl || '').trim();
    } catch {
      /* not a department site — read shared settings from the current web */
    }

    const currentUrl: string = this.context.pageContext.web.absoluteUrl;
    const norm = (u: string): string => u.toLowerCase().replace(/\/+$/, '');
    const useMain: boolean = !!mainSiteUrl && norm(mainSiteUrl) !== norm(currentUrl);
    // Shared settings/nav come from the Main Site when configured, else the current web.
    const sharedSp: SPFI = useMain ? spfi(mainSiteUrl).using(SPFx(this.context)) : this._sp;

    // Resolve the current user (Graph), web title and central settings/nav in parallel.
    const graphService: GraphUserService = new GraphUserService(this.context.msGraphClientFactory);
    const fallbackName: string = this.context.pageContext.user.displayName || strings.LogoAlt;
    const settingsService: SettingsService = new SettingsService(sharedSp);

    // Each read is guarded independently so a first-load failure of one (e.g. a
    // Graph or cross-site token that isn't warm yet) can't reject the whole batch
    // and leave the page without its header/footer until a manual refresh.
    const fallbackUser: ICurrentUser = { displayName: fallbackName, initials: getInitials(fallbackName) };
    const [user, webTitle, settings, navLinks] = await Promise.all([
      graphService.getCurrentUser(fallbackName).catch(() => fallbackUser),
      this._getWebTitle().catch(() => ''),
      settingsService.getSettings().catch(() => DEFAULT_SETTINGS),
      settingsService.getNavLinks().catch(() => [] as INavLink[])
    ]);

    this._user = user;
    this._logoAlt = webTitle || strings.LogoAlt;
    this._settings = settings;
    // Seeds the browser's language preference from the site default on first
    // visit (a no-op once the visitor has toggled explicitly) and applies
    // dir="rtl"/lang="ar" to the document so the whole page — including every
    // web part on it — renders in the resolved language/direction.
    this._language = initializeLanguage(settings.defaultLanguage);
    // Load a custom font stylesheet (e.g. Google Fonts) if the client configured one.
    this._applyCustomFontStylesheet();
    // On a department site the shared nav points at the Main Site's pages, so
    // resolve any site-relative URLs against the Main Site (absolute URLs pass
    // through unchanged).
    this._navLinks = useMain ? navLinks.map((n) => ({ ...n, url: this._resolveAgainstMain(n.url, mainSiteUrl) })) : navLinks;

    // Re-apply chrome now that the live settings are known.
    this._applyChromeOverrides();

    // Render as soon as the Top/Bottom placeholders are available (this is the
    // reliable first-paint signal — it also fires on initial load), and again on
    // client-side navigation so the active nav link stays in sync.
    this.context.placeholderProvider.changedEvent.add(this, this._render);
    this.context.application.navigatedEvent.add(this, this._render);
    this._render();

    return Promise.resolve();
  }

  /**
   * When the client has configured a custom font stylesheet URL (e.g. Google
   * Fonts), inject it once so the font families set via --intranet-font-* actually
   * load. Blank leaves the bundled Intranet fonts (no external request). Idempotent.
   */
  private _applyCustomFontStylesheet(): void {
    const url: string = (this._settings.fontStylesheetUrl || '').trim();
    const id: string = 'intranet-custom-font';
    const existing: HTMLElement | null = document.getElementById(id);
    if (!url) {
      if (existing) {
        existing.parentNode?.removeChild(existing);
      }
      return;
    }
    if (existing && (existing as HTMLLinkElement).href === url) {
      return;
    }
    if (existing) {
      existing.parentNode?.removeChild(existing);
    }
    const link: HTMLLinkElement = document.createElement('link');
    link.id = id;
    link.rel = 'stylesheet';
    link.href = url;
    document.head.appendChild(link);
  }

  /**
   * Injects CSS to hide native SharePoint chrome below the Intranet header, so it
   * doesn't duplicate our header/nav. Toggled via web part properties.
   */
  private _applyChromeOverrides(): void {
    // Live settings win; the static component properties act as the first-paint
    // fallback before the settings list has been read.
    const hideSiteHeader: boolean =
      this._settings.hideSiteHeader !== undefined ? this._settings.hideSiteHeader : this.properties.hideSiteHeader !== false;
    const hideSiteNav: boolean =
      this._settings.hideSiteNav !== undefined ? this._settings.hideSiteNav : this.properties.hideSiteNav === true;
    const hideCommandBar: boolean =
      this._settings.hideCommandBar !== undefined ? this._settings.hideCommandBar : this.properties.hideCommandBar !== false;

    // Never hide the command bar while the page is being edited — owners need
    // the Edit/Publish controls. Modern pages carry Mode=Edit in the URL.
    const isEditMode: boolean = /[?&]Mode=Edit/i.test(window.location.href);

    const rules: string[] = [];
    // Theme accent + fonts — consumed by web part SCSS via var(--intranet-accent / --intranet-font-*, …).
    // Fallback matches DEFAULT_SETTINGS.accentColor so first paint doesn't flash a different colour.
    const accent: string = this._settings.accentColor || '#b88a4a';
    const accentDark: string = (this._settings.accentColorDark || '').trim() || darken(accent, 0.18);
    const rootVars: string[] = [`--intranet-accent:${accent};`, `--intranet-accent-dark:${accentDark};`];
    if ((this._settings.fontHead || '').trim()) {
      rootVars.push(`--intranet-font-head:${this._settings.fontHead.trim()};`);
    }
    if ((this._settings.fontBody || '').trim()) {
      rootVars.push(`--intranet-font-body:${this._settings.fontBody.trim()};`);
    }
    rules.push(`:root{${rootVars.join('')}}`);
    if (hideSiteHeader) {
      rules.push('#spSiteHeader{display:none !important;}');
    }
    if (hideSiteNav) {
      rules.push(
        '#spSiteHeader [data-automation-id="SiteHeaderNavigationRegion"],.ms-HorizontalNav,#horizontalNav{display:none !important;}'
      );
    }
    if (hideCommandBar && !isEditMode) {
      rules.push('[data-automation-id="commandBarWrapper"]{display:none !important;}');
    }

    // Recolour the Microsoft 365 / SharePoint suite (tenant) top bar. Blank = leave the tenant default.
    // The bar splits into separately-coloured regions and elements: the brand
    // app-name link (.o365sx-appName) and the right-hand icon buttons
    // (.o365sx-button) each carry their own brand background, on top of the bar
    // regions (#leftRegion / #centerRegion / #rightRegion / #topLevelRegion).
    // All of them — including :hover/:visited — must be overridden.
    const suiteBar: string = (this._settings.suiteBarColor || '').trim();
    if (suiteBar) {
      rules.push(
        [
          '.o365sx-navbar',
          '#SuiteNavWrapper',
          '#O365_NavHeader',
          '#topLevelRegion',
          '#leftRegion',
          '#centerRegion',
          '#rightRegion',
          '#O365_SuiteBranding_container',
          '#O365_MainLink_TenantLogo_container',
          '.o365cs-base .o365sx-appName',
          '.o365cs-base .o365sx-appName:visited',
          '.o365cs-base .o365sx-appName:hover',
          '.o365cs-base .o365sx-button',
          '.o365cs-base .o365sx-button:hover',
          '.o365cs-base .o365sx-button:focus'
        ].join(',') + '{background-color:' + suiteBar + ' !important;background-image:none !important;}'
      );
      const suiteText: string = (this._settings.suiteBarTextColor || '').trim();
      if (suiteText) {
        rules.push(
          [
            '.o365cs-base .o365sx-appName',
            '.o365cs-base .o365sx-appName *',
            '.o365cs-base .o365sx-button',
            '.o365cs-base .o365sx-button *',
            '#SuiteNavWrapper a',
            '#SuiteNavWrapper span',
            '#SuiteNavWrapper i'
          ].join(',') + '{color:' + suiteText + ' !important;fill:' + suiteText + ' !important;}'
        );
      }
    }

    const existing: HTMLElement | null = document.getElementById('intranet-chrome-overrides');
    if (existing) {
      existing.parentNode?.removeChild(existing);
    }
    if (rules.length === 0) {
      return;
    }

    const style: HTMLStyleElement = document.createElement('style');
    style.id = 'intranet-chrome-overrides';
    style.appendChild(document.createTextNode(rules.join('')));
    document.head.appendChild(style);
  }

  /**
   * Resolves a navigation URL against the Main Site. Absolute (http[s]) URLs are
   * returned as-is; site-relative URLs are prefixed with the Main Site's
   * server-relative path so they point at the Main Site's pages, not the
   * department site.
   */
  private _resolveAgainstMain(url: string, mainSiteUrl: string): string {
    const u: string = (url || '').trim();
    if (!u || /^https?:\/\//i.test(u) || u.indexOf('mailto:') === 0 || u.indexOf('tel:') === 0) {
      return u;
    }
    try {
      const main: URL = new URL(mainSiteUrl);
      const basePath: string = main.pathname.replace(/\/+$/, ''); // e.g. /sites/Intranet
      const rel: string = u.charAt(0) === '/' ? u : `/${u}`;
      // If the stored URL is already server-relative to the Main Site
      // (e.g. "/sites/Intranet/SitePages/Home.aspx"), don't prefix it again.
      const lowerRel: string = rel.toLowerCase();
      const lowerBase: string = basePath.toLowerCase();
      if (lowerBase && (lowerRel === lowerBase || lowerRel.indexOf(`${lowerBase}/`) === 0)) {
        return `${main.origin}${rel}`;
      }
      return `${main.origin}${basePath}${rel}`;
    } catch {
      return u;
    }
  }

  private async _getWebTitle(): Promise<string> {
    try {
      const web: { Title: string } = await this._sp.web.select('Title')();
      return web.Title;
    } catch (error) {
      Log.warn(LOG_SOURCE, `Unable to read web title: ${error instanceof Error ? error.message : error}`);
      return '';
    }
  }

  private _render(): void {
    this._renderHeader();
    this._renderFooter();
  }

  private _renderHeader(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: this._onDispose
      });
    }

    if (!this._topPlaceholder || !this._topPlaceholder.domElement) {
      return;
    }

    // Clear previous content so the active-link state reflects the current page.
    const host: HTMLElement = this._topPlaceholder.domElement;
    while (host.firstChild) {
      host.removeChild(host.firstChild);
    }

    // Prefer the live navigation list; fall back to component properties, then defaults.
    // Labels are resolved to the active language here (Arabic falls back to the
    // English label when no translation was entered).
    const liveNav: INavItem[] = this._navLinks.map((n) => ({
      label: pickLocalized(n.label, n.labelAR, this._language),
      url: n.url,
      newTab: n.newTab
    }));
    const navItems: INavItem[] =
      liveNav.length > 0
        ? liveNav
        : this.properties.navItems && this.properties.navItems.length > 0
        ? this.properties.navItems
        : DEFAULT_NAV;

    const searchPlaceholder: string = pickLocalized(
      this._settings.searchPlaceholder || this.properties.searchPlaceholder || strings.SearchPlaceholder,
      this._settings.searchPlaceholderAR,
      this._language
    );
    const logoAlt: string = pickLocalized(this._settings.clientName || this._logoAlt, this._settings.clientNameAR, this._language);

    const header: IntranetHeader = new IntranetHeader({
      navItems,
      searchPlaceholder,
      searchAriaLabel: strings.SearchAriaLabel,
      notificationsLabel: strings.NotificationsLabel,
      logoUrl: this._settings.logoUrl || LOGO_URL,
      logoAlt,
      subLabel: this._settings.logoSubLabel,
      homeUrl: this.context.pageContext.web.absoluteUrl,
      webAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
      currentPath: window.location.pathname,
      user: this._user,
      language: this._language,
      onLanguageToggle: () => setLanguage(this._language === 'ar' ? 'en' : 'ar')
    });
    header.render(this._topPlaceholder.domElement);
  }

  private _renderFooter(): void {
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, {
        onDispose: this._onDispose
      });
    }

    if (!this._bottomPlaceholder || !this._bottomPlaceholder.domElement) {
      return;
    }

    const host: HTMLElement = this._bottomPlaceholder.domElement;

    // Footer can be hidden centrally from the Admin screen.
    if (this._settings.footerVisible === false) {
      while (host.firstChild) {
        host.removeChild(host.firstChild);
      }
      return;
    }

    if (host.childElementCount > 0) {
      return; // Footer is static; render once.
    }

    const clientName: string = this._settings.clientName || strings.LogoAlt;
    const defaultCopyright: string = `© ${new Date().getFullYear()} ${clientName}. All rights reserved.`;
    const footer: IntranetFooter = new IntranetFooter({
      wordmark: pickLocalized(clientName, this._settings.clientNameAR, this._language),
      copyright: pickLocalized(this._settings.footerCopyright || defaultCopyright, this._settings.footerCopyrightAR, this._language),
      tagline: pickLocalized(this._settings.footerTagline || '', this._settings.footerTaglineAR, this._language)
    });
    footer.render(host);
  }

  private _onDispose = (): void => {
    Log.info(LOG_SOURCE, 'Disposed Intranet header/footer placeholders');
  };
}
