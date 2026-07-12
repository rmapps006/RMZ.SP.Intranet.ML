import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/security';
import '@pnp/sp/user-custom-actions';
import '@pnp/sp/clientside-pages';
import { ClientsideWebpart } from '@pnp/sp/clientside-pages';
import { PermissionKind } from '@pnp/sp/security';
import { IUserCustomActionInfo } from '@pnp/sp/user-custom-actions';

/** Component ID of the IntranetHeaderFooter application customizer (shared with the main site). */
const HEADER_FOOTER_COMPONENT_ID: string = '1083fb8f-96fc-49dc-9284-2d77402b58a9';
/** Component ID of the Content Manager web part (contributor dashboard). */
const CONTENT_MANAGER_COMPONENT_ID: string = 'e7767646-0e20-4cd6-ab84-5e2c44f05673';

/** Minimal shape of a client-side web part definition from getClientsideWebParts. */
interface IPartDef {
  Id?: string;
  Manifest?: string;
}

/** Title of the list (local to a department site) that stores the department configuration. */
export const DEPT_SETTINGS_LIST: string = 'Department Settings';
/** Title of the single settings item that holds the JSON configuration blob. */
export const DEPT_SETTINGS_KEY: string = 'config';
/** Internal field name that stores the JSON blob. */
export const DEPT_SETTINGS_VALUE_FIELD: string = 'SettingValue';

/** A single quick-link button shown in the department Quick Links row. */
export interface IDepartmentQuickAction {
  label: string;
  labelAR?: string;
  url: string;
}

/**
 * Per-department configuration, edited from the Department Admin web part on the
 * department site and consumed by the department web parts. The navigation,
 * header/footer and theme come from the Main Site (see mainSiteUrl) — only the
 * department content lives here.
 */
export interface IDepartmentSettings {
  // Link to the main portal site — its navigation/header/footer/theme are reused.
  mainSiteUrl: string;
  // Hero
  departmentName: string;
  departmentNameAR?: string;
  eyebrow: string;
  eyebrowAR?: string;
  description: string;
  descriptionAR?: string;
  ownerName: string;
  ownerNameAR?: string;
  ownerRole: string;
  ownerRoleAR?: string;
  // Quick Links row (Forms / Documents / Events …)
  quickActions: IDepartmentQuickAction[];
  // Department News
  newsList: string;
  allNewsUrl: string;
  // Department Events
  eventsList: string;
  calendarUrl: string;
  eventsMaxItems: number;
  detailPageUrl: string;
  // Forms & Documents
  policiesLibrary: string;
  documentsLibrary: string;
  panel1Title: string;
  panel1TitleAR?: string;
  panel2Title: string;
  panel2TitleAR?: string;
  documentHubUrl: string;
  // When true, department widget links open in a new browser tab.
  openLinksInNewTab: boolean;
}

/** Defaults used before the department has been configured (or the list is missing). */
export const DEFAULT_DEPARTMENT_SETTINGS: IDepartmentSettings = {
  mainSiteUrl: '',
  departmentName: '',
  departmentNameAR: '',
  eyebrow: 'Department',
  eyebrowAR: '',
  description: '',
  descriptionAR: '',
  ownerName: '',
  ownerNameAR: '',
  ownerRole: '',
  ownerRoleAR: '',
  quickActions: [
    { label: 'Forms', labelAR: 'النماذج', url: '' },
    { label: 'Documents', labelAR: 'المستندات', url: '' },
    { label: 'Events', labelAR: 'الفعاليات', url: '' }
  ],
  newsList: 'News',
  allNewsUrl: '',
  eventsList: 'Events',
  calendarUrl: '',
  eventsMaxItems: 3,
  detailPageUrl: '',
  policiesLibrary: 'Forms',
  documentsLibrary: 'Documents',
  panel1Title: 'Forms & Templates',
  panel1TitleAR: '',
  panel2Title: 'Shared Documents',
  panel2TitleAR: '',
  documentHubUrl: '',
  openLinksInNewTab: false
};

const CACHE_TTL_MS: number = 120000; // 2 minutes
const DEPT_CACHE_KEY: string = 'intranet.deptsettings.v1';

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
    window.sessionStorage.setItem(key, JSON.stringify({ ts: Date.now(), data }));
  } catch {
    /* sessionStorage unavailable — ignore */
  }
}

/**
 * Synchronous best-effort read of the cached department settings. The
 * header/footer extension warms this cache on every page of a department site,
 * so the department web parts can render without each making a list call.
 */
export function getCachedDepartmentSettings(): IDepartmentSettings {
  const cached: IDepartmentSettings | undefined = readCache<IDepartmentSettings>(DEPT_CACHE_KEY);
  return cached ? { ...DEFAULT_DEPARTMENT_SETTINGS, ...cached } : { ...DEFAULT_DEPARTMENT_SETTINGS };
}

/** Reads and writes the per-department settings on the department site. */
export class DepartmentSettingsService {
  private readonly sp: SPFI;

  constructor(sp: SPFI) {
    this.sp = sp;
  }

  /** Returns the merged department settings (defaults overlaid with stored values). */
  public async getSettings(useCache: boolean = true): Promise<IDepartmentSettings> {
    if (useCache) {
      const cached: IDepartmentSettings | undefined = readCache<IDepartmentSettings>(DEPT_CACHE_KEY);
      if (cached) {
        return { ...DEFAULT_DEPARTMENT_SETTINGS, ...cached };
      }
    }
    try {
      const items: Record<string, string>[] = await this.sp.web.lists
        .getByTitle(DEPT_SETTINGS_LIST)
        .items.filter(`Title eq '${DEPT_SETTINGS_KEY}'`)
        .top(1)
        .select('Title', DEPT_SETTINGS_VALUE_FIELD)();
      const raw: string = items && items.length > 0 ? items[0][DEPT_SETTINGS_VALUE_FIELD] || '' : '';
      const parsed: Partial<IDepartmentSettings> = raw ? JSON.parse(raw) : {};
      const merged: IDepartmentSettings = { ...DEFAULT_DEPARTMENT_SETTINGS, ...parsed };
      writeCache(DEPT_CACHE_KEY, merged);
      return merged;
    } catch {
      // List missing (not a department site, or not yet set up) — return defaults.
      return { ...DEFAULT_DEPARTMENT_SETTINGS };
    }
  }

  /** Ensures the (hidden) department settings list exists. Idempotent. */
  public async ensureList(): Promise<void> {
    const list = this.sp.web.lists.getByTitle(DEPT_SETTINGS_LIST);
    let exists: boolean = true;
    try {
      await list.select('Id')();
    } catch {
      exists = false;
    }
    if (!exists) {
      await this.sp.web.lists.add(DEPT_SETTINGS_LIST, 'Department configuration (managed by Department Admin).', 100, false, {
        Hidden: true,
        OnQuickLaunch: false
      });
      await this.sp.web.lists.getByTitle(DEPT_SETTINGS_LIST).fields.addMultilineText(DEPT_SETTINGS_VALUE_FIELD, {
        NumberOfLines: 12,
        RichText: false
      });
    }
  }

  /**
   * Registers the Intranet header/footer application customizer on this (department)
   * site so it renders the shared navigation/header/footer. Idempotent — skips
   * if already registered. Returns true if it newly registered.
   */
  public async registerHeaderFooter(): Promise<boolean> {
    const actions: { ClientSideComponentId?: string }[] = await this.sp.web.userCustomActions();
    const exists: boolean = actions.some(
      (a) => (a.ClientSideComponentId || '').toLowerCase() === HEADER_FOOTER_COMPONENT_ID
    );
    if (exists) {
      return false;
    }
    // ClientSideComponentId/Properties are accepted by the REST API but are not
    // part of the PnP IUserCustomActionInfo type, so cast through unknown.
    const action: Record<string, string> = {
      Title: 'IntranetHeaderFooter',
      Name: 'IntranetHeaderFooter',
      Location: 'ClientSideExtension.ApplicationCustomizer',
      ClientSideComponentId: HEADER_FOOTER_COMPONENT_ID,
      ClientSideComponentProperties: '{}'
    };
    await this.sp.web.userCustomActions.add(action as unknown as Partial<IUserCustomActionInfo>);
    return true;
  }

  /** Ensures a list/library exists (idempotent). template 100 = list, 101 = document library. */
  private async ensureContentList(title: string, template: number): Promise<void> {
    if (!title) {
      return;
    }
    try {
      await this.sp.web.lists.getByTitle(title).select('Id')();
    } catch {
      await this.sp.web.lists.add(title, `${title} (provisioned by Department Admin).`, template);
    }
  }

  /** True when a field with the given internal name/title already exists. */
  private async fieldExists(listTitle: string, fieldName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listTitle).fields.getByInternalNameOrTitle(fieldName).select('Id')();
      return true;
    } catch {
      return false;
    }
  }

  /**
   * Adds a field to a list only when it isn't already present (idempotent).
   * Checks existence by name first — some field types (e.g. the modern Image
   * column) don't error on a duplicate add and would otherwise create a new
   * suffixed column on every re-run.
   */
  private async ensureField(listTitle: string, fieldName: string, add: () => Promise<unknown>): Promise<void> {
    if (await this.fieldExists(listTitle, fieldName)) {
      return;
    }
    try {
      await add();
    } catch {
      /* lost a race or the field appeared — ignore */
    }
  }

  /**
   * Provisions the department content lists/libraries this site's web parts use:
   * a News list (with image + link), an Events list, and the Forms and Documents
   * libraries. Idempotent. Uses the names from the supplied settings.
   */
  public async provisionContentLists(s: IDepartmentSettings): Promise<void> {
    const newsList: string = s.newsList || 'News';
    const eventsList: string = s.eventsList || 'Events';
    const formsLib: string = s.policiesLibrary || 'Forms';
    const docsLib: string = s.documentsLibrary || 'Documents';

    await this.ensureContentList(newsList, 100);
    const news = this.sp.web.lists.getByTitle(newsList);
    await this.ensureField(newsList, 'Category', () => news.fields.addChoice('Category', { Choices: ['General', 'Announcement', 'Update', 'Milestone'] }));
    await this.ensureField(newsList, 'NewsDate', () => news.fields.addDateTime('NewsDate'));
    await this.ensureField(newsList, 'Source', () => news.fields.addText('Source'));
    await this.ensureField(newsList, 'LinkUrl', () => news.fields.addText('LinkUrl', { MaxLength: 500 }));
    // Paste a hosted image URL here to show a picture on the news card.
    await this.ensureField(newsList, 'ImageUrl', () => news.fields.addText('ImageUrl', { MaxLength: 500 }));
    await this.ensureField(newsList, 'Body', () => news.fields.addMultilineText('Body', { RichText: true }));
    // Arabic translations — same content, rendered when the visitor's language is Arabic.
    await this.ensureField(newsList, 'TitleAR', () => news.fields.addText('TitleAR'));
    await this.ensureField(newsList, 'SourceAR', () => news.fields.addText('SourceAR'));
    await this.ensureField(newsList, 'BodyAR', () => news.fields.addMultilineText('BodyAR', { RichText: true }));

    await this.ensureContentList(eventsList, 100);
    const events = this.sp.web.lists.getByTitle(eventsList);
    await this.ensureField(eventsList, 'EventDate', () => events.fields.addDateTime('EventDate'));
    await this.ensureField(eventsList, 'EndDate', () => events.fields.addDateTime('EndDate'));
    await this.ensureField(eventsList, 'Location', () => events.fields.addText('Location'));
    await this.ensureField(eventsList, 'Category', () => events.fields.addChoice('Category', { Choices: ['Town Hall', 'Learning', 'Wellness', 'Training', 'General'] }));
    await this.ensureField(eventsList, 'Description', () => events.fields.addMultilineText('Description', { RichText: true }));
    // Arabic translations — same content, rendered when the visitor's language is Arabic.
    await this.ensureField(eventsList, 'TitleAR', () => events.fields.addText('TitleAR'));
    await this.ensureField(eventsList, 'LocationAR', () => events.fields.addText('LocationAR'));
    await this.ensureField(eventsList, 'DescriptionAR', () => events.fields.addMultilineText('DescriptionAR', { RichText: true }));

    await this.ensureContentList(formsLib, 101);
    await this.ensureContentList(docsLib, 101);
  }

  /** Adds the given rows to a list only when it is currently empty. Idempotent. */
  private async seedIfEmpty(listTitle: string, rows: Record<string, unknown>[]): Promise<void> {
    try {
      const existing: { Id: number }[] = await this.sp.web.lists.getByTitle(listTitle).items.top(1)();
      if (existing && existing.length > 0) {
        return;
      }
      const list = this.sp.web.lists.getByTitle(listTitle);
      for (const row of rows) {
        await list.items.add(row);
      }
    } catch {
      /* non-fatal — seeding is best-effort. */
    }
  }

  /**
   * Seeds a couple of sample News and Events rows into the department content
   * lists, but only when each list is empty (so re-running never duplicates).
   */
  public async seedSampleContent(s: IDepartmentSettings): Promise<void> {
    const newsList: string = s.newsList || 'News';
    const eventsList: string = s.eventsList || 'Events';
    const deptName: string = s.departmentName || 'Department';

    await this.seedIfEmpty(newsList, [
      {
        Title: `${deptName} Kicks Off 2026 Priorities`,
        TitleAR: `${deptName} يطلق أولويات عام ٢٠٢٦`,
        Category: 'Announcement',
        NewsDate: '2026-07-01T08:00:00Z',
        Source: deptName,
        SourceAR: deptName,
        Body: `<p>The ${deptName} team has shared its priorities and key initiatives for the year ahead.</p>`,
        BodyAR: `<p>شارك فريق ${deptName} أولوياته ومبادراته الرئيسية للعام المقبل.</p>`
      },
      {
        Title: 'New Team Members Join Us',
        TitleAR: 'أعضاء جدد ينضمون إلى فريقنا',
        Category: 'Update',
        NewsDate: '2026-06-15T08:00:00Z',
        Source: deptName,
        SourceAR: deptName
      }
    ]);

    await this.seedIfEmpty(eventsList, [
      {
        Title: `${deptName} Monthly Town Hall`,
        TitleAR: `اللقاء الشهري المفتوح لـ ${deptName}`,
        EventDate: '2026-07-15T09:00:00Z',
        EndDate: '2026-07-15T10:00:00Z',
        Location: 'Microsoft Teams',
        LocationAR: 'Microsoft Teams',
        Category: 'Town Hall'
      },
      {
        Title: 'Quarterly Learning Session',
        TitleAR: 'جلسة التعلم الفصلية',
        EventDate: '2026-07-22T13:00:00Z',
        EndDate: '2026-07-22T14:30:00Z',
        Location: 'Training Room 2',
        LocationAR: 'قاعة التدريب ٢',
        Category: 'Learning'
      }
    ]);
  }

  /** Finds a client-side web part definition by component id (brace/case tolerant). */
  private findPartDef(defs: IPartDef[], targetId: string): IPartDef | undefined {
    const norm = (v: string | undefined): string => (v || '').replace(/[{}]/g, '').toLowerCase();
    const target: string = norm(targetId);
    return defs.filter((c) => {
      if (norm(c.Id) === target) {
        return true;
      }
      try {
        const manifest: { id?: string } = JSON.parse(c.Manifest || '{}');
        return norm(manifest.id) === target;
      } catch {
        return false;
      }
    })[0];
  }

  /**
   * Creates a Content.aspx page on the department site with the Content Manager
   * web part (News + Events for this site). Idempotent: leaves an existing page
   * untouched. Best-effort — never throws.
   */
  private async ensureContentPage(s: IDepartmentSettings): Promise<void> {
    try {
      const existing: { Id: number }[] = await this.sp.web.lists
        .getByTitle('Site Pages')
        .items.filter(`FileLeafRef eq 'Content.aspx'`)
        .top(1)
        .select('Id')();
      if (existing && existing.length > 0) {
        return;
      }
      const defs: IPartDef[] = await this.sp.web.getClientsideWebParts();
      const def: IPartDef | undefined = this.findPartDef(defs, CONTENT_MANAGER_COMPONENT_ID);
      if (!def) {
        return; // package not deployed/updated on this site yet
      }
      const page = await this.sp.web.addClientsidePage('Content.aspx', 'Content Manager', 'Article');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const part = ClientsideWebpart.fromComponentDef(def as any);
      part.setProperties({
        title: 'Content Manager',
        newsList: s.newsList || 'News',
        eventsList: s.eventsList || 'Events',
        showNews: true,
        showEvents: true,
        showBenefits: false
      });
      page.addSection().addControl(part);
      await page.save();
    } catch {
      /* non-fatal — the web part can still be added to a page manually. */
    }
  }

  /** Runs department setup: settings list, header/footer registration and content lists. */
  public async runSetup(settings: IDepartmentSettings, seedSampleData: boolean = false): Promise<void> {
    await this.ensureList();
    await this.registerHeaderFooter();
    await this.provisionContentLists(settings);
    if (seedSampleData) {
      await this.seedSampleContent(settings);
    }
    await this.ensureContentPage(settings);
  }

  /** Persists the settings JSON blob (Admin only). Clears the read cache. */
  public async saveSettings(settings: IDepartmentSettings): Promise<void> {
    await this.ensureList();
    const list = this.sp.web.lists.getByTitle(DEPT_SETTINGS_LIST);
    const value: string = JSON.stringify(settings);
    const existing: { Id: number }[] = await list.items
      .filter(`Title eq '${DEPT_SETTINGS_KEY}'`)
      .top(1)
      .select('Id')();
    if (existing && existing.length > 0) {
      await list.items.getById(existing[0].Id).update({ [DEPT_SETTINGS_VALUE_FIELD]: value });
    } else {
      await list.items.add({ Title: DEPT_SETTINGS_KEY, [DEPT_SETTINGS_VALUE_FIELD]: value });
    }
    try {
      window.sessionStorage.removeItem(DEPT_CACHE_KEY);
    } catch {
      /* ignore */
    }
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
