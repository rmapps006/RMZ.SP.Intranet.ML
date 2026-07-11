import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/user-custom-actions';
import '@pnp/sp/folders';
import { IUserCustomActionInfo } from '@pnp/sp/user-custom-actions';
import '@pnp/sp/clientside-pages';
import { ClientsideWebpart } from '@pnp/sp/clientside-pages';
import {
  SETTINGS_LIST,
  NAV_LIST,
  SETTINGS_KEY,
  SETTINGS_VALUE_FIELD,
  NAV_URL_FIELD,
  NAV_ORDER_FIELD,
  NAV_NEWTAB_FIELD,
  DEFAULT_SETTINGS,
  DEFAULT_NAV
} from '../../../common/services/SettingsService';

export type ProvisionStatus = 'ok' | 'skipped' | 'error';

export interface IProvisionResult {
  name: string;
  status: ProvisionStatus;
  message: string;
}

/** Shape of a client-side web part definition returned by getClientsideWebParts. */
interface IPartDef {
  Id?: string;
  Manifest?: string;
  Name?: string;
}

const HEADER_FOOTER_COMPONENT_ID: string = 'c4e5110d-1d14-4170-8db9-987c63c015ac';
const DETAIL_VIEW_COMPONENT_ID: string = 'fa5a4c01-9751-44ce-88d2-9ca0a1443cab';
const CONTENT_MANAGER_COMPONENT_ID: string = 'c6324e68-dcc1-469d-86f4-bea2a661e608';

const HEADER_FOOTER_PROPERTIES: string = JSON.stringify({
  navItems: [
    { label: 'Home', url: '/' },
    { label: 'Departments', url: '/SitePages/Departments.aspx' },
    { label: 'Policies', url: '/SitePages/Policies.aspx' },
    { label: 'Documents', url: '/SitePages/Documents.aspx' },
    { label: 'News', url: '/SitePages/News.aspx' },
    { label: 'Events', url: '/SitePages/Events.aspx' },
    { label: 'Onboarding', url: '/SitePages/Onboarding.aspx' },
    { label: 'Approvals', url: '/SitePages/Approvals.aspx' }
  ],
  searchPlaceholder: 'Search intranet…',
  copyright: '',
  tagline: '',
  hideSiteHeader: true,
  hideCommandBar: true
});

export interface IProvisioningOptions {
  eventsList: string;
  policiesList: string;
  newsList: string;
  benefitsList: string;
  formsLibrary: string;
  templatesLibrary: string;
  pagesLibrary: string;
  registerHeaderFooter: boolean;
  seedSampleData: boolean;
  createViewAllPages: boolean;
}

export class ProvisioningService {
  private readonly sp: SPFI;

  constructor(context: WebPartContext) {
    this.sp = getSP(context);
  }

  /** Runs every provisioning task in sequence and returns a result per task. */
  public async runAll(options: IProvisioningOptions): Promise<IProvisionResult[]> {
    const results: IProvisionResult[] = [];

    results.push(
      await this._ensureList(options.eventsList, 100, 'Company events for the Events web part', async (listTitle) => {
        await this._ensureDateField(listTitle, 'EventDate');
        await this._ensureDateField(listTitle, 'EndDate');
        await this._ensureTextField(listTitle, 'Location');
        await this._ensureChoiceField(listTitle, 'Category', [
          'All-Company',
          'Learning & Development',
          'Finance',
          'HR'
        ]);
        await this._ensureMultilineField(listTitle, 'Description');
      })
    );

    results.push(
      // A document library (101): each policy IS the uploaded file, with a bit of
      // metadata. The detail page previews the file directly.
      await this._ensureList(options.policiesList, 101, 'Policy documents for the Policies web part', async (listTitle) => {
        await this._ensureChoiceField(listTitle, 'Department', [
          'HR',
          'Finance',
          'IT',
          'Operations',
          'Marketing',
          'Legal'
        ]);
        await this._ensureTextField(listTitle, 'PolicyVersion');
      })
    );

    results.push(
      await this._ensureList(options.newsList, 100, 'Company news for the News web part', async (listTitle) => {
        await this._ensureChoiceField(listTitle, 'Category', [
          'General Announcements',
          'Construction Updates',
          'HR Updates',
          'Finance'
        ]);
        await this._ensureDateField(listTitle, 'NewsDate');
        await this._ensureTextField(listTitle, 'Source');
        await this._ensureTextField(listTitle, 'LinkUrl');
        // Paste a hosted image URL here to show a picture on the news card.
        await this._ensureTextField(listTitle, 'ImageUrl');
        await this._ensureMultilineField(listTitle, 'Body');
      })
    );

    results.push(
      await this._ensureList(options.benefitsList, 100, 'Employee benefits catalogue', async (listTitle) => {
        await this._ensureChoiceField(listTitle, 'Category', [
          'Health',
          'Wellness',
          'Leave',
          'Financial',
          'Learning',
          'Family'
        ]);
        await this._ensureTextField(listTitle, 'Summary');
        await this._ensureTextField(listTitle, 'Eligibility');
        await this._ensureTextField(listTitle, 'Coverage');
        await this._ensureMultilineField(listTitle, 'Details');
      })
    );

    results.push(await this._ensureList(options.formsLibrary, 101, 'Forms document library'));
    results.push(await this._ensureList(options.templatesLibrary, 101, 'Templates document library'));

    results.push(await this._ensureSettingsList());
    results.push(await this._ensureNavigationList());
    results.push(await this._ensureLogosFolder());

    if (options.registerHeaderFooter) {
      results.push(await this._ensureHeaderFooter());
    }

    if (options.seedSampleData) {
      results.push(await this._seedEvents(options.eventsList));
      // Policies is a document library — seeding needs real files, so it's
      // skipped; upload policy documents to the library instead.
      results.push(await this._seedNews(options.newsList));
      results.push(await this._seedBenefits(options.benefitsList));
    }

    if (options.createViewAllPages) {
      const pageResults: IProvisionResult[] = await this._ensureViewAllPages(options);
      pageResults.forEach((r) => results.push(r));
      results.push(await this._ensureDetailPageSetting());
    }

    return results;
  }

  /**
   * Points the central `detailPageUrl` setting at the provisioned Detail page so
   * News/Events/Employee cards can link to it. Only fills the value when an admin
   * has not already set one, so a manual override is never clobbered. Idempotent.
   */
  private async _ensureDetailPageSetting(): Promise<IProvisionResult> {
    const name: string = 'Detail page URL';
    try {
      const web: { ServerRelativeUrl?: string } = await this.sp.web.select('ServerRelativeUrl')();
      const base: string = (web.ServerRelativeUrl || '').replace(/\/$/, '');
      const detailUrl: string = `${base}/SitePages/Detail.aspx`;

      const items: { Id: number; [key: string]: unknown }[] = await this.sp.web.lists
        .getByTitle(SETTINGS_LIST)
        .items.filter(`Title eq '${SETTINGS_KEY}'`)
        .top(1)
        .select('Id', SETTINGS_VALUE_FIELD)();
      if (!items || items.length === 0) {
        return { name, status: 'error', message: 'Settings config item not found — run the settings step first.' };
      }

      const raw: string = (items[0][SETTINGS_VALUE_FIELD] as string) || '{}';
      let parsed: Record<string, unknown>;
      try {
        parsed = JSON.parse(raw);
      } catch {
        parsed = {};
      }
      if (parsed.detailPageUrl && String(parsed.detailPageUrl).trim()) {
        return { name, status: 'skipped', message: `Already set to ${String(parsed.detailPageUrl)}.` };
      }

      parsed.detailPageUrl = detailUrl;
      await this.sp.web.lists
        .getByTitle(SETTINGS_LIST)
        .items.getById(items[0].Id)
        .update({ [SETTINGS_VALUE_FIELD]: JSON.stringify(parsed) });
      return { name, status: 'ok', message: `Set to ${detailUrl}.` };
    } catch (error) {
      return { name, status: 'error', message: this._msg(error) };
    }
  }

  private async _listExists(title: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(title).select('Id')();
      return true;
    } catch {
      return false;
    }
  }

  private async _ensureList(
    title: string,
    template: number,
    description: string,
    configureFields?: (listTitle: string) => Promise<void>
  ): Promise<IProvisionResult> {
    try {
      const existed: boolean = await this._listExists(title);
      if (!existed) {
        await this.sp.web.lists.add(title, description, template);
      }
      if (configureFields) {
        await configureFields(title);
      }
      return {
        name: title,
        status: existed ? 'skipped' : 'ok',
        message: existed ? 'Already existed — fields verified.' : 'Created with required columns.'
      };
    } catch (error) {
      return { name: title, status: 'error', message: this._msg(error) };
    }
  }

  private async _fieldExists(listTitle: string, fieldName: string): Promise<boolean> {
    try {
      await this.sp.web.lists.getByTitle(listTitle).fields.getByInternalNameOrTitle(fieldName).select('Id')();
      return true;
    } catch {
      return false;
    }
  }

  private async _ensureTextField(listTitle: string, name: string): Promise<void> {
    if (!(await this._fieldExists(listTitle, name))) {
      await this.sp.web.lists.getByTitle(listTitle).fields.addText(name);
    }
  }

  private async _ensureDateField(listTitle: string, name: string): Promise<void> {
    if (!(await this._fieldExists(listTitle, name))) {
      await this.sp.web.lists.getByTitle(listTitle).fields.addDateTime(name);
    }
  }

  private async _ensureMultilineField(listTitle: string, name: string): Promise<void> {
    if (!(await this._fieldExists(listTitle, name))) {
      await this.sp.web.lists.getByTitle(listTitle).fields.addMultilineText(name, { RichText: true });
    }
  }

  private async _ensureChoiceField(listTitle: string, name: string, choices: string[]): Promise<void> {
    if (!(await this._fieldExists(listTitle, name))) {
      await this.sp.web.lists.getByTitle(listTitle).fields.addChoice(name, { Choices: choices });
      return;
    }
    // Field already exists — keep its choices in sync with the code (e.g. so a
    // removed category disappears). Existing items retain any stored value.
    try {
      await this.sp.web.lists
        .getByTitle(listTitle)
        .fields.getByInternalNameOrTitle(name)
        .update({ Choices: choices }, 'SP.FieldChoice');
    } catch {
      /* non-fatal — the column is still usable with its current choices. */
    }
  }

  /**
   * Creates a "logos" folder under the Site Assets library so admins have a
   * standard place to upload app logos (used by the Quick Links tile `img`
   * field). Idempotent.
   */
  private async _ensureLogosFolder(): Promise<IProvisionResult> {
    try {
      const web: { ServerRelativeUrl?: string } = await this.sp.web.select('ServerRelativeUrl')();
      const base: string = (web.ServerRelativeUrl || '').replace(/\/+$/, '');
      const folderPath: string = `${base}/SiteAssets/logos`;

      let exists: boolean = false;
      try {
        const info: { Exists?: boolean } = await this.sp.web.getFolderByServerRelativePath(folderPath).select('Exists')();
        exists = info && info.Exists === true;
      } catch {
        exists = false;
      }
      if (exists) {
        return { name: 'Logos folder', status: 'skipped', message: 'SiteAssets/logos already exists.' };
      }

      await this.sp.web.folders.addUsingPath('SiteAssets/logos');
      return { name: 'Logos folder', status: 'ok', message: 'Created SiteAssets/logos — upload app logos here and use their URL in Quick Links tiles.' };
    } catch (error) {
      return { name: 'Logos folder', status: 'error', message: this._msg(error) };
    }
  }

  private async _ensureHeaderFooter(): Promise<IProvisionResult> {
    try {
      const actions: { ClientSideComponentId?: string }[] = await this.sp.web.userCustomActions();
      const exists: boolean = actions.some(
        (a) => (a.ClientSideComponentId || '').toLowerCase() === HEADER_FOOTER_COMPONENT_ID
      );
      if (exists) {
        return { name: 'Header / Footer registration', status: 'skipped', message: 'Already registered on this site.' };
      }
      // ClientSideComponentId/Properties are accepted by the REST API but are not
      // part of the PnP IUserCustomActionInfo type, so cast through unknown.
      const action: Record<string, string> = {
        Title: 'ORAHeaderFooter',
        Name: 'ORAHeaderFooter',
        Location: 'ClientSideExtension.ApplicationCustomizer',
        ClientSideComponentId: HEADER_FOOTER_COMPONENT_ID,
        ClientSideComponentProperties: HEADER_FOOTER_PROPERTIES
      };
      await this.sp.web.userCustomActions.add(action as unknown as Partial<IUserCustomActionInfo>);
      return { name: 'Header / Footer registration', status: 'ok', message: 'Registered — reload to see the header & footer.' };
    } catch (error) {
      return { name: 'Header / Footer registration', status: 'error', message: this._msg(error) };
    }
  }

  /**
   * Ensures the central settings list exists and holds a single `config` item
   * seeded with the defaults. Settings are stored as a JSON blob in a note field.
   */
  private async _ensureSettingsList(): Promise<IProvisionResult> {
    try {
      const existed: boolean = await this._listExists(SETTINGS_LIST);
      if (!existed) {
        await this.sp.web.lists.add(SETTINGS_LIST, 'Central configuration for the intranet', 100);
        await this._hideList(SETTINGS_LIST);
      }
      if (!(await this._fieldExists(SETTINGS_LIST, SETTINGS_VALUE_FIELD))) {
        await this.sp.web.lists.getByTitle(SETTINGS_LIST).fields.addMultilineText(SETTINGS_VALUE_FIELD, {
          NumberOfLines: 12,
          RichText: false
        });
      }
      const items: { Id: number }[] = await this.sp.web.lists
        .getByTitle(SETTINGS_LIST)
        .items.filter(`Title eq '${SETTINGS_KEY}'`)
        .top(1)
        .select('Id')();
      if (!items || items.length === 0) {
        await this.sp.web.lists.getByTitle(SETTINGS_LIST).items.add({
          Title: SETTINGS_KEY,
          [SETTINGS_VALUE_FIELD]: JSON.stringify(DEFAULT_SETTINGS)
        });
      }
      return {
        name: SETTINGS_LIST,
        status: existed ? 'skipped' : 'ok',
        message: existed ? 'Already existed — config item verified.' : 'Created with default configuration.'
      };
    } catch (error) {
      return { name: SETTINGS_LIST, status: 'error', message: this._msg(error) };
    }
  }

  /**
   * Ensures the navigation list exists with its columns, seeded with the
   * default links the header used to ship hard-coded.
   */
  private async _ensureNavigationList(): Promise<IProvisionResult> {
    try {
      const existed: boolean = await this._listExists(NAV_LIST);
      if (!existed) {
        await this.sp.web.lists.add(NAV_LIST, 'Top navigation links for the intranet header', 100);
        await this._hideList(NAV_LIST);
      }
      if (!(await this._fieldExists(NAV_LIST, NAV_URL_FIELD))) {
        await this.sp.web.lists.getByTitle(NAV_LIST).fields.addText(NAV_URL_FIELD);
      }
      if (!(await this._fieldExists(NAV_LIST, NAV_ORDER_FIELD))) {
        await this.sp.web.lists.getByTitle(NAV_LIST).fields.addNumber(NAV_ORDER_FIELD);
      }
      if (!(await this._fieldExists(NAV_LIST, NAV_NEWTAB_FIELD))) {
        await this.sp.web.lists.getByTitle(NAV_LIST).fields.addBoolean(NAV_NEWTAB_FIELD);
      }
      const count: number = (await this.sp.web.lists.getByTitle(NAV_LIST).items.top(1)()).length;
      if (count === 0) {
        let order: number = 1;
        for (const link of DEFAULT_NAV) {
          await this.sp.web.lists.getByTitle(NAV_LIST).items.add({
            Title: link.label,
            [NAV_URL_FIELD]: link.url,
            [NAV_ORDER_FIELD]: order,
            [NAV_NEWTAB_FIELD]: link.newTab
          });
          order += 1;
        }
      }
      return {
        name: NAV_LIST,
        status: existed ? 'skipped' : 'ok',
        message: existed ? 'Already existed — columns verified.' : 'Created and seeded with default links.'
      };
    } catch (error) {
      return { name: NAV_LIST, status: 'error', message: this._msg(error) };
    }
  }

  /** Hides a list from Site Contents, search and the quick launch. */
  private async _hideList(title: string): Promise<void> {
    try {
      await this.sp.web.lists.getByTitle(title).update({ Hidden: true, OnQuickLaunch: false, NoCrawl: true });
    } catch {
      /* non-fatal — the list is still usable if it could not be hidden. */
    }
  }

  /**
   * Creates the destination "View All" pages (News, Events, Directory, Policies)
   * with the matching web part pre-placed and its "View All" link switched off,
   * so the full-list page never shows a redundant button. Idempotent: existing
   * pages are left untouched.
   */
  private async _ensureViewAllPages(options: IProvisioningOptions): Promise<IProvisionResult[]> {
    const specs: { file: string; title: string; partId: string; props: Record<string, unknown> }[] = [
      {
        file: 'News',
        title: 'News & Announcements',
        partId: 'b975901c-fe7c-4efe-8c53-be86d2f9c691',
        props: { newsList: options.newsList, archiveUrl: '', maxItems: 24, showTitle: true, showViewAll: false }
      },
      {
        file: 'Events',
        title: 'Company Events',
        partId: 'e34006b6-56c7-4f1d-8cd6-395b909e7549',
        props: { eventsList: options.eventsList, calendarUrl: '', maxItems: '24', showTitle: true, showViewAll: false }
      },
      {
        file: 'Directory',
        title: 'Employee Directory',
        partId: '62faf2e1-930e-4cb8-acf3-6e2690952613',
        props: { fullDirectoryUrl: '', pageSize: 100, showTitle: true, showViewAll: false }
      },
      {
        file: 'Policies',
        title: 'Policies & Procedures',
        partId: '577e282d-a89f-48ea-baa8-b303a8865538',
        props: { policiesList: options.policiesList, allPoliciesUrl: '', showTitle: true, showViewAll: false }
      },
      {
        // Single shared detail page — News/Events/Employee/Policy cards link
        // here with ?type=&recId=. The web part reads those params itself.
        file: 'Detail',
        title: 'Detail',
        partId: DETAIL_VIEW_COMPONENT_ID,
        props: {
          newsList: options.newsList,
          eventsList: options.eventsList,
          policiesList: options.policiesList,
          benefitsList: options.benefitsList,
          backUrl: ''
        }
      },
      {
        // Contributor dashboard — add/edit/delete News, Events and HR Benefits.
        file: 'Content',
        title: 'Content Manager',
        partId: CONTENT_MANAGER_COMPONENT_ID,
        props: {
          title: 'Content Manager',
          newsList: options.newsList,
          eventsList: options.eventsList,
          benefitsList: options.benefitsList,
          showNews: true,
          showEvents: true,
          showBenefits: true
        }
      }
    ];

    let partDefs: IPartDef[];
    try {
      partDefs = await this.sp.web.getClientsideWebParts();
    } catch (error) {
      return [{ name: 'View All pages', status: 'error', message: `Could not read available web parts: ${this._msg(error)}` }];
    }

    const results: IProvisionResult[] = [];
    for (const spec of specs) {
      results.push(await this._ensurePage(options.pagesLibrary, spec, partDefs));
    }
    return results;
  }

  /**
   * Finds a client-side web part definition by component id. PnP returns the id
   * in `Id`, but depending on the part it may be wrapped in braces or differ in
   * case, and for SPFx parts the canonical id is also inside the `Manifest`
   * JSON — so we normalise and check both.
   */
  private _findPartDef(partDefs: IPartDef[], targetId: string): IPartDef | undefined {
    const norm = (v: string | undefined): string => (v || '').replace(/[{}]/g, '').toLowerCase();
    const target: string = norm(targetId);
    return partDefs.find((c) => {
      if (norm(c.Id) === target) {
        return true;
      }
      try {
        const manifest: { id?: string } = JSON.parse(c.Manifest || '{}');
        return norm(manifest.id) === target;
      } catch {
        return false;
      }
    });
  }

  private async _ensurePage(
    pagesLibrary: string,
    spec: { file: string; title: string; partId: string; props: Record<string, unknown> },
    partDefs: IPartDef[]
  ): Promise<IProvisionResult> {
    const fileName: string = `${spec.file}.aspx`;
    const name: string = `Page · ${spec.file}`;
    try {
      const existing: { Id: number }[] = await this.sp.web.lists
        .getByTitle(pagesLibrary)
        .items.filter(`FileLeafRef eq '${fileName}'`)
        .top(1)
        .select('Id')();
      if (existing && existing.length > 0) {
        return { name, status: 'skipped', message: 'Page already exists — left untouched.' };
      }

      const def: IPartDef | undefined = this._findPartDef(partDefs, spec.partId);
      if (!def) {
        return {
          name,
          status: 'error',
          message: `Web part ${spec.partId} not found among ${partDefs.length} available parts. ` +
            'Ensure v1.0.8.0 is deployed and the app is added/updated on this site, then re-run.'
        };
      }

      const page = await this.sp.web.addClientsidePage(fileName, spec.title, 'Article');
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const part = ClientsideWebpart.fromComponentDef(def as any);
      part.setProperties(spec.props);
      page.addSection().addControl(part);
      await page.save();
      return { name, status: 'ok', message: 'Created with the web part (View All hidden).' };
    } catch (error) {
      return { name, status: 'error', message: this._msg(error) };
    }
  }

  private async _seedEvents(listTitle: string): Promise<IProvisionResult> {
    try {
      const count: number = (await this.sp.web.lists.getByTitle(listTitle).items.top(1)()).length;
      if (count > 0) {
        return { name: `${listTitle} · sample data`, status: 'skipped', message: 'List already has items.' };
      }
      const list = this.sp.web.lists.getByTitle(listTitle);
      await list.items.add({
        Title: 'Annual Strategy Summit 2026',
        EventDate: '2026-06-18T09:00:00Z',
        EndDate: '2026-06-18T17:00:00Z',
        Location: 'Atlantis The Royal, Dubai',
        Category: 'All-Company'
      });
      await list.items.add({
        Title: 'Digital Transformation Workshop',
        EventDate: '2026-06-24T10:00:00Z',
        EndDate: '2026-06-24T13:00:00Z',
        Location: 'Virtual via Microsoft Teams',
        Category: 'Learning & Development'
      });
      return { name: `${listTitle} · sample data`, status: 'ok', message: 'Added 2 sample events.' };
    } catch (error) {
      return { name: `${listTitle} · sample data`, status: 'error', message: this._msg(error) };
    }
  }

  private async _seedNews(listTitle: string): Promise<IProvisionResult> {
    try {
      const count: number = (await this.sp.web.lists.getByTitle(listTitle).items.top(1)()).length;
      if (count > 0) {
        return { name: `${listTitle} · sample data`, status: 'skipped', message: 'List already has items.' };
      }
      const list = this.sp.web.lists.getByTitle(listTitle);
      await list.items.add({
        Title: "2026 Employee Benefits: What's New This Year",
        Category: 'HR Updates',
        NewsDate: '2026-07-01T08:00:00Z',
        Source: 'People & Culture',
        Body:
          "<p>We're pleased to announce enhancements to your 2026 benefits package, effective 1 July 2026.</p>" +
          '<ul><li><strong>Medical:</strong> Upgraded regional cover including dependents, with dental and optical.</li>' +
          '<li><strong>Wellness allowance:</strong> AED 2,000 per year toward fitness or mental-health services.</li>' +
          '<li><strong>Annual leave:</strong> +2 days for employees with 3+ years of service.</li>' +
          '<li><strong>Learning:</strong> AED 5,000 annual professional-development budget.</li></ul>' +
          '<p>Enrolment opens in the HR portal from 1 July.</p>'
      });
      await list.items.add({
        Title: 'Breaking Ground on the New Waterfront District',
        Category: 'Construction Updates',
        NewsDate: '2026-06-20T08:00:00Z',
        Source: 'Development Office'
      });
      await list.items.add({
        Title: 'Q2 Financial Results Exceed Targets',
        Category: 'Finance',
        NewsDate: '2026-06-10T08:00:00Z',
        Source: 'Finance'
      });
      return { name: `${listTitle} · sample data`, status: 'ok', message: 'Added 3 sample news items.' };
    } catch (error) {
      return { name: `${listTitle} · sample data`, status: 'error', message: this._msg(error) };
    }
  }

  private async _seedBenefits(listTitle: string): Promise<IProvisionResult> {
    try {
      const count: number = (await this.sp.web.lists.getByTitle(listTitle).items.top(1)()).length;
      if (count > 0) {
        return { name: `${listTitle} · sample data`, status: 'skipped', message: 'List already has items.' };
      }
      const list = this.sp.web.lists.getByTitle(listTitle);
      await list.items.add({
        Title: 'Comprehensive Medical Insurance',
        Category: 'Health',
        Summary: 'Regional health cover for you and your dependents, including dental and optical.',
        Eligibility: 'All full-time employees from day one.',
        Coverage: 'Employee + spouse + up to 3 children',
        Details:
          '<p>Full inpatient and outpatient cover across the region, with direct billing at network ' +
          'hospitals and clinics. Dental and optical benefits are included up to annual limits.</p>'
      });
      await list.items.add({
        Title: 'Annual Wellness Allowance',
        Category: 'Wellness',
        Summary: 'AED 2,000 per year toward gym, fitness and mental-health services.',
        Eligibility: 'All employees after probation.',
        Coverage: 'AED 2,000 / year',
        Details: '<p>Reimbursable against gym memberships, fitness classes, and approved mental-health support.</p>'
      });
      await list.items.add({
        Title: 'Professional Development Budget',
        Category: 'Learning',
        Summary: 'AED 5,000 per year for courses, certifications and conferences.',
        Eligibility: 'All employees after probation.',
        Coverage: 'AED 5,000 / year',
        Details: '<p>Covers role-relevant training, professional certifications, and conference attendance with manager approval.</p>'
      });
      await list.items.add({
        Title: 'Enhanced Parental Leave',
        Category: 'Family',
        Summary: 'Extended paid maternity and paternity leave beyond statutory minimums.',
        Eligibility: 'Employees with 1+ year of service.',
        Coverage: 'Up to 90 days maternity / 10 days paternity',
        Details: '<p>Paid parental leave with a phased return-to-work option in the first month back.</p>'
      });
      return { name: `${listTitle} · sample data`, status: 'ok', message: 'Added 4 sample benefits.' };
    } catch (error) {
      return { name: `${listTitle} · sample data`, status: 'error', message: this._msg(error) };
    }
  }

  private _msg(error: unknown): string {
    return error instanceof Error ? error.message : String(error);
  }
}
