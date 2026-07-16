import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/user-custom-actions';
import '@pnp/sp/folders';
import '@pnp/sp/files';
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
  NAV_LABEL_AR_FIELD,
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

const HEADER_FOOTER_COMPONENT_ID: string = '1083fb8f-96fc-49dc-9284-2d77402b58a9';
const DETAIL_VIEW_COMPONENT_ID: string = 'e3f3748b-536f-4ee1-bcf8-0be1a0e71437';
const CONTENT_MANAGER_COMPONENT_ID: string = 'e7767646-0e20-4cd6-ab84-5e2c44f05673';

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
  docCenterLibrary: string;
  formsLibrary: string;
  templatesLibrary: string;
  pagesLibrary: string;
  registerHeaderFooter: boolean;
  seedSampleData: boolean;
  /**
   * When true (and seedSampleData is also true), every existing item in the
   * seeded content lists (News, Events, HR Benefits) is deleted before the
   * sample data is re-added — a destructive "reset to sample content". Ignored
   * when seedSampleData is false. Defaults to false.
   */
  resetSampleData: boolean;
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
        // Arabic translations — same content, rendered when the visitor's language is Arabic.
        await this._ensureTextField(listTitle, 'TitleAR');
        await this._ensureTextField(listTitle, 'LocationAR');
        await this._ensureMultilineField(listTitle, 'DescriptionAR');
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
        // Arabic translation of the policy title (the version/department taxonomy stays English).
        await this._ensureTextField(listTitle, 'TitleAR');
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
        // Arabic translations — same content, rendered when the visitor's language is Arabic.
        await this._ensureTextField(listTitle, 'TitleAR');
        await this._ensureTextField(listTitle, 'SourceAR');
        await this._ensureMultilineField(listTitle, 'BodyAR');
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
        // Arabic translations — same content, rendered when the visitor's language is Arabic.
        await this._ensureTextField(listTitle, 'TitleAR');
        await this._ensureTextField(listTitle, 'SummaryAR');
        await this._ensureTextField(listTitle, 'EligibilityAR');
        await this._ensureTextField(listTitle, 'CoverageAR');
        await this._ensureMultilineField(listTitle, 'DetailsAR');
      })
    );

    results.push(
      // Document Center — a real document library (101) with DMS metadata columns.
      // The custom Document Center web part renders a search/filter experience on
      // top of it while SharePoint provides versioning, check-in/out and permissions.
      await this._ensureList(options.docCenterLibrary, 101, 'Managed documents for the Document Center web part', async (listTitle) => {
        // "Category" holds the owning department; the web part labels it Department.
        await this._ensureChoiceField(listTitle, 'Category', ['HR', 'Finance', 'IT', 'Operations', 'Marketing', 'Legal', 'General']);
        await this._ensureChoiceField(listTitle, 'DocumentType', ['Policy', 'Procedure', 'Form', 'Template', 'Report', 'Guideline', 'Contract']);
        // Stage/status through the document lifecycle.
        await this._ensureChoiceField(listTitle, 'DocStatus', ['Draft', 'In Review', 'Published', 'Archived']);
        await this._ensureChoiceField(listTitle, 'Sensitivity', ['Public', 'Internal', 'Confidential', 'Restricted']);
        await this._ensureTextField(listTitle, 'DocumentNumber');
        await this._ensureTextField(listTitle, 'DocTags');
        await this._ensureTextField(listTitle, 'DocOwner');
        await this._ensureDateField(listTitle, 'ReviewDate');
        await this._ensureMultilineField(listTitle, 'Description');
        // Arabic translations — same content, rendered when the visitor's language is Arabic.
        await this._ensureTextField(listTitle, 'TitleAR');
        await this._ensureMultilineField(listTitle, 'DescriptionAR');
        // Turn on version history so "Publish & versions" / "Full history" work.
        try {
          await this.sp.web.lists.getByTitle(listTitle).update({ EnableVersioning: true, EnableMinorVersions: false });
        } catch {
          /* non-fatal — versioning can be enabled from library settings if this fails. */
        }
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
      // Reset first (when requested): wipe the seeded content lists so the seed
      // methods — which skip a non-empty list — re-add a clean sample set.
      if (options.resetSampleData) {
        results.push(await this._clearList(options.eventsList));
        results.push(await this._clearList(options.newsList));
        results.push(await this._clearList(options.benefitsList));
        results.push(await this._clearList(options.docCenterLibrary));
      }
      results.push(await this._seedEvents(options.eventsList));
      // Policies is a document library — seeding needs real files, so it's
      // skipped; upload policy documents to the library instead.
      results.push(await this._seedNews(options.newsList));
      results.push(await this._seedBenefits(options.benefitsList));
      results.push(await this._seedDocuments(options.docCenterLibrary));
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
        Title: 'IntranetHeaderFooter',
        Name: 'IntranetHeaderFooter',
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
      if (!(await this._fieldExists(NAV_LIST, NAV_LABEL_AR_FIELD))) {
        await this.sp.web.lists.getByTitle(NAV_LIST).fields.addText(NAV_LABEL_AR_FIELD);
      }
      const count: number = (await this.sp.web.lists.getByTitle(NAV_LIST).items.top(1)()).length;
      if (count === 0) {
        let order: number = 1;
        for (const link of DEFAULT_NAV) {
          await this.sp.web.lists.getByTitle(NAV_LIST).items.add({
            Title: link.label,
            [NAV_LABEL_AR_FIELD]: link.labelAR || '',
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
        partId: '71d6cf01-a010-4201-b219-9d6882dcc243',
        props: { newsList: options.newsList, archiveUrl: '', maxItems: 24, showTitle: true, showViewAll: false }
      },
      {
        file: 'Events',
        title: 'Company Events',
        partId: '86a3b1ee-7bbb-4d88-b6a9-b2f65ba2f737',
        props: { eventsList: options.eventsList, calendarUrl: '', maxItems: '24', showTitle: true, showViewAll: false }
      },
      {
        file: 'Directory',
        title: 'Employee Directory',
        partId: '80250f4d-704a-4320-94ea-a7b128cecec0',
        props: { fullDirectoryUrl: '', pageSize: 100, showTitle: true, showViewAll: false }
      },
      {
        file: 'Policies',
        title: 'Policies & Procedures',
        partId: 'dd4fd601-bbc6-4fd5-99ee-e441aa28bc65',
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

  /**
   * Deletes every item in a list — used by the reset option before re-seeding.
   * Pages through the list (100 at a time) so it works past the default view
   * threshold, with a safety cap so a misbehaving delete can never loop
   * forever. Idempotent: a missing or already-empty list is a no-op.
   */
  private async _clearList(listTitle: string): Promise<IProvisionResult> {
    const name: string = `${listTitle} · reset`;
    try {
      if (!(await this._listExists(listTitle))) {
        return { name, status: 'skipped', message: 'List does not exist yet — nothing to reset.' };
      }
      const list = this.sp.web.lists.getByTitle(listTitle);
      let deleted: number = 0;
      // SharePoint frees ids as we delete, so keep pulling the first page until
      // none remain. The guard bounds this to 200 pages (20k items) as a
      // backstop against an unexpected non-terminating delete.
      for (let guard = 0; guard < 200; guard++) {
        const rows: { Id: number }[] = await list.items.select('Id').top(100)();
        if (!rows || rows.length === 0) {
          break;
        }
        for (const row of rows) {
          await list.items.getById(row.Id).delete();
          deleted += 1;
        }
      }
      return {
        name,
        status: 'ok',
        message: deleted > 0 ? `Deleted ${deleted} existing item(s) before re-seeding.` : 'No existing items to delete.'
      };
    } catch (error) {
      return { name, status: 'error', message: this._msg(error) };
    }
  }

  /**
   * Seeds the Document Center library with a handful of sample documents (real
   * placeholder .txt files) carrying full bilingual metadata, so the Document
   * Center web part renders populated out of the box. Only runs on an empty
   * library; each upload is best-effort so one failure can't abort the rest.
   */
  private async _seedDocuments(listTitle: string): Promise<IProvisionResult> {
    const name: string = `${listTitle} · sample data`;
    try {
      const count: number = (await this.sp.web.lists.getByTitle(listTitle).items.top(1)()).length;
      if (count > 0) {
        return { name, status: 'skipped', message: 'Library already has items.' };
      }
      const list = this.sp.web.lists.getByTitle(listTitle);
      const body: string = 'Sample document — replace this placeholder with the real file. Managed via the Document Center.';
      const samples: {
        file: string;
        Title: string;
        TitleAR: string;
        DocumentNumber: string;
        Category: string;
        DocumentType: string;
        DocStatus: string;
        Sensitivity: string;
        DocTags: string;
        DocOwner: string;
        ReviewDate: string;
        Description: string;
        DescriptionAR: string;
      }[] = [
        {
          file: 'Employee-Handbook-2026.txt',
          Title: 'Employee Handbook 2026',
          TitleAR: 'دليل الموظف ٢٠٢٦',
          DocumentNumber: 'HR-2026-0001',
          Category: 'HR',
          DocumentType: 'Policy',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'handbook; policy; benefits',
          DocOwner: 'People & Culture',
          ReviewDate: '2026-12-31T00:00:00Z',
          Description: 'The complete guide to company policies, benefits and ways of working.',
          DescriptionAR: 'الدليل الكامل لسياسات الشركة والمزايا وأساليب العمل.'
        },
        {
          file: 'Expense-Claim-Form.txt',
          Title: 'Expense Claim Form',
          TitleAR: 'نموذج مطالبة النفقات',
          DocumentNumber: 'FIN-2026-0001',
          Category: 'Finance',
          DocumentType: 'Form',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'expense; claim; reimbursement',
          DocOwner: 'Finance',
          ReviewDate: '2026-09-30T00:00:00Z',
          Description: 'Submit business expenses for reimbursement.',
          DescriptionAR: 'تقديم نفقات العمل لاستردادها.'
        },
        {
          file: 'IT-Security-Guidelines.txt',
          Title: 'IT Security Guidelines',
          TitleAR: 'إرشادات أمن تقنية المعلومات',
          DocumentNumber: 'IT-2026-0001',
          Category: 'IT',
          DocumentType: 'Guideline',
          DocStatus: 'In Review',
          Sensitivity: 'Confidential',
          DocTags: 'security; passwords; devices',
          DocOwner: 'Information Technology',
          ReviewDate: '2026-08-15T00:00:00Z',
          Description: 'Security best practices for devices, passwords and data handling.',
          DescriptionAR: 'أفضل ممارسات الأمان للأجهزة وكلمات المرور والتعامل مع البيانات.'
        },
        {
          file: 'Project-Status-Report-Template.txt',
          Title: 'Project Status Report Template',
          TitleAR: 'قالب تقرير حالة المشروع',
          DocumentNumber: 'OPS-2026-0001',
          Category: 'Operations',
          DocumentType: 'Template',
          DocStatus: 'Draft',
          Sensitivity: 'Internal',
          DocTags: 'project; report; template',
          DocOwner: 'PMO',
          ReviewDate: '2026-10-31T00:00:00Z',
          Description: 'Standard template for weekly project status reporting.',
          DescriptionAR: 'قالب موحّد لإعداد تقارير حالة المشروع الأسبوعية.'
        },
        {
          file: 'Vendor-Contract-Template.txt',
          Title: 'Vendor Contract Template',
          TitleAR: 'قالب عقد المورّد',
          DocumentNumber: 'LEG-2026-0001',
          Category: 'Legal',
          DocumentType: 'Contract',
          DocStatus: 'Published',
          Sensitivity: 'Confidential',
          DocTags: 'contract; vendor; agreement',
          DocOwner: 'Legal',
          ReviewDate: '2027-01-31T00:00:00Z',
          Description: 'Standard agreement template for engaging third-party vendors.',
          DescriptionAR: 'قالب الاتفاقية الموحّد للتعاقد مع الموردين الخارجيين.'
        },
        {
          file: 'Onboarding-Procedure.txt',
          Title: 'Onboarding Procedure',
          TitleAR: 'إجراء التأهيل الوظيفي',
          DocumentNumber: 'HR-2026-0002',
          Category: 'HR',
          DocumentType: 'Procedure',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'onboarding; joiner; process',
          DocOwner: 'People & Culture',
          ReviewDate: '2026-11-30T00:00:00Z',
          Description: 'Step-by-step process for welcoming and setting up new joiners.',
          DescriptionAR: 'عملية تفصيلية لاستقبال الموظفين الجدد وتجهيزهم.'
        },
        {
          file: 'Leave-Policy.txt',
          Title: 'Leave & Time-Off Policy',
          TitleAR: 'سياسة الإجازات والوقت المستقطع',
          DocumentNumber: 'HR-2026-0003',
          Category: 'HR',
          DocumentType: 'Policy',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'leave; vacation; policy',
          DocOwner: 'People & Culture',
          ReviewDate: '2027-03-31T00:00:00Z',
          Description: 'Entitlements and rules for annual, sick and special leave.',
          DescriptionAR: 'الاستحقاقات وقواعد الإجازات السنوية والمرضية والخاصة.'
        },
        {
          file: 'Performance-Review-Form.txt',
          Title: 'Performance Review Form',
          TitleAR: 'نموذج تقييم الأداء',
          DocumentNumber: 'HR-2026-0004',
          Category: 'HR',
          DocumentType: 'Form',
          DocStatus: 'In Review',
          Sensitivity: 'Confidential',
          DocTags: 'performance; appraisal; form',
          DocOwner: 'People & Culture',
          ReviewDate: '2026-07-01T00:00:00Z',
          Description: 'Annual performance appraisal form for managers and staff.',
          DescriptionAR: 'نموذج تقييم الأداء السنوي للمديرين والموظفين.'
        },
        {
          file: 'Budget-Planning-Report.txt',
          Title: 'Annual Budget Planning Report',
          TitleAR: 'تقرير تخطيط الميزانية السنوية',
          DocumentNumber: 'FIN-2026-0002',
          Category: 'Finance',
          DocumentType: 'Report',
          DocStatus: 'Published',
          Sensitivity: 'Confidential',
          DocTags: 'budget; planning; finance',
          DocOwner: 'Finance',
          ReviewDate: '2027-01-15T00:00:00Z',
          Description: 'Consolidated budget planning and forecasts for the financial year.',
          DescriptionAR: 'تخطيط الميزانية الموحّد والتوقّعات المالية للسنة المالية.'
        },
        {
          file: 'Procurement-Policy.txt',
          Title: 'Procurement Policy',
          TitleAR: 'سياسة المشتريات',
          DocumentNumber: 'FIN-2026-0003',
          Category: 'Finance',
          DocumentType: 'Policy',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'procurement; purchasing; policy',
          DocOwner: 'Finance',
          ReviewDate: '2026-10-01T00:00:00Z',
          Description: 'Rules and approval thresholds for purchasing goods and services.',
          DescriptionAR: 'قواعد وحدود الموافقة لشراء السلع والخدمات.'
        },
        {
          file: 'Invoice-Template.txt',
          Title: 'Standard Invoice Template',
          TitleAR: 'قالب الفاتورة القياسي',
          DocumentNumber: 'FIN-2026-0004',
          Category: 'Finance',
          DocumentType: 'Template',
          DocStatus: 'Draft',
          Sensitivity: 'Internal',
          DocTags: 'invoice; billing; template',
          DocOwner: 'Finance',
          ReviewDate: '2026-09-01T00:00:00Z',
          Description: 'Company-branded invoice template for client billing.',
          DescriptionAR: 'قالب فاتورة يحمل هوية الشركة لإصدار فواتير العملاء.'
        },
        {
          file: 'Data-Backup-Procedure.txt',
          Title: 'Data Backup & Recovery Procedure',
          TitleAR: 'إجراء النسخ الاحتياطي واستعادة البيانات',
          DocumentNumber: 'IT-2026-0002',
          Category: 'IT',
          DocumentType: 'Procedure',
          DocStatus: 'Published',
          Sensitivity: 'Confidential',
          DocTags: 'backup; recovery; data',
          DocOwner: 'Information Technology',
          ReviewDate: '2026-08-31T00:00:00Z',
          Description: 'How data is backed up and restored, with recovery-time targets.',
          DescriptionAR: 'كيفية النسخ الاحتياطي للبيانات واستعادتها مع أهداف زمن الاستعادة.'
        },
        {
          file: 'Acceptable-Use-Policy.txt',
          Title: 'Acceptable Use Policy',
          TitleAR: 'سياسة الاستخدام المقبول',
          DocumentNumber: 'IT-2026-0003',
          Category: 'IT',
          DocumentType: 'Policy',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'acceptable use; devices; policy',
          DocOwner: 'Information Technology',
          ReviewDate: '2027-02-28T00:00:00Z',
          Description: 'Acceptable use of company systems, devices and internet access.',
          DescriptionAR: 'الاستخدام المقبول لأنظمة الشركة وأجهزتها والوصول إلى الإنترنت.'
        },
        {
          file: 'Software-Request-Form.txt',
          Title: 'Software Request Form',
          TitleAR: 'نموذج طلب البرمجيات',
          DocumentNumber: 'IT-2026-0004',
          Category: 'IT',
          DocumentType: 'Form',
          DocStatus: 'Draft',
          Sensitivity: 'Internal',
          DocTags: 'software; request; form',
          DocOwner: 'Information Technology',
          ReviewDate: '2026-07-20T00:00:00Z',
          Description: 'Request new software licences or access to applications.',
          DescriptionAR: 'طلب تراخيص برمجيات جديدة أو الوصول إلى التطبيقات.'
        },
        {
          file: 'Health-and-Safety-Guideline.txt',
          Title: 'Health & Safety Guideline',
          TitleAR: 'إرشادات الصحة والسلامة',
          DocumentNumber: 'OPS-2026-0002',
          Category: 'Operations',
          DocumentType: 'Guideline',
          DocStatus: 'Published',
          Sensitivity: 'Public',
          DocTags: 'health; safety; workplace',
          DocOwner: 'Operations',
          ReviewDate: '2026-12-15T00:00:00Z',
          Description: 'Workplace health and safety practices for all staff and visitors.',
          DescriptionAR: 'ممارسات الصحة والسلامة في مكان العمل لجميع الموظفين والزوار.'
        },
        {
          file: 'Facilities-Booking-Form.txt',
          Title: 'Facilities Booking Form',
          TitleAR: 'نموذج حجز المرافق',
          DocumentNumber: 'OPS-2026-0003',
          Category: 'Operations',
          DocumentType: 'Form',
          DocStatus: 'In Review',
          Sensitivity: 'Internal',
          DocTags: 'facilities; booking; rooms',
          DocOwner: 'Operations',
          ReviewDate: '2026-08-01T00:00:00Z',
          Description: 'Reserve meeting rooms, equipment and shared facilities.',
          DescriptionAR: 'حجز قاعات الاجتماعات والمعدّات والمرافق المشتركة.'
        },
        {
          file: 'Brand-Guidelines.txt',
          Title: 'Brand Guidelines',
          TitleAR: 'إرشادات العلامة التجارية',
          DocumentNumber: 'MKT-2026-0001',
          Category: 'Marketing',
          DocumentType: 'Guideline',
          DocStatus: 'Published',
          Sensitivity: 'Public',
          DocTags: 'brand; logo; guidelines',
          DocOwner: 'Marketing',
          ReviewDate: '2027-04-30T00:00:00Z',
          Description: 'Logo usage, colours, typography and tone of voice.',
          DescriptionAR: 'استخدام الشعار والألوان والخطوط ونبرة الخطاب.'
        },
        {
          file: 'Campaign-Report-Q2.txt',
          Title: 'Q2 Campaign Performance Report',
          TitleAR: 'تقرير أداء حملات الربع الثاني',
          DocumentNumber: 'MKT-2026-0002',
          Category: 'Marketing',
          DocumentType: 'Report',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'campaign; report; marketing',
          DocOwner: 'Marketing',
          ReviewDate: '2026-09-15T00:00:00Z',
          Description: 'Results and insights from the second-quarter marketing campaigns.',
          DescriptionAR: 'نتائج ورؤى حملات التسويق في الربع الثاني.'
        },
        {
          file: 'Social-Media-Policy.txt',
          Title: 'Social Media Policy',
          TitleAR: 'سياسة وسائل التواصل الاجتماعي',
          DocumentNumber: 'MKT-2026-0003',
          Category: 'Marketing',
          DocumentType: 'Policy',
          DocStatus: 'Draft',
          Sensitivity: 'Internal',
          DocTags: 'social media; policy; conduct',
          DocOwner: 'Marketing',
          ReviewDate: '2026-07-10T00:00:00Z',
          Description: 'Guidelines for representing the company on social platforms.',
          DescriptionAR: 'إرشادات تمثيل الشركة على منصات التواصل الاجتماعي.'
        },
        {
          file: 'NDA-Template.txt',
          Title: 'Non-Disclosure Agreement Template',
          TitleAR: 'قالب اتفاقية عدم الإفصاح',
          DocumentNumber: 'LEG-2026-0002',
          Category: 'Legal',
          DocumentType: 'Template',
          DocStatus: 'Published',
          Sensitivity: 'Confidential',
          DocTags: 'nda; confidentiality; template',
          DocOwner: 'Legal',
          ReviewDate: '2027-01-01T00:00:00Z',
          Description: 'Standard mutual non-disclosure agreement for partners.',
          DescriptionAR: 'اتفاقية عدم إفصاح متبادلة موحّدة للشركاء.'
        },
        {
          file: 'Compliance-Policy.txt',
          Title: 'Regulatory Compliance Policy',
          TitleAR: 'سياسة الامتثال التنظيمي',
          DocumentNumber: 'LEG-2026-0003',
          Category: 'Legal',
          DocumentType: 'Policy',
          DocStatus: 'In Review',
          Sensitivity: 'Restricted',
          DocTags: 'compliance; regulatory; policy',
          DocOwner: 'Legal',
          ReviewDate: '2026-06-30T00:00:00Z',
          Description: 'Company obligations under applicable laws and regulations.',
          DescriptionAR: 'التزامات الشركة بموجب القوانين واللوائح المعمول بها.'
        },
        {
          file: 'Company-Org-Chart.txt',
          Title: 'Company Organisation Chart',
          TitleAR: 'الهيكل التنظيمي للشركة',
          DocumentNumber: 'GEN-2026-0001',
          Category: 'General',
          DocumentType: 'Report',
          DocStatus: 'Published',
          Sensitivity: 'Internal',
          DocTags: 'org chart; structure; teams',
          DocOwner: 'People & Culture',
          ReviewDate: '2026-12-01T00:00:00Z',
          Description: 'Current organisation structure and reporting lines.',
          DescriptionAR: 'الهيكل التنظيمي الحالي وخطوط إعداد التقارير.'
        },
        {
          file: 'Meeting-Minutes-Template.txt',
          Title: 'Meeting Minutes Template',
          TitleAR: 'قالب محضر الاجتماع',
          DocumentNumber: 'GEN-2026-0002',
          Category: 'General',
          DocumentType: 'Template',
          DocStatus: 'Published',
          Sensitivity: 'Public',
          DocTags: 'meeting; minutes; template',
          DocOwner: 'Operations',
          ReviewDate: '2027-02-01T00:00:00Z',
          Description: 'Standard template for recording meeting minutes and actions.',
          DescriptionAR: 'قالب موحّد لتدوين محاضر الاجتماعات والإجراءات.'
        },
        {
          file: 'Business-Continuity-Plan.txt',
          Title: 'Business Continuity Plan',
          TitleAR: 'خطة استمرارية الأعمال',
          DocumentNumber: 'GEN-2026-0003',
          Category: 'General',
          DocumentType: 'Guideline',
          DocStatus: 'Archived',
          Sensitivity: 'Restricted',
          DocTags: 'continuity; risk; plan',
          DocOwner: 'Operations',
          ReviewDate: '2026-05-31T00:00:00Z',
          Description: 'How the business continues operating during a major disruption.',
          DescriptionAR: 'كيفية استمرار العمل أثناء أي اضطراب كبير.'
        }
      ];

      let uploaded: number = 0;
      let firstError: string = '';
      for (const s of samples) {
        try {
          // PnP v4 returns the file info directly; resolve the list item from its URL.
          const info: { ServerRelativeUrl?: string } = await list.rootFolder.files.addUsingPath(s.file, body, { Overwrite: true });
          uploaded += 1;
          const serverRelativeUrl: string = info.ServerRelativeUrl || '';
          const item = await this.sp.web.getFileByServerRelativePath(serverRelativeUrl).getItem();
          const fields: Record<string, unknown> = {
            Title: s.Title,
            TitleAR: s.TitleAR,
            DocumentNumber: s.DocumentNumber,
            Category: s.Category,
            DocumentType: s.DocumentType,
            DocStatus: s.DocStatus,
            Sensitivity: s.Sensitivity,
            DocTags: s.DocTags,
            DocOwner: s.DocOwner,
            ReviewDate: s.ReviewDate,
            Description: s.Description,
            DescriptionAR: s.DescriptionAR
          };
          try {
            // Fast path: set all metadata in one call.
            await item.update(fields);
          } catch (bulkError) {
            // A single mismatched column would fail the whole batch — fall back to
            // per-field updates so the file still gets whatever metadata is valid,
            // and remember the first error for reporting.
            if (!firstError) {
              firstError = this._msg(bulkError);
            }
            for (const key of Object.keys(fields)) {
              try {
                await item.update({ [key]: fields[key] });
              } catch {
                /* skip a field that doesn't accept the value / doesn't exist */
              }
            }
          }
        } catch (uploadError) {
          if (!firstError) {
            firstError = this._msg(uploadError);
          }
        }
      }
      if (uploaded > 0) {
        return {
          name,
          status: 'ok',
          message: firstError
            ? `Uploaded ${uploaded} sample document(s); some metadata was skipped (${firstError}).`
            : `Added ${uploaded} sample document(s).`
        };
      }
      return {
        name,
        status: 'error',
        message: firstError ? `Could not upload sample documents: ${firstError}` : 'Could not upload sample documents (check library permissions).'
      };
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
        TitleAR: 'قمة الاستراتيجية السنوية 2026',
        EventDate: '2026-06-18T09:00:00Z',
        EndDate: '2026-06-18T17:00:00Z',
        Location: 'Atlantis The Royal, Dubai',
        LocationAR: 'أتلانتس ذا رويال، دبي',
        Category: 'All-Company'
      });
      await list.items.add({
        Title: 'Digital Transformation Workshop',
        TitleAR: 'ورشة عمل التحول الرقمي',
        EventDate: '2026-06-24T10:00:00Z',
        EndDate: '2026-06-24T13:00:00Z',
        Location: 'Virtual via Microsoft Teams',
        LocationAR: 'افتراضي عبر مايكروسوفت تيمز',
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
        TitleAR: 'مزايا الموظفين لعام 2026: الجديد هذا العام',
        Category: 'HR Updates',
        NewsDate: '2026-07-01T08:00:00Z',
        Source: 'People & Culture',
        SourceAR: 'إدارة الموارد البشرية والثقافة',
        Body:
          "<p>We're pleased to announce enhancements to your 2026 benefits package, effective 1 July 2026.</p>" +
          '<ul><li><strong>Medical:</strong> Upgraded regional cover including dependents, with dental and optical.</li>' +
          '<li><strong>Wellness allowance:</strong> AED 2,000 per year toward fitness or mental-health services.</li>' +
          '<li><strong>Annual leave:</strong> +2 days for employees with 3+ years of service.</li>' +
          '<li><strong>Learning:</strong> AED 5,000 annual professional-development budget.</li></ul>' +
          '<p>Enrolment opens in the HR portal from 1 July.</p>',
        BodyAR:
          '<p>يسرّنا أن نعلن عن تحسينات على باقة مزاياكم لعام 2026، اعتباراً من 1 يوليو 2026.</p>' +
          '<ul><li><strong>التأمين الطبي:</strong> تغطية إقليمية مُحدّثة تشمل المُعالين، مع تغطية طب الأسنان والبصريات.</li>' +
          '<li><strong>بدل العافية:</strong> 2,000 درهم سنوياً لخدمات اللياقة البدنية أو الصحة النفسية.</li>' +
          '<li><strong>الإجازة السنوية:</strong> يومان إضافيان للموظفين ذوي الخدمة التي تزيد على 3 سنوات.</li>' +
          '<li><strong>التعلّم:</strong> ميزانية سنوية للتطوير المهني قدرها 5,000 درهم.</li></ul>' +
          '<p>يبدأ التسجيل في بوابة الموارد البشرية اعتباراً من 1 يوليو.</p>'
      });
      await list.items.add({
        Title: 'Breaking Ground on the New Waterfront District',
        TitleAR: 'وضع حجر الأساس لحي الواجهة البحرية الجديد',
        Category: 'Construction Updates',
        NewsDate: '2026-06-20T08:00:00Z',
        Source: 'Development Office',
        SourceAR: 'مكتب التطوير'
      });
      await list.items.add({
        Title: 'Q2 Financial Results Exceed Targets',
        TitleAR: 'النتائج المالية للربع الثاني تتجاوز المستهدفات',
        Category: 'Finance',
        NewsDate: '2026-06-10T08:00:00Z',
        Source: 'Finance',
        SourceAR: 'الإدارة المالية'
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
        TitleAR: 'تأمين طبي شامل',
        Category: 'Health',
        Summary: 'Regional health cover for you and your dependents, including dental and optical.',
        SummaryAR: 'تغطية صحية إقليمية لك ولمُعاليك، تشمل طب الأسنان والبصريات.',
        Eligibility: 'All full-time employees from day one.',
        EligibilityAR: 'جميع الموظفين بدوام كامل من اليوم الأول.',
        Coverage: 'Employee + spouse + up to 3 children',
        CoverageAR: 'الموظف + الزوج/الزوجة + حتى 3 أطفال',
        Details:
          '<p>Full inpatient and outpatient cover across the region, with direct billing at network ' +
          'hospitals and clinics. Dental and optical benefits are included up to annual limits.</p>',
        DetailsAR:
          '<p>تغطية كاملة للعلاج داخل المستشفى وخارجه على مستوى المنطقة، مع الدفع المباشر لدى المستشفيات ' +
          'والعيادات ضمن الشبكة. تشمل التغطية مزايا طب الأسنان والبصريات ضمن الحدود السنوية.</p>'
      });
      await list.items.add({
        Title: 'Annual Wellness Allowance',
        TitleAR: 'بدل العافية السنوي',
        Category: 'Wellness',
        Summary: 'AED 2,000 per year toward gym, fitness and mental-health services.',
        SummaryAR: '2,000 درهم سنوياً لخدمات النادي الرياضي واللياقة البدنية والصحة النفسية.',
        Eligibility: 'All employees after probation.',
        EligibilityAR: 'جميع الموظفين بعد فترة الاختبار.',
        Coverage: 'AED 2,000 / year',
        CoverageAR: '2,000 درهم / سنوياً',
        Details: '<p>Reimbursable against gym memberships, fitness classes, and approved mental-health support.</p>',
        DetailsAR: '<p>قابل للاسترداد مقابل اشتراكات النادي الرياضي وحصص اللياقة البدنية والدعم النفسي المعتمد.</p>'
      });
      await list.items.add({
        Title: 'Professional Development Budget',
        TitleAR: 'ميزانية التطوير المهني',
        Category: 'Learning',
        Summary: 'AED 5,000 per year for courses, certifications and conferences.',
        SummaryAR: '5,000 درهم سنوياً للدورات والشهادات والمؤتمرات.',
        Eligibility: 'All employees after probation.',
        EligibilityAR: 'جميع الموظفين بعد فترة الاختبار.',
        Coverage: 'AED 5,000 / year',
        CoverageAR: '5,000 درهم / سنوياً',
        Details: '<p>Covers role-relevant training, professional certifications, and conference attendance with manager approval.</p>',
        DetailsAR: '<p>يغطي التدريب المرتبط بالوظيفة والشهادات المهنية وحضور المؤتمرات بموافقة المدير.</p>'
      });
      await list.items.add({
        Title: 'Enhanced Parental Leave',
        TitleAR: 'إجازة والدية مُعزّزة',
        Category: 'Family',
        Summary: 'Extended paid maternity and paternity leave beyond statutory minimums.',
        SummaryAR: 'إجازة أمومة وأبوة مدفوعة الأجر ممتدة تتجاوز الحدود الدنيا القانونية.',
        Eligibility: 'Employees with 1+ year of service.',
        EligibilityAR: 'الموظفون ذوو الخدمة التي تزيد على سنة واحدة.',
        Coverage: 'Up to 90 days maternity / 10 days paternity',
        CoverageAR: 'حتى 90 يوماً إجازة أمومة / 10 أيام إجازة أبوة',
        Details: '<p>Paid parental leave with a phased return-to-work option in the first month back.</p>',
        DetailsAR: '<p>إجازة والدية مدفوعة الأجر مع خيار العودة التدريجية إلى العمل خلال الشهر الأول من العودة.</p>'
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
