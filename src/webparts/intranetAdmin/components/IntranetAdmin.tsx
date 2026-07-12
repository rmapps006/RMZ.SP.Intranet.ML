import * as React from 'react';
import styles from './IntranetAdmin.module.scss';
import { IIntranetAdminProps } from './IIntranetAdminProps';
import { ProvisioningService, IProvisionResult } from '../services/ProvisioningService';
import { getSP } from '../../../common/services/pnpService';
import {
  SettingsService,
  IIntranetSettings,
  INavLink,
  IQuickLinkSetting,
  IDirectoryDomain,
  IDepartmentEntry,
  DEFAULT_SETTINGS,
  getCachedSettings
} from '../../../common/services/SettingsService';
import {
  Pivot,
  PivotItem,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

type SaveState = 'idle' | 'saving' | 'saved' | 'error';

const IntranetAdmin: React.FunctionComponent<IIntranetAdminProps> = (props) => {
  const service: SettingsService = React.useMemo(() => new SettingsService(getSP(props.context)), [props.context]);

  const [checking, setChecking] = React.useState<boolean>(true);
  const [canManage, setCanManage] = React.useState<boolean>(false);

  const [settings, setSettings] = React.useState<IIntranetSettings>(DEFAULT_SETTINGS);
  const [nav, setNav] = React.useState<INavLink[]>([]);
  const [loaded, setLoaded] = React.useState<boolean>(false);

  // ---- Setup tab (provisioning) ----
  const [running, setRunning] = React.useState<boolean>(false);
  const [results, setResults] = React.useState<IProvisionResult[]>([]);
  const [done, setDone] = React.useState<boolean>(false);
  const [seedSampleData, setSeedSampleData] = React.useState<boolean>(props.seedSampleData === true);
  // Destructive: wipes the News/Events/Benefits lists before re-seeding. Never
  // defaults on; only meaningful when seedSampleData is also on.
  const [resetSampleData, setResetSampleData] = React.useState<boolean>(false);

  const [settingsSave, setSettingsSave] = React.useState<SaveState>('idle');
  const [navSave, setNavSave] = React.useState<SaveState>('idle');

  React.useEffect(() => {
    let active: boolean = true;
    (async (): Promise<void> => {
      const allowed: boolean = await service.canManage();
      if (!active) {
        return;
      }
      setCanManage(allowed);
      setChecking(false);
      if (allowed) {
        const [s, n] = await Promise.all([service.getSettings(false), service.getNavLinks(false)]);
        if (active) {
          setSettings(s);
          setNav(n);
          setLoaded(true);
        }
      }
    })().catch(() => {
      if (active) {
        setChecking(false);
      }
    });
    return () => {
      active = false;
    };
  }, [service]);

  const run = React.useCallback(async (): Promise<void> => {
    // Reset permanently deletes existing content — confirm before proceeding.
    const willReset: boolean = seedSampleData && resetSampleData;
    if (willReset) {
      const ok: boolean = window.confirm(
        `Reset will permanently DELETE all existing items in the "${props.newsList}", "${props.eventsList}" and ` +
          `"${props.benefitsList}" lists on this site, then re-add the sample data. This cannot be undone.\n\nContinue?`
      );
      if (!ok) {
        return;
      }
    }
    setRunning(true);
    setDone(false);
    setResults([]);
    const provisioning: ProvisioningService = new ProvisioningService(props.context);
    try {
      const res: IProvisionResult[] = await provisioning.runAll({
        eventsList: props.eventsList,
        policiesList: props.policiesList,
        newsList: props.newsList,
        benefitsList: props.benefitsList,
        formsLibrary: props.formsLibrary,
        templatesLibrary: props.templatesLibrary,
        pagesLibrary: props.pagesLibrary,
        registerHeaderFooter: props.registerHeaderFooter,
        seedSampleData: seedSampleData,
        resetSampleData: willReset,
        createViewAllPages: props.createViewAllPages
      });
      setResults(res);
    } catch (error) {
      setResults([{ name: 'Setup', status: 'error', message: error instanceof Error ? error.message : String(error) }]);
    } finally {
      setRunning(false);
      setDone(true);
    }
  }, [props, seedSampleData, resetSampleData]);

  const setField = <K extends keyof IIntranetSettings>(key: K, value: IIntranetSettings[K]): void => {
    setSettings((prev) => ({ ...prev, [key]: value }));
    setSettingsSave('idle');
  };

  const saveSettings = React.useCallback(async (): Promise<void> => {
    setSettingsSave('saving');
    try {
      await service.saveSettings(settings);
      setSettingsSave('saved');
    } catch {
      setSettingsSave('error');
    }
  }, [service, settings]);

  const saveNav = React.useCallback(async (): Promise<void> => {
    setNavSave('saving');
    try {
      await service.saveNavLinks(nav.filter((n) => n.label && n.url));
      setNavSave('saved');
    } catch {
      setNavSave('error');
    }
  }, [service, nav]);

  const updateNav = (index: number, patch: Partial<INavLink>): void => {
    setNav((prev) => prev.map((item, i) => (i === index ? { ...item, ...patch } : item)));
    setNavSave('idle');
  };
  const moveNav = (index: number, delta: number): void => {
    setNav((prev) => {
      const next: INavLink[] = [...prev];
      const target: number = index + delta;
      if (target < 0 || target >= next.length) {
        return prev;
      }
      const tmp: INavLink = next[index];
      next[index] = next[target];
      next[target] = tmp;
      return next;
    });
    setNavSave('idle');
  };
  const removeNav = (index: number): void => {
    setNav((prev) => prev.filter((_, i) => i !== index));
    setNavSave('idle');
  };
  const addNav = (): void => {
    setNav((prev) => [...prev, { label: 'New link', url: '/', newTab: false }]);
    setNavSave('idle');
  };

  // ---- Quick Links tile editor (stored in the settings blob) ----
  type LinkSet = 'quickLinks';
  const updateLink = (which: LinkSet, index: number, patch: Partial<IQuickLinkSetting>): void => {
    setSettings((prev) => ({
      ...prev,
      [which]: prev[which].map((it, i) => (i === index ? { ...it, ...patch } : it))
    }));
    setSettingsSave('idle');
  };
  const moveLink = (which: LinkSet, index: number, delta: number): void => {
    setSettings((prev) => {
      const next: IQuickLinkSetting[] = [...prev[which]];
      const target: number = index + delta;
      if (target < 0 || target >= next.length) {
        return prev;
      }
      const tmp: IQuickLinkSetting = next[index];
      next[index] = next[target];
      next[target] = tmp;
      return { ...prev, [which]: next };
    });
    setSettingsSave('idle');
  };
  const removeLink = (which: LinkSet, index: number): void => {
    setSettings((prev) => ({ ...prev, [which]: prev[which].filter((_, i) => i !== index) }));
    setSettingsSave('idle');
  };
  const addLink = (which: LinkSet): void => {
    setSettings((prev) => ({
      ...prev,
      [which]: [...prev[which], { label: 'New tile', abbr: '', icon: '', url: '#', bg: '#f5f3f0', color: '#262626', newTab: false }]
    }));
    setSettingsSave('idle');
  };

  // ---- Employee Directory filters (stored in the settings blob) ----
  const updateDomain = (index: number, patch: Partial<IDirectoryDomain>): void => {
    setSettings((prev) => ({
      ...prev,
      directoryDomains: prev.directoryDomains.map((it, i) => (i === index ? { ...it, ...patch } : it))
    }));
    setSettingsSave('idle');
  };
  const moveDomain = (index: number, delta: number): void => {
    setSettings((prev) => {
      const next: IDirectoryDomain[] = [...prev.directoryDomains];
      const target: number = index + delta;
      if (target < 0 || target >= next.length) {
        return prev;
      }
      const tmp: IDirectoryDomain = next[index];
      next[index] = next[target];
      next[target] = tmp;
      return { ...prev, directoryDomains: next };
    });
    setSettingsSave('idle');
  };
  const removeDomain = (index: number): void => {
    setSettings((prev) => ({ ...prev, directoryDomains: prev.directoryDomains.filter((_, i) => i !== index) }));
    setSettingsSave('idle');
  };
  const addDomain = (): void => {
    setSettings((prev) => ({ ...prev, directoryDomains: [...prev.directoryDomains, { label: 'New company', domain: '' }] }));
    setSettingsSave('idle');
  };
  const setExclude = (value: string): void => {
    const list: string[] = value.split('\n').map((s) => s.trim()).filter((s) => s.length > 0);
    setSettings((prev) => ({ ...prev, directoryExclude: list }));
    setSettingsSave('idle');
  };

  // ---- Departments directory editor ----
  const updateDept = (index: number, patch: Partial<IDepartmentEntry>): void => {
    setSettings((prev) => ({
      ...prev,
      departments: prev.departments.map((it, i) => (i === index ? { ...it, ...patch } : it))
    }));
    setSettingsSave('idle');
  };
  const moveDept = (index: number, delta: number): void => {
    setSettings((prev) => {
      const next: IDepartmentEntry[] = [...prev.departments];
      const target: number = index + delta;
      if (target < 0 || target >= next.length) {
        return prev;
      }
      const tmp: IDepartmentEntry = next[index];
      next[index] = next[target];
      next[target] = tmp;
      return { ...prev, departments: next };
    });
    setSettingsSave('idle');
  };
  const removeDept = (index: number): void => {
    setSettings((prev) => ({ ...prev, departments: prev.departments.filter((_, i) => i !== index) }));
    setSettingsSave('idle');
  };
  const addDept = (): void => {
    setSettings((prev) => ({
      ...prev,
      departments: [...prev.departments, { label: 'New department', description: '', url: '/sites/', icon: '', accent: '#f5f3f0' }]
    }));
    setSettingsSave('idle');
  };

  const renderLinkEditor = (which: LinkSet): React.ReactNode => {
    const links: IQuickLinkSetting[] = settings[which];
    return (
      <>
        {links.map((item, index) => (
          <div className={styles.qlRow} key={`${which}-${index}`}>
            <TextField ariaLabel="Label" placeholder="Label" value={item.label} onChange={(_, v) => updateLink(which, index, { label: v || '' })} styles={{ root: { width: 130 } }} />
            <TextField ariaLabel="Label (Arabic)" placeholder="Label (Arabic)" value={item.labelAR || ''} onChange={(_, v) => updateLink(which, index, { labelAR: v || '' })} styles={{ root: { width: 130 } }} dir="auto" />
            <TextField ariaLabel="Fluent icon name" placeholder="Icon (e.g. Mail)" value={item.icon || ''} onChange={(_, v) => updateLink(which, index, { icon: v || '' })} styles={{ root: { width: 130 } }} />
            <TextField ariaLabel="Short text" placeholder="Abbr" value={item.abbr || ''} onChange={(_, v) => updateLink(which, index, { abbr: v || '' })} styles={{ root: { width: 70 } }} />
            <TextField ariaLabel="URL" placeholder="URL" value={item.url} onChange={(_, v) => updateLink(which, index, { url: v || '' })} styles={{ root: { width: 200 } }} />
            <TextField ariaLabel="Background colour" placeholder="bg" value={item.bg} onChange={(_, v) => updateLink(which, index, { bg: v || '' })} styles={{ root: { width: 90 } }} />
            <TextField ariaLabel="Glyph colour" placeholder="colour" value={item.color} onChange={(_, v) => updateLink(which, index, { color: v || '' })} styles={{ root: { width: 90 } }} />
            <Toggle ariaLabel="Open in new tab" checked={item.newTab === true} onChange={(_, c) => updateLink(which, index, { newTab: c === true })} onText="New tab" offText="Same tab" />
            <IconButton iconProps={{ iconName: 'Up' }} title="Move up" ariaLabel="Move up" onClick={() => moveLink(which, index, -1)} />
            <IconButton iconProps={{ iconName: 'Down' }} title="Move down" ariaLabel="Move down" onClick={() => moveLink(which, index, 1)} />
            <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Remove" onClick={() => removeLink(which, index)} />
          </div>
        ))}
        <DefaultButton iconProps={{ iconName: 'Add' }} text="Add tile" onClick={() => addLink(which)} />
      </>
    );
  };

  const dotClass = (status: string): string => {
    if (status === 'ok') {
      return `${styles.dot} ${styles.ok}`;
    }
    if (status === 'error') {
      return `${styles.dot} ${styles.error}`;
    }
    return `${styles.dot} ${styles.skipped}`;
  };

  const saveBar = (state: SaveState): React.ReactNode => {
    if (state === 'saved') {
      return <MessageBar messageBarType={MessageBarType.success}>Saved. Changes appear within ~2 minutes (or on next reload).</MessageBar>;
    }
    if (state === 'error') {
      return <MessageBar messageBarType={MessageBarType.error}>Could not save. Run Setup first so the settings lists exist.</MessageBar>;
    }
    return null;
  };

  if (checking) {
    return (
      <section className={styles.admin}>
        <div className={styles.panel}>
          <Spinner size={SpinnerSize.large} label="Checking permissions…" />
        </div>
      </section>
    );
  }

  if (!canManage) {
    return (
      <section className={styles.admin}>
        <div className={styles.panel}>
          <div className={styles.eyebrow}>{getCachedSettings().clientName} · Administration</div>
          <h2 className={styles.title}>Administrators only</h2>
          <MessageBar messageBarType={MessageBarType.warning}>
            The Intranet Setup &amp; Settings screen is available to Site Owners only. Ask a site owner if you need a change here.
          </MessageBar>
        </div>
      </section>
    );
  }

  return (
    <section className={styles.admin}>
      <div className={styles.panel}>
        <div className={styles.eyebrow}>{getCachedSettings().clientName} · Administration</div>
        <h2 className={styles.title}>Intranet Setup &amp; Settings</h2>

        <Pivot>
          <PivotItem headerText="Setup">
            <p className={styles.intro}>
              Provisions the lists, libraries and columns the home-page web parts use (Events, Policies, Templates,
              News category), the central <strong>Settings</strong> and <strong>Navigation</strong> lists
              {props.registerHeaderFooter ? ', and registers the global header &amp; footer' : ''}. The action is
              idempotent — safe to re-run.
            </p>
            <Toggle
              label="Seed sample data (adds example News, Events, Policies and HR Benefits when each list is empty)"
              checked={seedSampleData}
              onChange={(_, c) => {
                setSeedSampleData(c === true);
                if (c !== true) {
                  setResetSampleData(false); // reset only applies alongside seeding
                }
              }}
              onText="On"
              offText="Off"
            />
            <Toggle
              label="Reset before seeding — permanently delete existing News, Events and HR Benefits items first, then re-add the sample data"
              checked={resetSampleData}
              disabled={!seedSampleData}
              onChange={(_, c) => setResetSampleData(c === true)}
              onText="On"
              offText="Off"
            />
            {seedSampleData && resetSampleData ? (
              <div className={styles.warn} role="alert">
                ⚠ Reset is destructive: all existing items in the News, Events and HR Benefits lists will be deleted
                before the sample data is re-added. You will be asked to confirm.
              </div>
            ) : null}
            <button className={styles.btn} onClick={run} disabled={running} type="button">
              {running ? (seedSampleData && resetSampleData ? 'Resetting…' : 'Running setup…') : seedSampleData && resetSampleData ? 'Reset & run setup' : 'Run setup'}
            </button>
            {results.length > 0 ? (
              <div className={styles.results}>
                {results.map((r, i) => (
                  <div className={styles.row} key={`${r.name}-${i}`}>
                    <span className={dotClass(r.status)} aria-hidden="true" />
                    <div>
                      <div className={styles.rowName}>{r.name}</div>
                      <div className={styles.rowMsg}>{r.message}</div>
                    </div>
                  </div>
                ))}
              </div>
            ) : null}
            {done ? (
              <div className={styles.hint}>
                Setup complete. Reload to see header/footer changes, then use the other tabs to configure the intranet.
              </div>
            ) : null}
          </PivotItem>

          <PivotItem headerText="Navigation">
            <p className={styles.intro}>
              Top-navigation links shown in the global header. Reorder, add or remove links and Save — the header reads
              this list live.
            </p>
            {nav.map((item, index) => (
              <div className={styles.navRow} key={`nav-${index}`}>
                <TextField
                  className={styles.navLabel}
                  ariaLabel="Label"
                  value={item.label}
                  onChange={(_, v) => updateNav(index, { label: v || '' })}
                />
                <TextField
                  className={styles.navLabel}
                  ariaLabel="Label (Arabic)"
                  placeholder="Label (Arabic)"
                  value={item.labelAR || ''}
                  onChange={(_, v) => updateNav(index, { labelAR: v || '' })}
                  dir="auto"
                />
                <TextField
                  className={styles.navUrl}
                  ariaLabel="URL"
                  value={item.url}
                  onChange={(_, v) => updateNav(index, { url: v || '' })}
                />
                <Toggle
                  ariaLabel="Open in new tab"
                  checked={item.newTab}
                  onChange={(_, c) => updateNav(index, { newTab: c === true })}
                  onText="New tab"
                  offText="Same tab"
                />
                <IconButton iconProps={{ iconName: 'Up' }} title="Move up" ariaLabel="Move up" onClick={() => moveNav(index, -1)} />
                <IconButton iconProps={{ iconName: 'Down' }} title="Move down" ariaLabel="Move down" onClick={() => moveNav(index, 1)} />
                <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Remove" onClick={() => removeNav(index)} />
              </div>
            ))}
            <div className={styles.actions}>
              <DefaultButton iconProps={{ iconName: 'Add' }} text="Add link" onClick={addNav} />
              <PrimaryButton text={navSave === 'saving' ? 'Saving…' : 'Save navigation'} disabled={navSave === 'saving'} onClick={saveNav} />
            </div>
            {saveBar(navSave)}
          </PivotItem>

          <PivotItem headerText="General">
            <p className={styles.intro}>Header, footer and embedded-portal options applied across the intranet.</p>
            <h3 className={styles.subTitle}>Language</h3>
            <Dropdown
              label="Default language"
              selectedKey={settings.defaultLanguage}
              options={[
                { key: 'en', text: 'English' },
                { key: 'ar', text: 'Arabic' }
              ]}
              onChange={(_, option?: IDropdownOption) => setField('defaultLanguage', (option ? option.key : 'en') as IIntranetSettings['defaultLanguage'])}
            />
            <h3 className={styles.subTitle}>Branding</h3>
            <p className={styles.intro}>Client name and logo — applied across the portal and every department site (department sites inherit these from here).</p>
            <TextField label="Client name" value={settings.clientName} onChange={(_, v) => setField('clientName', v || '')} />
            <TextField label="Client name (Arabic)" value={settings.clientNameAR || ''} onChange={(_, v) => setField('clientNameAR', v || '')} dir="auto" />
            <TextField
              label="Logo URL"
              placeholder="e.g. /SiteAssets/logos/client-logo.png — blank uses the built-in logo"
              value={settings.logoUrl}
              onChange={(_, v) => setField('logoUrl', v || '')}
            />
            <TextField
              label="Header sub-label (small text beside the logo — blank to hide)"
              value={settings.logoSubLabel}
              onChange={(_, v) => setField('logoSubLabel', v || '')}
            />
            <TextField
              label="Accent colour (hex)"
              placeholder="#b88a4a"
              value={settings.accentColor}
              onChange={(_, v) => setField('accentColor', v || '')}
            />
            <TextField
              label="Accent colour — dark/secondary (hex)"
              placeholder="blank = auto-derived (a darker shade of the accent)"
              value={settings.accentColorDark}
              onChange={(_, v) => setField('accentColorDark', v || '')}
            />
            <TextField
              label="Heading font"
              placeholder="e.g. Poppins — blank keeps the built-in font"
              value={settings.fontHead}
              onChange={(_, v) => setField('fontHead', v || '')}
            />
            <TextField
              label="Body font"
              placeholder="e.g. Poppins — blank keeps the built-in font"
              value={settings.fontBody}
              onChange={(_, v) => setField('fontBody', v || '')}
            />
            <TextField
              label="Font stylesheet URL (optional — e.g. a Google Fonts link to load a custom font)"
              placeholder="https://fonts.googleapis.com/css2?family=Poppins:wght@300;500;700&display=swap"
              value={settings.fontStylesheetUrl}
              onChange={(_, v) => setField('fontStylesheetUrl', v || '')}
            />
            <h3 className={styles.subTitle}>Footer &amp; chrome</h3>
            <Toggle
              label="Show footer"
              checked={settings.footerVisible}
              onChange={(_, c) => setField('footerVisible', c === true)}
              onText="Visible"
              offText="Hidden"
            />
            <TextField label="Footer copyright" placeholder="Blank = © {year} {client name}. All rights reserved." value={settings.footerCopyright} onChange={(_, v) => setField('footerCopyright', v || '')} />
            <TextField label="Footer copyright (Arabic)" value={settings.footerCopyrightAR || ''} onChange={(_, v) => setField('footerCopyrightAR', v || '')} dir="auto" />
            <TextField label="Footer tagline" value={settings.footerTagline} onChange={(_, v) => setField('footerTagline', v || '')} />
            <TextField label="Footer tagline (Arabic)" value={settings.footerTaglineAR || ''} onChange={(_, v) => setField('footerTaglineAR', v || '')} dir="auto" />
            <TextField label="Search placeholder" value={settings.searchPlaceholder} onChange={(_, v) => setField('searchPlaceholder', v || '')} />
            <TextField label="Search placeholder (Arabic)" value={settings.searchPlaceholderAR || ''} onChange={(_, v) => setField('searchPlaceholderAR', v || '')} dir="auto" />
            <Toggle label="Hide native SharePoint site header" checked={settings.hideSiteHeader} onChange={(_, c) => setField('hideSiteHeader', c === true)} />
            <Toggle label="Hide native site navigation only" checked={settings.hideSiteNav} onChange={(_, c) => setField('hideSiteNav', c === true)} />
            <Toggle label="Hide SharePoint command bar (view mode)" checked={settings.hideCommandBar} onChange={(_, c) => setField('hideCommandBar', c === true)} />
            <Toggle label="Hide URL in embedded-portal frames" checked={settings.embedHideUrl} onChange={(_, c) => setField('embedHideUrl', c === true)} />
            <TextField
              label="SharePoint top bar colour (hex) — blank = tenant default"
              placeholder="#262626"
              value={settings.suiteBarColor}
              onChange={(_, v) => setField('suiteBarColor', v || '')}
            />
            <TextField
              label="Top bar text/icon colour (hex)"
              placeholder="#ffffff"
              value={settings.suiteBarTextColor}
              onChange={(_, v) => setField('suiteBarTextColor', v || '')}
            />
            <div className={styles.actions}>
              <PrimaryButton text={settingsSave === 'saving' ? 'Saving…' : 'Save settings'} disabled={settingsSave === 'saving' || !loaded} onClick={saveSettings} />
            </div>
            {saveBar(settingsSave)}
          </PivotItem>

          <PivotItem headerText="Web parts">
            <p className={styles.intro}>Defaults applied to the home-page web parts (section headers, item counts, accent colour).</p>
            <Toggle
              label="Show 'View All' links on home-page sections"
              checked={settings.showViewAll}
              onChange={(_, c) => setField('showViewAll', c === true)}
              onText="Shown"
              offText="Hidden"
            />
            <Toggle
              label="Open widget links in a new tab"
              checked={settings.openLinksInNewTab}
              onChange={(_, c) => setField('openLinksInNewTab', c === true)}
              onText="New tab"
              offText="Same tab"
            />
            <TextField label="Detail page URL (hosts the Detail View web part)" value={settings.detailPageUrl} onChange={(_, v) => setField('detailPageUrl', v || '')} />
            <TextField
              label="News — items shown"
              type="number"
              value={String(settings.newsMaxItems)}
              onChange={(_, v) => setField('newsMaxItems', parseInt(v || '0', 10) || DEFAULT_SETTINGS.newsMaxItems)}
            />
            <TextField
              label="Events — items shown"
              type="number"
              value={String(settings.eventsMaxItems)}
              onChange={(_, v) => setField('eventsMaxItems', parseInt(v || '0', 10) || DEFAULT_SETTINGS.eventsMaxItems)}
            />
            <TextField
              label="Directory — page size"
              type="number"
              value={String(settings.directoryPageSize)}
              onChange={(_, v) => setField('directoryPageSize', parseInt(v || '0', 10) || DEFAULT_SETTINGS.directoryPageSize)}
            />
            <div className={styles.actions}>
              <PrimaryButton text={settingsSave === 'saving' ? 'Saving…' : 'Save settings'} disabled={settingsSave === 'saving' || !loaded} onClick={saveSettings} />
            </div>
            {saveBar(settingsSave)}
          </PivotItem>

          <PivotItem headerText="Quick Links">
            <p className={styles.intro}>
              Tiles shown by the <strong>Quick Links</strong> web part and the Productivity Strip. Each tile uses a Fluent
              UI <em>icon</em> name (e.g. <code>Mail</code>, <code>Calendar</code>); if blank, the short <em>abbr</em> text
              is shown. A web part with its own JSON in the property pane overrides these central tiles.
            </p>
            {renderLinkEditor('quickLinks')}
            <div className={styles.actions}>
              <PrimaryButton text={settingsSave === 'saving' ? 'Saving…' : 'Save links'} disabled={settingsSave === 'saving' || !loaded} onClick={saveSettings} />
            </div>
            {saveBar(settingsSave)}
          </PivotItem>

          <PivotItem headerText="Directory">
            <p className={styles.intro}>
              Controls the <strong>Employee Directory</strong>. Search queries the full tenant directory (Microsoft
              Graph), not just the loaded page. Add the company email domains to show — each becomes a filter chip and
              limits which people appear. Leave the list empty to show everyone in the tenant.
            </p>
            <h3 className={styles.subTitle}>Companies (email domains)</h3>
            {settings.directoryDomains.map((item, index) => (
              <div className={styles.qlRow} key={`dom-${index}`}>
                <TextField ariaLabel="Company label" placeholder="Company (chip label)" value={item.label} onChange={(_, v) => updateDomain(index, { label: v || '' })} styles={{ root: { width: 200 } }} />
                <TextField ariaLabel="Email domain" placeholder="email domain e.g. example.com" value={item.domain} onChange={(_, v) => updateDomain(index, { domain: v || '' })} styles={{ root: { width: 260 } }} />
                <IconButton iconProps={{ iconName: 'Up' }} title="Move up" ariaLabel="Move up" onClick={() => moveDomain(index, -1)} />
                <IconButton iconProps={{ iconName: 'Down' }} title="Move down" ariaLabel="Move down" onClick={() => moveDomain(index, 1)} />
                <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Remove" onClick={() => removeDomain(index)} />
              </div>
            ))}
            <DefaultButton iconProps={{ iconName: 'Add' }} text="Add company" onClick={addDomain} />
            <h3 className={styles.subTitle}>Exclude (one term per line)</h3>
            <p className={styles.intro}>
              People whose name, email, <strong>position</strong> or department contains any of these terms are hidden —
              e.g. <code>Service Account</code>, <code>svc-</code>, <code>no-reply</code>. Use this to drop accounts whose
              job title / position is set to “Service Account”.
            </p>
            <TextField
              ariaLabel="Exclude terms"
              multiline
              rows={5}
              value={(settings.directoryExclude || []).join('\n')}
              onChange={(_, v) => setExclude(v || '')}
            />
            <div className={styles.actions}>
              <PrimaryButton text={settingsSave === 'saving' ? 'Saving…' : 'Save directory'} disabled={settingsSave === 'saving' || !loaded} onClick={saveSettings} />
            </div>
            {saveBar(settingsSave)}
          </PivotItem>

          <PivotItem headerText="Departments">
            <p className={styles.intro}>
              Tiles shown by the <strong>Departments Directory</strong> web part — one per department site. Each uses a
              Fluent UI <em>icon</em> name (e.g. <code>People</code>, <code>Money</code>); if blank, the department
              initials are shown. The <em>URL</em> links to that department&apos;s site (e.g. <code>/sites/HR</code>).
            </p>
            {settings.departments.map((item, index) => (
              <div className={styles.qlRow} key={`dept-${index}`}>
                <TextField ariaLabel="Name" placeholder="Department name" value={item.label} onChange={(_, v) => updateDept(index, { label: v || '' })} styles={{ root: { width: 190 } }} />
                <TextField ariaLabel="Name (Arabic)" placeholder="Department name (Arabic)" value={item.labelAR || ''} onChange={(_, v) => updateDept(index, { labelAR: v || '' })} styles={{ root: { width: 190 } }} dir="auto" />
                <TextField ariaLabel="Description" placeholder="Short description" value={item.description} onChange={(_, v) => updateDept(index, { description: v || '' })} styles={{ root: { width: 300 } }} />
                <TextField ariaLabel="Description (Arabic)" placeholder="Short description (Arabic)" value={item.descriptionAR || ''} onChange={(_, v) => updateDept(index, { descriptionAR: v || '' })} styles={{ root: { width: 300 } }} dir="auto" />
                <TextField ariaLabel="Site URL" placeholder="/sites/HR" value={item.url} onChange={(_, v) => updateDept(index, { url: v || '' })} styles={{ root: { width: 180 } }} />
                <TextField ariaLabel="Fluent icon name" placeholder="Icon" value={item.icon || ''} onChange={(_, v) => updateDept(index, { icon: v || '' })} styles={{ root: { width: 120 } }} />
                <TextField ariaLabel="Accent" placeholder="accent" value={item.accent || ''} onChange={(_, v) => updateDept(index, { accent: v || '' })} styles={{ root: { width: 90 } }} />
                <IconButton iconProps={{ iconName: 'Up' }} title="Move up" ariaLabel="Move up" onClick={() => moveDept(index, -1)} />
                <IconButton iconProps={{ iconName: 'Down' }} title="Move down" ariaLabel="Move down" onClick={() => moveDept(index, 1)} />
                <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Remove" onClick={() => removeDept(index)} />
              </div>
            ))}
            <DefaultButton iconProps={{ iconName: 'Add' }} text="Add department" onClick={addDept} />
            <div className={styles.actions}>
              <PrimaryButton text={settingsSave === 'saving' ? 'Saving…' : 'Save departments'} disabled={settingsSave === 'saving' || !loaded} onClick={saveSettings} />
            </div>
            {saveBar(settingsSave)}
          </PivotItem>
        </Pivot>
      </div>
    </section>
  );
};

export default IntranetAdmin;
