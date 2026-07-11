import * as React from 'react';
import styles from './DepartmentAdmin.module.scss';
import { IDepartmentAdminProps } from './IDepartmentAdminProps';
import { getSP } from '../../../common/services/pnpService';
import {
  DepartmentSettingsService,
  IDepartmentSettings,
  IDepartmentQuickAction,
  DEFAULT_DEPARTMENT_SETTINGS
} from '../../../common/services/DepartmentSettingsService';
import { getCachedSettings } from '../../../common/services/SettingsService';
import {
  Pivot,
  PivotItem,
  TextField,
  PrimaryButton,
  DefaultButton,
  IconButton,
  Toggle,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize
} from '@fluentui/react';

type SaveState = 'idle' | 'saving' | 'saved' | 'error';

const DepartmentAdmin: React.FunctionComponent<IDepartmentAdminProps> = (props) => {
  const service: DepartmentSettingsService = React.useMemo(
    () => new DepartmentSettingsService(getSP(props.context)),
    [props.context]
  );

  const [checking, setChecking] = React.useState<boolean>(true);
  const [canManage, setCanManage] = React.useState<boolean>(false);
  const [loaded, setLoaded] = React.useState<boolean>(false);
  const [settings, setSettings] = React.useState<IDepartmentSettings>(DEFAULT_DEPARTMENT_SETTINGS);
  const [save, setSave] = React.useState<SaveState>('idle');
  const [setupState, setSetupState] = React.useState<SaveState>('idle');
  const [seedSampleData, setSeedSampleData] = React.useState<boolean>(false);

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
        const s: IDepartmentSettings = await service.getSettings(false);
        if (active) {
          setSettings(s);
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

  const setField = <K extends keyof IDepartmentSettings>(key: K, value: IDepartmentSettings[K]): void => {
    setSettings((prev) => ({ ...prev, [key]: value }));
    setSave('idle');
  };

  const saveSettings = React.useCallback(async (): Promise<void> => {
    setSave('saving');
    try {
      await service.saveSettings(settings);
      setSave('saved');
    } catch {
      setSave('error');
    }
  }, [service, settings]);

  const runSetup = React.useCallback(async (): Promise<void> => {
    setSetupState('saving');
    try {
      // Save first so the Main site URL is stored, then create the list and
      // register the shared header/footer on this department site.
      await service.saveSettings(settings);
      await service.runSetup(settings, seedSampleData);
      setSetupState('saved');
      setSave('saved');
    } catch {
      setSetupState('error');
    }
  }, [service, settings, seedSampleData]);

  // ---- Quick Links editor ----
  const updateAction = (index: number, patch: Partial<IDepartmentQuickAction>): void => {
    setSettings((prev) => ({
      ...prev,
      quickActions: prev.quickActions.map((a, i) => (i === index ? { ...a, ...patch } : a))
    }));
    setSave('idle');
  };
  const moveAction = (index: number, delta: number): void => {
    setSettings((prev) => {
      const next: IDepartmentQuickAction[] = [...prev.quickActions];
      const target: number = index + delta;
      if (target < 0 || target >= next.length) {
        return prev;
      }
      const tmp: IDepartmentQuickAction = next[index];
      next[index] = next[target];
      next[target] = tmp;
      return { ...prev, quickActions: next };
    });
    setSave('idle');
  };
  const removeAction = (index: number): void => {
    setSettings((prev) => ({ ...prev, quickActions: prev.quickActions.filter((_, i) => i !== index) }));
    setSave('idle');
  };
  const addAction = (): void => {
    setSettings((prev) => ({ ...prev, quickActions: [...prev.quickActions, { label: 'New', url: '' }] }));
    setSave('idle');
  };

  const saveBar = (state: SaveState, okText: string): React.ReactNode => {
    if (state === 'saved') {
      return <MessageBar messageBarType={MessageBarType.success}>{okText}</MessageBar>;
    }
    if (state === 'error') {
      return <MessageBar messageBarType={MessageBarType.error}>Could not complete. Check that you are a site owner and try again.</MessageBar>;
    }
    return null;
  };

  const saveButton = (
    <div className={styles.actions}>
      <PrimaryButton text={save === 'saving' ? 'Saving…' : 'Save'} disabled={save === 'saving' || !loaded} onClick={saveSettings} />
    </div>
  );

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
          <div className={styles.eyebrow}>{getCachedSettings().clientName} · Department Administration</div>
          <h2 className={styles.title}>Administrators only</h2>
          <MessageBar messageBarType={MessageBarType.warning}>
            The Department Setup &amp; Settings screen is available to Site Owners only.
          </MessageBar>
        </div>
      </section>
    );
  }

  return (
    <section className={styles.admin}>
      <div className={styles.panel}>
        <div className={styles.eyebrow}>{getCachedSettings().clientName} · Department Administration</div>
        <h2 className={styles.title}>Department Setup &amp; Settings</h2>

        <Pivot>
          <PivotItem headerText="Setup">
            <p className={styles.intro}>
              This department site reuses the navigation, header/footer and theme from the main portal. Enter the
              <strong> main portal site URL</strong> below, then <strong>Run setup</strong> — this saves the URL, creates
              the local settings list and registers the shared header/footer on this site. Reload after setup to see the
              portal navigation. Everything else on this page configures the department web parts on this site.
            </p>
            <TextField
              label="Main portal site URL (e.g. https://tenant.sharepoint.com/sites/Intranet)"
              value={settings.mainSiteUrl}
              onChange={(_, v) => setField('mainSiteUrl', v || '')}
            />
            <Toggle
              label="Seed sample data"
              inlineLabel
              checked={seedSampleData}
              onChange={(_, checked) => setSeedSampleData(checked === true)}
              onText="On"
              offText="Off"
            />
            <div className={styles.actions}>
              <DefaultButton text={setupState === 'saving' ? 'Running…' : 'Run setup'} disabled={setupState === 'saving'} onClick={runSetup} />
              <PrimaryButton text={save === 'saving' ? 'Saving…' : 'Save'} disabled={save === 'saving' || !loaded} onClick={saveSettings} />
            </div>
            {saveBar(setupState, 'Setup complete — the Content Manager page (SitePages/Content.aspx) is ready. Reload to see the portal navigation and header.')}
          </PivotItem>

          <PivotItem headerText="Hero">
            <p className={styles.intro}>The department banner shown by the Department Hero web part.</p>
            <TextField label="Eyebrow" value={settings.eyebrow} onChange={(_, v) => setField('eyebrow', v || '')} />
            <TextField label="Eyebrow (Arabic)" value={settings.eyebrowAR} onChange={(_, v) => setField('eyebrowAR', v || '')} dir="auto" />
            <TextField label="Department name" value={settings.departmentName} onChange={(_, v) => setField('departmentName', v || '')} />
            <TextField label="Department name (Arabic)" value={settings.departmentNameAR} onChange={(_, v) => setField('departmentNameAR', v || '')} dir="auto" />
            <TextField label="Description" multiline rows={3} value={settings.description} onChange={(_, v) => setField('description', v || '')} />
            <TextField label="Description (Arabic)" multiline rows={3} value={settings.descriptionAR} onChange={(_, v) => setField('descriptionAR', v || '')} dir="auto" />
            <TextField label="Owner name" value={settings.ownerName} onChange={(_, v) => setField('ownerName', v || '')} />
            <TextField label="Owner role" value={settings.ownerRole} onChange={(_, v) => setField('ownerRole', v || '')} />
            {saveButton}
            {saveBar(save, 'Saved. Reload the department pages to see changes.')}
          </PivotItem>

          <PivotItem headerText="Quick Links">
            <p className={styles.intro}>The outlined buttons shown by the Department Quick Actions web part.</p>
            {settings.quickActions.map((action, index) => (
              <div className={styles.qlRow} key={`qa-${index}`}>
                <TextField ariaLabel="Label" placeholder="Label" value={action.label} onChange={(_, v) => updateAction(index, { label: v || '' })} styles={{ root: { width: 160 } }} />
                <TextField ariaLabel="Label (Arabic)" placeholder="Label (Arabic)" value={action.labelAR} onChange={(_, v) => updateAction(index, { labelAR: v || '' })} styles={{ root: { width: 160 } }} dir="auto" />
                <TextField ariaLabel="URL" placeholder="URL" value={action.url} onChange={(_, v) => updateAction(index, { url: v || '' })} styles={{ root: { width: 300 } }} />
                <IconButton iconProps={{ iconName: 'Up' }} title="Move up" ariaLabel="Move up" onClick={() => moveAction(index, -1)} />
                <IconButton iconProps={{ iconName: 'Down' }} title="Move down" ariaLabel="Move down" onClick={() => moveAction(index, 1)} />
                <IconButton iconProps={{ iconName: 'Delete' }} title="Remove" ariaLabel="Remove" onClick={() => removeAction(index)} />
              </div>
            ))}
            <DefaultButton iconProps={{ iconName: 'Add' }} text="Add button" onClick={addAction} />
            {saveButton}
            {saveBar(save, 'Saved. Reload the department pages to see changes.')}
          </PivotItem>

          <PivotItem headerText="Content">
            <p className={styles.intro}>Data sources for the Department News, Events and Documents web parts on this site.</p>
            <h3 className={styles.subTitle}>News</h3>
            <TextField label="News list" value={settings.newsList} onChange={(_, v) => setField('newsList', v || '')} />
            <TextField label="All-news URL" value={settings.allNewsUrl} onChange={(_, v) => setField('allNewsUrl', v || '')} />
            <h3 className={styles.subTitle}>Events</h3>
            <TextField label="Events list" value={settings.eventsList} onChange={(_, v) => setField('eventsList', v || '')} />
            <TextField label="Calendar URL" value={settings.calendarUrl} onChange={(_, v) => setField('calendarUrl', v || '')} />
            <TextField label="Detail page URL (hosts the Detail View web part)" value={settings.detailPageUrl} onChange={(_, v) => setField('detailPageUrl', v || '')} />
            <TextField
              label="Events shown"
              type="number"
              value={String(settings.eventsMaxItems)}
              onChange={(_, v) => setField('eventsMaxItems', parseInt(v || '0', 10) || DEFAULT_DEPARTMENT_SETTINGS.eventsMaxItems)}
            />
            <h3 className={styles.subTitle}>Forms &amp; Documents</h3>
            <TextField label="Left panel title" value={settings.panel1Title} onChange={(_, v) => setField('panel1Title', v || '')} />
            <TextField label="Left library (Forms & Templates)" value={settings.policiesLibrary} onChange={(_, v) => setField('policiesLibrary', v || '')} />
            <TextField label="Right panel title" value={settings.panel2Title} onChange={(_, v) => setField('panel2Title', v || '')} />
            <TextField label="Right library (Shared Documents)" value={settings.documentsLibrary} onChange={(_, v) => setField('documentsLibrary', v || '')} />
            <TextField label="Document hub URL" value={settings.documentHubUrl} onChange={(_, v) => setField('documentHubUrl', v || '')} />
            <h3 className={styles.subTitle}>Links</h3>
            <Toggle
              label="Open widget links in a new tab"
              checked={settings.openLinksInNewTab}
              onChange={(_, c) => setField('openLinksInNewTab', c === true)}
              onText="New tab"
              offText="Same tab"
            />
            {saveButton}
            {saveBar(save, 'Saved. Reload the department pages to see changes.')}
          </PivotItem>
        </Pivot>
      </div>
    </section>
  );
};

export default DepartmentAdmin;
