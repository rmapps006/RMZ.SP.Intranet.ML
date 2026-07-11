import * as React from 'react';
import styles from './ContentManager.module.scss';
import { IContentManagerProps } from './IContentManagerProps';
import { Pivot, PivotItem, MessageBar, MessageBarType, Spinner, SpinnerSize } from '@fluentui/react';
import { NEWS_TYPE, EVENTS_TYPE, BENEFITS_TYPE, IContentTypeDef } from '../services/contentTypes';
import { canContribute, listExists } from '../services/ContentService';
import { getCachedSettings } from '../../../common/services/SettingsService';
import ContentPanel from './ContentPanel';

interface ITab {
  def: IContentTypeDef;
  listTitle: string;
}

const ContentManager: React.FunctionComponent<IContentManagerProps> = (props) => {
  const [checking, setChecking] = React.useState<boolean>(true);
  const [allowed, setAllowed] = React.useState<boolean>(false);
  const [tabs, setTabs] = React.useState<ITab[]>([]);

  // Build the configured tab set (property toggles + list name overrides).
  const configured: ITab[] = React.useMemo(() => {
    const list: ITab[] = [];
    if (props.showNews) {
      list.push({ def: NEWS_TYPE, listTitle: props.newsList || 'News' });
    }
    if (props.showEvents) {
      list.push({ def: EVENTS_TYPE, listTitle: props.eventsList || 'Events' });
    }
    if (props.showBenefits) {
      list.push({ def: BENEFITS_TYPE, listTitle: props.benefitsList || 'HR Benefits' });
    }
    return list;
  }, [props.showNews, props.showEvents, props.showBenefits, props.newsList, props.eventsList, props.benefitsList]);

  React.useEffect(() => {
    let active: boolean = true;
    (async (): Promise<void> => {
      // Keep only tabs whose list actually exists on this site, and gate on the
      // user being able to add items to at least one of them.
      const present: ITab[] = [];
      for (const t of configured) {
        if (await listExists(props.context, t.listTitle)) {
          present.push(t);
        }
      }
      let canAny: boolean = false;
      for (const t of present) {
        if (await canContribute(props.context, t.listTitle)) {
          canAny = true;
          break;
        }
      }
      if (active) {
        setTabs(present);
        setAllowed(canAny);
        setChecking(false);
      }
    })().catch(() => {
      if (active) {
        setChecking(false);
      }
    });
    return () => {
      active = false;
    };
  }, [props.context, configured]);

  if (checking) {
    return (
      <section className={styles.cm}>
        <Spinner size={SpinnerSize.large} label="Loading content manager…" />
      </section>
    );
  }

  if (tabs.length === 0) {
    return (
      <section className={styles.cm}>
        <MessageBar messageBarType={MessageBarType.warning}>
          None of the configured content lists were found on this site. Run Setup (Admin) first, or check the list names in
          this web part’s properties.
        </MessageBar>
      </section>
    );
  }

  if (!allowed) {
    return (
      <section className={styles.cm}>
        <MessageBar messageBarType={MessageBarType.warning}>
          You don’t have permission to add content here. Ask a site owner to grant you contribute access.
        </MessageBar>
      </section>
    );
  }

  return (
    <section className={styles.cm}>
      <div className={styles.head}>
        <div className={styles.eyebrow}>{getCachedSettings().clientName} · Content</div>
        <h2 className={styles.title}>{props.title || 'Content Manager'}</h2>
        <p className={styles.intro}>Add, edit and remove portal content without leaving this page.</p>
      </div>
      <Pivot>
        {tabs.map((t) => (
          <PivotItem headerText={t.def.label} key={t.def.key}>
            <ContentPanel context={props.context} listTitle={t.listTitle} def={t.def} />
          </PivotItem>
        ))}
      </Pivot>
    </section>
  );
};

export default ContentManager;
