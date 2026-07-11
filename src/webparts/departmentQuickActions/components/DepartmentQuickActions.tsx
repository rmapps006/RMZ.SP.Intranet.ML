import * as React from 'react';
import styles from './DepartmentQuickActions.module.scss';
import { IDepartmentQuickActionsProps } from './IDepartmentQuickActionsProps';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';
import { getCurrentLanguage, pickLocalized, Language } from '../../../common/services/languageService';

interface IQuickAction {
  label: string;
  url: string;
}

const formsIcon: React.ReactElement = (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth={1.3} strokeLinecap="round">
    <rect x="2.5" y="1.5" width="9" height="11" rx="1" />
    <line x1="4.5" y1="4.5" x2="9.5" y2="4.5" />
    <line x1="4.5" y1="6.5" x2="9.5" y2="6.5" />
    <line x1="4.5" y1="8.5" x2="7.5" y2="8.5" />
  </svg>
);

const documentsIcon: React.ReactElement = (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth={1.3} strokeLinecap="round" strokeLinejoin="round">
    <path d="M4 1.5h4L11 4.5V12a1 1 0 01-1 1H4a1 1 0 01-1-1V2.5a1 1 0 011-1z" />
    <path d="M8 1.5V4.5H11" />
  </svg>
);

const eventsIcon: React.ReactElement = (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth={1.3} strokeLinecap="round">
    <rect x="2" y="3" width="10" height="9" rx="1" />
    <line x1="2" y1="6" x2="12" y2="6" />
    <line x1="5" y1="1.5" x2="5" y2="4.5" />
    <line x1="9" y1="1.5" x2="9" y2="4.5" />
  </svg>
);

const defaultIcon: React.ReactElement = (
  <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth={1.3} strokeLinecap="round">
    <circle cx="7" cy="7" r="5" />
  </svg>
);

/** Picks an icon from the action label so custom actions still get a sensible glyph. */
function iconFor(label: string): React.ReactElement {
  const l: string = label.toLowerCase();
  if (l.indexOf('form') !== -1) {
    return formsIcon;
  }
  if (l.indexOf('event') !== -1 || l.indexOf('calendar') !== -1) {
    return eventsIcon;
  }
  if (l.indexOf('doc') !== -1 || l.indexOf('polic') !== -1 || l.indexOf('librar') !== -1) {
    return documentsIcon;
  }
  return defaultIcon;
}

function parseActions(json: string): IQuickAction[] {
  try {
    const parsed: unknown = JSON.parse(json);
    if (Array.isArray(parsed)) {
      return (parsed as Array<Partial<IQuickAction>>).map((a) => ({
        label: typeof a.label === 'string' ? a.label : '',
        url: typeof a.url === 'string' ? a.url : ''
      }));
    }
  } catch {
    /* invalid JSON — render no actions */
  }
  return [];
}

const DepartmentQuickActions: React.FunctionComponent<IDepartmentQuickActionsProps> = (props) => {
  // Property-pane JSON wins; otherwise use the central Department Admin quick links.
  const ds = useDepartmentSettings(props.context);
  const language: Language = getCurrentLanguage();
  const actions: IQuickAction[] =
    props.actionsJson && props.actionsJson.trim().length > 0
      ? parseActions(props.actionsJson)
      : (ds.quickActions || []).map((a) => ({ label: pickLocalized(a.label || '', a.labelAR, language), url: a.url || '' }));

  return (
    <section className={styles.qaStrip}>
      <div className={styles.qaRow}>
        {actions.map((action, i) => (
          <a className={styles.qaBtn} href={action.url || '#'} key={i}>
            {iconFor(action.label)}
            {action.label}
          </a>
        ))}
      </div>
    </section>
  );
};

export default DepartmentQuickActions;
