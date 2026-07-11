import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './QuickLinks.module.scss';
import { IQuickLinksProps } from './IQuickLinksProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { getCachedSettings, IQuickLinkSetting } from '../../../common/services/SettingsService';
import { getCurrentLanguage, pickLocalized, Language } from '../../../common/services/languageService';
import { linkTarget } from '../../../common/util/format';

/** A single quick-access tile. */
interface IQuickLink {
  label: string;
  abbr: string;
  img?: string; // image/logo URL or data URI — takes priority over icon and abbr
  icon?: string; // Fluent UI (Office) icon name, e.g. "Mail" — takes priority over abbr
  url: string;
  bg: string;
  color: string;
  newTab?: boolean;
}

/** Normalises a raw tile (from JSON or central settings) into a renderable tile. */
function normalizeLink(q: IQuickLinkSetting, language: Language): IQuickLink {
  return {
    label: pickLocalized(q.label || '', q.labelAR, language),
    abbr: q.abbr || (q.label ? q.label.charAt(0).toUpperCase() : '•'),
    img: q.img || undefined,
    icon: q.icon || undefined,
    url: q.url || '#',
    bg: q.bg || '#f5f3f0',
    color: q.color || '#262626',
    newTab: q.newTab === true
  };
}

/**
 * Resolves the tiles to render. A non-empty property-pane JSON acts as a
 * per-instance override; otherwise the centrally-managed Quick Links (edited
 * in the Admin screen) are used.
 */
function resolveLinks(json: string, language: Language): IQuickLink[] {
  if (json && json.trim().length > 0) {
    try {
      const parsed: unknown = JSON.parse(json);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return (parsed as IQuickLinkSetting[]).map((q) => normalizeLink(q, language));
      }
    } catch {
      /* fall through to central settings */
    }
  }
  return (getCachedSettings().quickLinks || []).map((q) => normalizeLink(q, language));
}

const QuickLinks: React.FunctionComponent<IQuickLinksProps> = (props) => {
  const language: Language = getCurrentLanguage();
  const links: IQuickLink[] = resolveLinks(props.linksJson, language);
  // The central setting forces a new tab; a per-tile newTab can still opt in.
  const globalNewTab: boolean = getCachedSettings().openLinksInNewTab;

  return (
    <section className={styles.quickLinks}>
      <SectionHeader
        title={props.title}
        linkText="View All"
        linkUrl={props.viewAllUrl}
        showTitle={props.showTitle}
        showLink={props.showViewAll}
      />
      {links.length === 0 ? (
        <div className={styles.empty}>No quick links configured.</div>
      ) : (
        <div className={styles.grid}>
          {links.map((link, idx) => (
            <a
              className={styles.tile}
              href={link.url || '#'}
              key={`${link.label}-${idx}`}
              aria-label={link.label}
              {...linkTarget(link.newTab || globalNewTab)}
            >
              <div className={styles.icon} style={{ background: link.bg, color: link.color }}>
                {link.img ? (
                  <img className={styles.iconImg} src={link.img} alt="" aria-hidden="true" />
                ) : link.icon ? (
                  <Icon iconName={link.icon} />
                ) : (
                  link.abbr
                )}
              </div>
              <div className={styles.label}>{link.label}</div>
            </a>
          ))}
        </div>
      )}
    </section>
  );
};

export default QuickLinks;
