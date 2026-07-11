import * as React from 'react';
import styles from './SectionHeader.module.scss';
import { getCachedSettings } from '../services/SettingsService';

export interface ISectionHeaderProps {
  title: string;
  linkText?: string;
  linkUrl?: string;
  /**
   * Master switch for the "View All" link. Defaults to true. Set false (e.g.
   * from the central settings or a web part toggle) to hide it everywhere.
   */
  showLink?: boolean;
  /** Hide the title text (keep the link). Defaults to true (title shown). */
  showTitle?: boolean;
}

/**
 * Resolves a link target to a pathname and compares it with the current page.
 * Used so a web part placed on its own "View All" destination page does not
 * render a redundant link back to the page you are already on.
 */
function isCurrentPage(linkUrl: string | undefined): boolean {
  if (!linkUrl) {
    return false;
  }
  try {
    const target: string = new URL(linkUrl, window.location.origin).pathname.toLowerCase();
    const here: string = window.location.pathname.toLowerCase();
    if (!target || target === '/') {
      return false;
    }
    return here === target;
  } catch {
    return false;
  }
}

export const SectionHeader: React.FunctionComponent<ISectionHeaderProps> = (props) => {
  // Global "View All" default from the Admin screen (read synchronously from
  // cache). The accent colour is themed via the --intranet-accent CSS variable that
  // the header/footer extension injects, so it is not handled here.
  const globalShowViewAll: boolean = getCachedSettings().showViewAll !== false;

  const showTitle: boolean = props.showTitle !== false;
  // The link shows when: explicitly allowed (prop), enabled centrally, there IS
  // link text, and we are not already sitting on the page it points to (#6).
  const showLink: boolean =
    props.showLink !== false &&
    globalShowViewAll &&
    !!props.linkText &&
    !isCurrentPage(props.linkUrl);

  return (
    <div className={styles.head}>
      {showTitle ? <h2 className={styles.title}>{props.title}</h2> : <span />}
      {showLink ? (
        <a className={styles.link} href={props.linkUrl || '#'}>
          {props.linkText}
        </a>
      ) : null}
    </div>
  );
};
