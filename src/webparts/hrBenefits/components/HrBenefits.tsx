import * as React from 'react';
import styles from './HrBenefits.module.scss';
import { IHrBenefitsProps } from './IHrBenefitsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { IBenefit } from '../../../common/models';
import { getBenefits } from '../services/HrBenefitsService';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';

// The 6 design SVG icons, mapped to the fallback benefits by index.
// SVG attributes are converted to React (strokeWidth, strokeLinecap, etc.).
const ICONS: JSX.Element[] = [
  // Health & Medical
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <path d="M17 6l2 5h5l-4 3 2 5-5-3-5 3 2-5-4-3h5z" />
    <circle cx="17" cy="22" r="6" />
  </svg>,
  // Leave Policy
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <rect x="6" y="8" width="22" height="20" rx="1" />
    <line x1="6" y1="14" x2="28" y2="14" />
    <line x1="12" y1="5" x2="12" y2="11" />
    <line x1="22" y1="5" x2="22" y2="11" />
  </svg>,
  // Financial
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <rect x="6" y="10" width="22" height="16" rx="2" />
    <line x1="6" y1="16" x2="28" y2="16" />
    <line x1="11" y1="21" x2="15" y2="21" />
    <line x1="20" y1="21" x2="24" y2="21" />
  </svg>,
  // Learning & Development
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <rect x="6" y="6" width="22" height="22" rx="2" />
    <line x1="11" y1="13" x2="23" y2="13" />
    <line x1="11" y1="17" x2="20" y2="17" />
    <line x1="11" y1="21" x2="17" y2="21" />
  </svg>,
  // Wellness
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <circle cx="17" cy="14" r="6" />
    <path d="M9 28c0-4.418 3.582-8 8-8s8 3.582 8 8" />
    <line x1="22" y1="22" x2="27" y2="22" />
    <line x1="24.5" y1="19.5" x2="24.5" y2="24.5" />
  </svg>,
  // Education
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <path d="M8 26V12l9-6 9 6v14" />
    <rect x="13" y="18" width="8" height="8" />
    <line x1="17" y1="12" x2="17" y2="15" />
  </svg>
];

// A sensible default icon for configurable (JSON) benefit items.
const DEFAULT_ICON: JSX.Element = (
  <svg viewBox="0 0 34 34" fill="none" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round">
    <circle cx="17" cy="17" r="11" />
    <path d="M17 11v6l4 3" />
  </svg>
);

const FALLBACK: IBenefit[] = [];

function parseBenefits(json: string): { items: IBenefit[]; isFallback: boolean } {
  if (json && json.trim().length > 0) {
    try {
      const parsed: unknown = JSON.parse(json);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const items: IBenefit[] = (parsed as IBenefit[]).map((b) => ({
          category: b.category || '',
          title: b.title || '',
          description: b.description || '',
          linkText: b.linkText || 'View Details',
          url: b.url
        }));
        return { items, isFallback: false };
      }
    } catch {
      /* fall back to design benefits */
    }
  }
  return { items: FALLBACK, isFallback: true };
}

const HrBenefits: React.FunctionComponent<IHrBenefitsProps> = (props) => {
  // Property-pane JSON acts as an override; otherwise read the HR Benefits list.
  const override = parseBenefits(props.benefitsJson);
  const [listItems, setListItems] = React.useState<IBenefit[]>([]);

  const settings = useSettings(props.context);
  const navKey: string = useNavKey();
  const detailUrl: string = settings.detailPageUrl;
  const newTab: boolean = settings.openLinksInNewTab;

  React.useEffect(() => {
    if (!override.isFallback) {
      return; // a JSON override is configured — don't read the list
    }
    let active: boolean = true;
    getBenefits(props.context, props.benefitsList)
      .then((result: IBenefit[]) => {
        if (active) {
          setListItems(result);
        }
      })
      .catch(() => {
        /* keep empty */
      });
    return () => {
      active = false;
    };
  }, [props.context, props.benefitsList, override.isFallback, navKey]);

  const items: IBenefit[] = override.isFallback ? listItems : override.items;
  const useDesignIcons: boolean = override.isFallback;

  const hrefFor = (b: IBenefit): string | undefined => {
    if (detailUrl && b.id !== undefined) {
      return `${detailUrl}?type=benefit&recId=${b.id}`;
    }
    return b.url || undefined;
  };

  return (
    <section className={styles.benefits}>
      <SectionHeader title={props.title} linkText="All Benefits" linkUrl={props.allBenefitsUrl} showTitle={props.showTitle} showLink={props.showViewAll} />
      <div className={styles.grid}>
        {items.map((b, i) => {
          const href: string | undefined = hrefFor(b);
          return (
            <div className={styles.card} key={`${b.title}-${i}`}>
              <span className={styles.icon}>{useDesignIcons ? ICONS[i] || DEFAULT_ICON : DEFAULT_ICON}</span>
              <div className={styles.cat}>{b.category}</div>
              <div className={styles.tt}>{b.title}</div>
              <div className={styles.desc}>{b.description}</div>
              <a className={styles.lnk} href={href || '#'} {...linkTarget(newTab)}>
                {b.linkText || 'View Details'}
              </a>
            </div>
          );
        })}
      </div>
    </section>
  );
};

export default HrBenefits;
