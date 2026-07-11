import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './ProductivityStrip.module.scss';
import { IProductivityStripProps } from './IProductivityStripProps';
import { getTodaySchedule, IScheduleItem } from '../services/ProductivityService';
import { IQuickLinkSetting } from '../../../common/services/SettingsService';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';

/** A single quick-access tile definition. */
interface IQuickLink {
  label: string;
  abbr: string;
  img?: string; // image/logo URL or data URI; takes priority over icon and abbr
  icon?: string; // Fluent UI (Office) icon name; takes priority over abbr
  url: string;
  bg: string;
  color: string;
  newTab?: boolean;
}

const ORA_OCEAN: string = '#a4e0e6';
const ORA_SAND: string = '#e3cba8';

/** Normalises a raw tile (from JSON or central settings) into a renderable tile. */
function normalizeLink(q: IQuickLinkSetting): IQuickLink {
  return {
    label: q.label || '',
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
 * per-instance override; otherwise the centrally-managed set (edited in the
 * Admin screen) is used.
 */
function resolveLinks(json: string, central: IQuickLinkSetting[]): IQuickLink[] {
  if (json && json.trim().length > 0) {
    try {
      const parsed: unknown = JSON.parse(json);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return (parsed as IQuickLinkSetting[]).map(normalizeLink);
      }
    } catch {
      /* fall through to central settings */
    }
  }
  return (central || []).map(normalizeLink);
}

function LinkTiles(props: { links: IQuickLink[]; newTab: boolean }): JSX.Element {
  return (
    <div className={styles.qlG}>
      {props.links.map((link: IQuickLink, idx: number) => (
        <a
          className={styles.qlI}
          href={link.url || '#'}
          key={idx}
          {...linkTarget(link.newTab || props.newTab)}
        >
          <div className={styles.qlIc} style={{ background: link.bg, color: link.color }}>
            {link.img ? (
              <img className={styles.qlImg} src={link.img} alt="" aria-hidden="true" />
            ) : link.icon ? (
              <Icon iconName={link.icon} />
            ) : (
              link.abbr
            )}
          </div>
          <div className={styles.qlLb}>{link.label}</div>
        </a>
      ))}
    </div>
  );
}

const ProductivityStrip: React.FunctionComponent<IProductivityStripProps> = (props) => {
  const [schedule, setSchedule] = React.useState<IScheduleItem[]>([]);
  const settings = useSettings(props.context);
  const navKey: string = useNavKey();

  const maxEvents: number = (() => {
    const parsed: number = parseInt(props.maxEvents, 10);
    return isNaN(parsed) || parsed <= 0 ? 3 : parsed;
  })();

  React.useEffect(() => {
    let active: boolean = true;
    getTodaySchedule(props.context, maxEvents)
      .then((items: IScheduleItem[]) => {
        if (active) {
          setSchedule(items);
        }
      })
      .catch(() => {
        /* service already falls back */
      });
    return () => {
      active = false;
    };
  }, [props.context, maxEvents, navKey]);

  const quickLinks: IQuickLink[] = resolveLinks(props.quickLinksJson, settings.quickLinks);

  // "Today" and the schedule follow the viewer's own local timezone.
  const now: Date = new Date();
  const dayNum: string = now.getDate().toString();
  const weekday: string = now.toLocaleDateString('en-US', { weekday: 'long' });
  const monthYear: string = now.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

  return (
    <section className={styles.prod}>
      <div className={styles.prodG}>
        {/* Today's Schedule (left) */}
        <div className={styles.pc}>
          <div className={styles.pcLbl}>Today&apos;s Schedule</div>
          <div className={styles.calTop}>
            <div className={styles.calNum}>{dayNum}</div>
            <div>
              <div className={styles.calDn}>{weekday}</div>
              <div className={styles.calMo}>{monthYear}</div>
            </div>
          </div>
          {schedule.length === 0 ? (
            <div className={styles.calEmpty}>No events scheduled today.</div>
          ) : (
            schedule.map((item: IScheduleItem, idx: number) => (
              <div className={styles.calItem} key={idx}>
                <div className={styles.calT}>{item.time}</div>
                <div
                  className={styles.calDot}
                  style={{ background: idx % 2 === 0 ? ORA_OCEAN : ORA_SAND }}
                />
                <div>{item.subject}</div>
              </div>
            ))
          )}
        </div>

        {/* Quick Links (right) */}
        <div className={styles.pc}>
          <div className={styles.pcLbl}>Quick Links</div>
          <LinkTiles links={quickLinks} newTab={settings.openLinksInNewTab} />
        </div>
      </div>
    </section>
  );
};

export default ProductivityStrip;
