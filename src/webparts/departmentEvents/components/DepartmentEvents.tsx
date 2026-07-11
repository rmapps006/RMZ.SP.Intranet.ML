import * as React from 'react';
import styles from './DepartmentEvents.module.scss';
import { IDepartmentEventsProps } from './IDepartmentEventsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { IEventItem } from '../../../common/models';
import { splitDateChip, linkTarget } from '../../../common/util/format';
import { getDepartmentEvents } from '../services/DepartmentEventsService';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';
import { useNavKey } from '../../../common/services/useNavKey';

const DepartmentEvents: React.FunctionComponent<IDepartmentEventsProps> = (props) => {
  const [items, setItems] = React.useState<IEventItem[]>([]);

  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const eventsList: string = props.eventsList || ds.eventsList;
  const calendarUrl: string = props.calendarUrl || ds.calendarUrl;
  const detailUrl: string = ds.detailPageUrl;
  const newTab: boolean = ds.openLinksInNewTab;
  const navKey: string = useNavKey();

  // Detail page link when configured, else no link.
  const hrefFor = (item: IEventItem): string | undefined =>
    detailUrl && item.id !== undefined ? `${detailUrl}?type=event&recId=${item.id}` : item.url || undefined;

  const max: number = (() => {
    const parsed: number = parseInt(props.maxItems, 10);
    if (!isNaN(parsed) && parsed > 0) {
      return parsed;
    }
    return ds.eventsMaxItems > 0 ? ds.eventsMaxItems : 3;
  })();

  React.useEffect(() => {
    let active: boolean = true;
    getDepartmentEvents(props.context, eventsList, max)
      .then((result: IEventItem[]) => {
        if (active) {
          setItems(result);
        }
      })
      .catch(() => {
        /* service already returns fallback */
      });
    return () => {
      active = false;
    };
  }, [props.context, eventsList, max, navKey]);

  return (
    <section className={styles.events}>
      <SectionHeader title="Department Events" linkText="View Calendar" linkUrl={calendarUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.evtsG}>
        {items.map((item: IEventItem, index: number) => {
          const chip: { mo: string; dy: string } = splitDateChip(item.date);
          const inner: JSX.Element = (
            <>
              <div className={styles.ecDate} style={{ background: item.accent }}>
                <div className={styles.ecMo}>{chip.mo}</div>
                <div className={styles.ecDy}>{chip.dy}</div>
              </div>
              <div className={styles.ecBd}>
                <div className={styles.ecTp}>{item.type}</div>
                <div className={styles.ecTt}>{item.title}</div>
                {item.time ? <div className={styles.ecTime}>{item.time}</div> : null}
                {item.location ? <div className={styles.ecLoc}>{item.location}</div> : null}
              </div>
            </>
          );
          const key: string = `${item.title}-${index}`;
          const href: string | undefined = hrefFor(item);
          return href ? (
            <a key={key} className={styles.ec} href={href} {...linkTarget(newTab)} style={{ textDecoration: 'none', color: 'inherit' }}>
              {inner}
            </a>
          ) : (
            <div key={key} className={styles.ec}>
              {inner}
            </div>
          );
        })}
      </div>
    </section>
  );
};

export default DepartmentEvents;
