import * as React from 'react';
import styles from './EventsList.module.scss';
import { IEventsListProps } from './IEventsListProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { IEventItem } from '../../../common/models';
import { splitDateChip, linkTarget } from '../../../common/util/format';
import { getEvents } from '../services/EventsService';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';

const EventsList: React.FunctionComponent<IEventsListProps> = (props) => {
  const [items, setItems] = React.useState<IEventItem[]>([]);

  // Reactive settings so central values apply on first paint (not only after
  // the cache-warming extension has run).
  const settings = useSettings(props.context);
  const navKey: string = useNavKey();
  const max: number = (() => {
    const parsed: number = parseInt(props.maxItems, 10);
    // Fall back to the central Admin setting when no per-web-part value is set.
    return isNaN(parsed) || parsed <= 0 ? settings.eventsMaxItems : parsed;
  })();

  const detailUrl: string = settings.detailPageUrl;
  const newTab: boolean = settings.openLinksInNewTab;

  // Detail page link when configured (else the item's own DispForm link, else no link).
  const hrefFor = (item: IEventItem): string | undefined => {
    if (detailUrl && item.id !== undefined) {
      return `${detailUrl}?type=event&recId=${item.id}`;
    }
    return item.url || undefined;
  };

  React.useEffect(() => {
    let active: boolean = true;
    getEvents(props.context, props.eventsList, max)
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
  }, [props.context, props.eventsList, max, navKey]);

  return (
    <section className={styles.events}>
      <SectionHeader title="Company Events" linkText="View Calendar" linkUrl={props.calendarUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

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

export default EventsList;
