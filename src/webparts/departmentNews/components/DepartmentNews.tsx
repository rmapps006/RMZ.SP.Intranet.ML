import * as React from 'react';
import styles from './DepartmentNews.module.scss';
import { IDepartmentNewsProps } from './IDepartmentNewsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { INewsItem } from '../../../common/models';
import { formatLongDate, linkTarget } from '../../../common/util/format';
import { getDepartmentNews } from '../services/DepartmentNewsService';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';
import { useNavKey } from '../../../common/services/useNavKey';

const DepartmentNews: React.FunctionComponent<IDepartmentNewsProps> = (props) => {
  const [items, setItems] = React.useState<INewsItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);

  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const newsList: string = props.newsList || ds.newsList;
  const allNewsUrl: string = props.allNewsUrl || ds.allNewsUrl;
  const detailUrl: string = ds.detailPageUrl;
  const newTab: boolean = ds.openLinksInNewTab;
  const navKey: string = useNavKey();

  // Detail page link when configured (else the item's own LinkUrl, else no link).
  const hrefFor = (item: INewsItem): string | undefined => {
    if (detailUrl && item.id !== undefined) {
      return `${detailUrl}?type=news&recId=${item.id}`;
    }
    return item.url || undefined;
  };

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    getDepartmentNews(props.context, newsList)
      .then((result: INewsItem[]) => {
        if (active) {
          setItems(result);
        }
      })
      .catch(() => {
        /* service already returns fallback */
      })
      .finally(() => {
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [props.context, newsList, navKey]);

  return (
    <section className={styles.news}>
      <SectionHeader title="Department News" linkText="All News" linkUrl={allNewsUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      {loading ? (
        <div className={styles.empty}>Loading…</div>
      ) : items.length === 0 ? (
        <div className={styles.empty}>No news to show.</div>
      ) : (
        <div className={styles.newsG}>
          {items.map((item: INewsItem, index: number) => {
            const inner: JSX.Element = (
              <>
                <div
                  className={styles.ncImg}
                  style={
                    item.imageUrl
                      ? { backgroundImage: `url("${item.imageUrl}")`, backgroundSize: 'cover', backgroundPosition: 'center' }
                      : { background: item.imageGradient }
                  }
                >
                  {item.imageCaption && !item.imageUrl ? <span className={styles.ncCap}>{item.imageCaption}</span> : null}
                </div>
                <div className={styles.ncBd}>
                  <div className={styles.ncTag}>{item.category}</div>
                  <h3 className={styles.ncH}>{item.title}</h3>
                  <div className={styles.ncFt}>
                    <span className={styles.ncDt}>{formatLongDate(item.date)}</span>
                    {item.author ? <span className={styles.ncAu}>{item.author}</span> : null}
                  </div>
                </div>
              </>
            );
            const key: string = `${item.title}-${index}`;
            const href: string | undefined = hrefFor(item);
            return href ? (
              <a key={key} className={`${styles.nc} ${styles.ncActive}`} href={href} {...linkTarget(newTab)} style={{ textDecoration: 'none', color: 'inherit' }}>
                {inner}
              </a>
            ) : (
              <div key={key} className={styles.nc}>
                {inner}
              </div>
            );
          })}
        </div>
      )}
    </section>
  );
};

export default DepartmentNews;
