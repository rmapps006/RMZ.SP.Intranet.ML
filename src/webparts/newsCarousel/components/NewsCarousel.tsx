import * as React from 'react';
import styles from './NewsCarousel.module.scss';
import { INewsCarouselProps } from './INewsCarouselProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { INewsItem } from '../../../common/models';
import { getNews, NEWS_CATEGORIES } from '../services/NewsService';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';
import { getCurrentLanguage, isRtl, Language } from '../../../common/services/languageService';
import { t, localizeChoice } from '../../../common/services/uiStrings';

const NewsCarousel: React.FunctionComponent<INewsCarouselProps> = (props) => {
  const [items, setItems] = React.useState<INewsItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [activeCategory, setActiveCategory] = React.useState<string>('All');
  const language: Language = getCurrentLanguage();

  // Reactive settings: seeds from cache then fetches + re-renders, so central
  // values (max items, detail page, new-tab preference) are honoured on first
  // paint instead of only after the cache-warming extension has run.
  const settings = useSettings(props.context);
  const navKey: string = useNavKey();
  const maxItems: number = props.maxItems && props.maxItems > 0 ? props.maxItems : settings.newsMaxItems;
  const detailUrl: string = settings.detailPageUrl;
  const newTab: boolean = settings.openLinksInNewTab;
  const showAll: boolean = maxItems > 3;

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
    getNews(props.context, props.newsList, Math.max(maxItems, 9))
      .then((result: INewsItem[]) => {
        if (active) {
          setItems(result);
          setLoading(false);
        }
      })
      .catch(() => {
        /* service already returns fallback */
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [props.context, props.newsList, maxItems, navKey]);

  // Client-side category filter: "All" shows everything, otherwise substring match.
  const filtered: INewsItem[] =
    activeCategory === 'All'
      ? items
      : items.filter((item: INewsItem) =>
          (item.category || '').toLowerCase().indexOf(activeCategory.toLowerCase()) !== -1
        );
  const visibleItems: INewsItem[] = filtered.slice(0, maxItems);

  return (
    <section className={styles.news} dir={isRtl(language) ? 'rtl' : 'ltr'}>
      <SectionHeader title={props.title} linkText={props.linkText} linkUrl={props.archiveUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.newsCats}>
        {NEWS_CATEGORIES.map((cat: string) => (
          <button
            key={cat}
            type="button"
            className={`${styles.ncat} ${activeCategory === cat ? styles.on : ''}`}
            onClick={() => setActiveCategory(cat)}
          >
            {localizeChoice(cat, language)}
          </button>
        ))}
      </div>

      {loading ? (
        <div className={styles.newsEmpty}>{t('loadingNews', language)}</div>
      ) : visibleItems.length === 0 ? (
        <div className={styles.newsEmpty}>{t('noNews', language)}</div>
      ) : (
        <div className={styles.newsG}>
          {visibleItems.map((item: INewsItem, index: number) => {
            const inner: JSX.Element = (
              <>
                <div
                  className={styles.ncImg}
                  aria-hidden="true"
                  style={
                    item.imageUrl
                      ? { backgroundImage: `url("${item.imageUrl}")`, backgroundSize: 'cover', backgroundPosition: 'center' }
                      : { background: item.imageGradient }
                  }
                />
                <div className={styles.ncBd}>
                  <div className={styles.ncTag}>{localizeChoice(item.category, language)}</div>
                  <h3 className={styles.ncH}>{item.title}</h3>
                  {item.excerpt ? <p className={styles.ncEx}>{item.excerpt}</p> : null}
                  <div className={styles.ncFt}>
                    <span className={styles.ncDt}>{item.date}</span>
                    <span className={styles.ncAu}>{item.author}</span>
                  </div>
                </div>
              </>
            );
            const key: string = `${item.title}-${index}`;
            const href: string | undefined = hrefFor(item);
            return href ? (
              <a key={key} className={`${styles.nc} ${styles.ncLink}`} href={href} {...linkTarget(newTab)}>
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

      {showAll ? null : (
        <div className={styles.newsNav} aria-hidden="true">
          {[0, 1, 2, 3].map((i: number) => (
            <span key={i} className={`${styles.newsDot} ${i === 0 ? styles.on : ''}`} />
          ))}
        </div>
      )}
    </section>
  );
};

export default NewsCarousel;
