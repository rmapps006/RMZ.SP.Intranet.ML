import * as React from 'react';
import styles from './DocumentCenter.module.scss';
import { IDocumentCenterProps } from './IDocumentCenterProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { getDocuments, IDocEntry, IDocCenterResult, DocFileType } from '../services/DocumentCenterService';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget, formatLongDate } from '../../../common/util/format';
import { getCurrentLanguage, isRtl, Language } from '../../../common/services/languageService';
import { t, localizeChoice } from '../../../common/services/uiStrings';

const ALL: string = '__all__';

/** Distinct, non-empty values for a field across the document set (for filter options). */
function distinct(docs: IDocEntry[], pick: (d: IDocEntry) => string): string[] {
  const seen: { [key: string]: boolean } = {};
  const out: string[] = [];
  docs.forEach((d) => {
    const v: string = (pick(d) || '').trim();
    if (v && !seen[v]) {
      seen[v] = true;
      out.push(v);
    }
  });
  return out.sort();
}

function fileTypeClass(fileType: DocFileType): string {
  switch (fileType) {
    case 'pdf':
      return styles.pdf;
    case 'xls':
      return styles.xls;
    case 'ppt':
      return styles.ppt;
    case 'doc':
      return styles.doc;
    default:
      return styles.other;
  }
}

function fileTypeLabel(fileType: DocFileType): string {
  return fileType === 'other' ? 'FILE' : fileType.toUpperCase();
}

function statusClass(status: string): string {
  const s: string = (status || '').toLowerCase();
  if (s.indexOf('approved') !== -1) {
    return `${styles.status} ${styles.stApproved}`;
  }
  if (s.indexOf('review') !== -1) {
    return `${styles.status} ${styles.stReview}`;
  }
  if (s.indexOf('archiv') !== -1) {
    return `${styles.status} ${styles.stArchived}`;
  }
  return `${styles.status} ${styles.stDraft}`;
}

const DocumentCenter: React.FunctionComponent<IDocumentCenterProps> = (props) => {
  const [docs, setDocs] = React.useState<IDocEntry[]>([]);
  const [libraryUrl, setLibraryUrl] = React.useState<string | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [search, setSearch] = React.useState<string>('');
  const [category, setCategory] = React.useState<string>(ALL);
  const [docType, setDocType] = React.useState<string>(ALL);
  const [status, setStatus] = React.useState<string>(ALL);

  const language: Language = getCurrentLanguage();
  const navKey: string = useNavKey();
  const newTab: boolean = false;

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    getDocuments(props.context, props.libraryTitle, props.pageSize)
      .then((result: IDocCenterResult) => {
        if (active) {
          setDocs(result.documents);
          setLibraryUrl(result.libraryUrl);
        }
      })
      .catch(() => {
        /* service already returns an empty set */
      })
      .finally(() => {
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [props.context, props.libraryTitle, props.pageSize, navKey]);

  const categories: string[] = distinct(docs, (d) => d.category);
  const types: string[] = distinct(docs, (d) => d.docType);
  const statuses: string[] = distinct(docs, (d) => d.status);

  const filtered: IDocEntry[] = docs.filter((d) => {
    const matchesCat: boolean = category === ALL || d.category === category;
    const matchesType: boolean = docType === ALL || d.docType === docType;
    const matchesStatus: boolean = status === ALL || d.status === status;
    const q: string = search.trim().toLowerCase();
    const matchesSearch: boolean =
      q.length === 0 ||
      d.title.toLowerCase().indexOf(q) !== -1 ||
      (d.description || '').toLowerCase().indexOf(q) !== -1;
    return matchesCat && matchesType && matchesStatus && matchesSearch;
  });

  const catChips: { value: string; label: string }[] = [
    { value: ALL, label: localizeChoice('All', language) },
    ...categories.map((c) => ({ value: c, label: localizeChoice(c, language) }))
  ];

  return (
    <section className={styles.dc} dir={isRtl(language) ? 'rtl' : 'ltr'}>
      <SectionHeader title={props.title} linkText={props.linkText} linkUrl={libraryUrl || ''} showTitle={props.showTitle} showLink={props.showViewAll && !!libraryUrl} />

      <div className={styles.toolbar}>
        <div className={styles.search}>
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" aria-hidden="true">
            <circle cx="11" cy="11" r="7" stroke="currentColor" strokeWidth="2" />
            <path d="M21 21l-4.3-4.3" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
          </svg>
          <input
            className={styles.searchInput}
            type="text"
            value={search}
            placeholder={t('searchDocuments', language)}
            aria-label={t('searchDocuments', language)}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
        <div className={styles.filters}>
          <select className={styles.select} aria-label={t('type', language)} value={docType} onChange={(e) => setDocType(e.target.value)}>
            <option value={ALL}>{t('allTypes', language)}</option>
            {types.map((ty) => (
              <option key={ty} value={ty}>
                {localizeChoice(ty, language)}
              </option>
            ))}
          </select>
          <select className={styles.select} aria-label={t('status', language)} value={status} onChange={(e) => setStatus(e.target.value)}>
            <option value={ALL}>{t('allStatuses', language)}</option>
            {statuses.map((st) => (
              <option key={st} value={st}>
                {localizeChoice(st, language)}
              </option>
            ))}
          </select>
        </div>
      </div>

      {categories.length > 0 ? (
        <div className={styles.chips}>
          {catChips.map((chip) => (
            <button
              key={chip.value}
              type="button"
              aria-pressed={chip.value === category}
              className={chip.value === category ? `${styles.chip} ${styles.on}` : styles.chip}
              onClick={() => setCategory(chip.value)}
            >
              {chip.label}
            </button>
          ))}
        </div>
      ) : null}

      {loading ? (
        <div className={styles.empty}>{t('loading', language)}</div>
      ) : docs.length === 0 ? (
        <div className={styles.empty}>{t('noDocuments', language)}</div>
      ) : filtered.length === 0 ? (
        <div className={styles.empty}>{t('noDocumentsMatch', language)}</div>
      ) : (
        <>
          <div className={styles.count}>
            {filtered.length} {t('documentsCount', language)}
          </div>
          <div className={styles.list}>
            {filtered.map((d, i) => {
              const inner: JSX.Element = (
                <>
                  <div className={`${styles.dic} ${fileTypeClass(d.fileType)}`}>{fileTypeLabel(d.fileType)}</div>
                  <div className={styles.info}>
                    <div className={styles.nm}>{d.title}</div>
                    {d.description ? <div className={styles.desc}>{d.description}</div> : null}
                    <div className={styles.metaRow}>
                      {d.category ? <span className={styles.tag}>{localizeChoice(d.category, language)}</span> : null}
                      {d.docType ? <span className={styles.tag}>{localizeChoice(d.docType, language)}</span> : null}
                      {d.owner ? <span className={styles.metaBit}>{t('owner', language)}: {d.owner}</span> : null}
                      {d.reviewDate ? <span className={styles.metaBit}>{t('reviewDate', language)}: {formatLongDate(d.reviewDate)}</span> : null}
                      {d.modified ? <span className={styles.metaBit}>{formatLongDate(d.modified)}</span> : null}
                    </div>
                  </div>
                  {d.status ? <span className={statusClass(d.status)}>{localizeChoice(d.status, language)}</span> : null}
                </>
              );
              const key: string = `${d.title}-${i}`;
              return d.url ? (
                <a
                  key={key}
                  className={`${styles.row} ${styles.rowLink}`}
                  href={d.url}
                  {...linkTarget(newTab)}
                  style={{ textDecoration: 'none', color: 'inherit' }}
                >
                  {inner}
                </a>
              ) : (
                <div key={key} className={styles.row}>
                  {inner}
                </div>
              );
            })}
          </div>
        </>
      )}
    </section>
  );
};

export default DocumentCenter;
