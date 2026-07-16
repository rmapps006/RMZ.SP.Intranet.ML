import * as React from 'react';
import styles from './DashboardTab.module.scss';
import { IDocCtx } from '../IDocTabProps';
import { IDocEntry, versionHistoryUrl, isOverdue } from '../../services/DocumentCenterService';
import { t, localizeChoice } from '../../../../common/services/uiStrings';
import { formatLongDate } from '../../../../common/util/format';

interface IKpi {
  key: string;
  label: string;
  value: number;
  warn: boolean;
}

function statusClass(status: string): string {
  const s: string = (status || '').toLowerCase();
  if (s.indexOf('publish') !== -1) {
    return `${styles.pill} ${styles.stPublished}`;
  }
  if (s.indexOf('review') !== -1) {
    return `${styles.pill} ${styles.stReview}`;
  }
  if (s.indexOf('archiv') !== -1) {
    return `${styles.pill} ${styles.stArchived}`;
  }
  return `${styles.pill} ${styles.stDraft}`;
}

const DashboardTab: React.FunctionComponent<{ ctx: IDocCtx }> = ({ ctx }) => {
  const language = ctx.language;
  const { stats, docs } = ctx;

  const kpis: IKpi[] = [
    { key: 'total', label: t('totalDocuments', language), value: stats.total, warn: false },
    { key: 'published', label: t('publishedCount', language), value: stats.published, warn: false },
    { key: 'inReview', label: t('inReviewCount', language), value: stats.inReview, warn: false },
    { key: 'draft', label: t('draftCount', language), value: stats.draft, warn: false },
    { key: 'overdue', label: t('overdueReview', language), value: stats.overdue, warn: stats.overdue > 0 },
    { key: 'my', label: t('myDocuments', language), value: stats.myCount, warn: false }
  ];

  const overdueDocs: IDocEntry[] = docs.filter(isOverdue);
  const recentDocs: IDocEntry[] = docs.slice(0, 6);

  return (
    <div className={styles.wrap}>
      <div className={styles.tiles}>
        {kpis.map((k) => (
          <div className={k.warn ? `${styles.tile} ${styles.tileWarn}` : styles.tile} key={k.key}>
            <div className={styles.tileNum}>{k.value}</div>
            <div className={styles.tileLabel}>{k.label}</div>
          </div>
        ))}
      </div>

      <div className={styles.panels}>
        <section className={styles.panel}>
          <h3 className={styles.panelTitle}>{t('needsAttention', language)}</h3>
          {overdueDocs.length === 0 ? (
            <div className={styles.empty}>{t('nothingOverdue', language)}</div>
          ) : (
            <div className={styles.list}>
              {overdueDocs.map((d) => (
                <div className={styles.attRow} key={d.id}>
                  <div className={styles.attInfo}>
                    <div className={styles.attTitleRow}>
                      {d.documentNumber ? <span className={styles.docNum}>{d.documentNumber}</span> : null}
                      <span className={styles.nm}>{d.title}</span>
                    </div>
                    {d.department ? <span className={styles.attDept}>{localizeChoice(d.department, language)}</span> : null}
                  </div>
                  {d.reviewDate ? (
                    <span className={styles.attDue}>
                      {t('reviewDate', language)}: {formatLongDate(d.reviewDate)}
                    </span>
                  ) : null}
                </div>
              ))}
            </div>
          )}
        </section>

        <section className={styles.panel}>
          <h3 className={styles.panelTitle}>{t('recentDocuments', language)}</h3>
          {recentDocs.length === 0 ? (
            <div className={styles.empty}>{t('noDocumentsMatch', language)}</div>
          ) : (
            <div className={styles.list}>
              {recentDocs.map((d) => (
                <div className={styles.recRow} key={d.id}>
                  <div className={styles.recInfo}>
                    <div className={styles.recTitleRow}>
                      {d.documentNumber ? <span className={styles.docNum}>{d.documentNumber}</span> : null}
                      <span className={styles.nm}>{d.title}</span>
                      {d.status ? <span className={statusClass(d.status)}>{localizeChoice(d.status, language)}</span> : null}
                    </div>
                    <div className={styles.recMeta}>
                      <span className={styles.recDate}>{formatLongDate(d.modified)}</span>
                      <span className={styles.recActions}>
                        {d.url ? (
                          <a className={styles.act} href={d.url} target="_blank" rel="noopener noreferrer">
                            {t('open', language)}
                          </a>
                        ) : null}
                        <a
                          className={styles.act}
                          href={versionHistoryUrl(ctx.webUrl, ctx.listId, d.id)}
                          target="_blank"
                          rel="noopener noreferrer"
                        >
                          {t('versionHistory', language)}
                        </a>
                      </span>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </section>
      </div>
    </div>
  );
};

export default DashboardTab;
