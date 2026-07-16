import * as React from 'react';
import styles from './MyDocumentsTab.module.scss';
import { IDocCtx } from '../IDocTabProps';
import { IDocEntry, versionHistoryUrl } from '../../services/DocumentCenterService';
import { t, localizeChoice } from '../../../../common/services/uiStrings';
import { formatLongDate } from '../../../../common/util/format';

interface IStageGroup {
  status: string;
  heading: string;
  items: IDocEntry[];
}

function fileTypeClass(ft: IDocEntry['fileType']): string {
  switch (ft) {
    case 'pdf':
      return styles.pdf;
    case 'doc':
      return styles.doc;
    case 'xls':
      return styles.xls;
    case 'ppt':
      return styles.ppt;
    default:
      return styles.other;
  }
}

function fileTypeLabel(ft: IDocEntry['fileType']): string {
  return ft === 'other' ? 'FILE' : ft.toUpperCase();
}

const MyDocumentsTab: React.FunctionComponent<{ ctx: IDocCtx }> = ({ ctx }) => {
  const language = ctx.language;
  const mine: IDocEntry[] = ctx.docs.filter((d) => d.authorId === ctx.currentUserId);

  const groups: IStageGroup[] = [
    { status: 'Draft', heading: t('stageDraft', language), items: mine.filter((d) => d.status === 'Draft') },
    { status: 'In Review', heading: t('stageInReview', language), items: mine.filter((d) => d.status === 'In Review') },
    { status: 'Published', heading: t('stagePublished', language), items: mine.filter((d) => d.status === 'Published') },
    { status: 'Archived', heading: localizeChoice('Archived', language), items: mine.filter((d) => d.status === 'Archived') },
  ].filter((g) => g.items.length > 0);

  if (mine.length === 0) {
    return <div className={styles.empty}>{t('noMyDocuments', language)}</div>;
  }

  return (
    <div className={styles.wrap}>
      {groups.map((g) => (
        <section className={styles.section} key={g.status}>
          <div className={styles.sectionHead}>
            <span className={styles.sectionTitle}>{g.heading}</span>
            <span className={styles.sectionCount}>{g.items.length}</span>
          </div>
          <div className={styles.list}>
            {g.items.map((d) => (
              <div className={styles.row} key={d.id}>
                <div className={`${styles.dic} ${fileTypeClass(d.fileType)}`}>{fileTypeLabel(d.fileType)}</div>
                <div className={styles.info}>
                  <div className={styles.nmRow}>
                    {d.documentNumber ? <span className={styles.docNum}>{d.documentNumber}</span> : null}
                    <span className={styles.nm}>{d.title}</span>
                  </div>
                  <div className={styles.metaRow}>
                    {d.department ? <span className={styles.tag}>{localizeChoice(d.department, language)}</span> : null}
                    {d.docType ? <span className={styles.tag}>{localizeChoice(d.docType, language)}</span> : null}
                    {d.modified ? <span className={styles.metaBit}>{formatLongDate(d.modified)}</span> : null}
                  </div>
                </div>
                <div className={styles.actions}>
                  {d.url ? (
                    <a className={styles.act} href={d.url} target="_blank" rel="noopener noreferrer">
                      {t('open', language)}
                    </a>
                  ) : null}
                  <a className={styles.act} href={versionHistoryUrl(ctx.webUrl, ctx.listId, d.id)} target="_blank" rel="noopener noreferrer">
                    {t('versionHistory', language)}
                  </a>
                </div>
              </div>
            ))}
          </div>
        </section>
      ))}
    </div>
  );
};

export default MyDocumentsTab;
