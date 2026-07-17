import * as React from 'react';
import styles from './DocumentsTab.module.scss';
import { IDocCtx } from '../IDocTabProps';
import { IDocEntry, DocFileType, DocStage, setStage, versionHistoryUrl, isOverdue } from '../../services/DocumentCenterService';
import { t, localizeChoice } from '../../../../common/services/uiStrings';
import { formatLongDate } from '../../../../common/util/format';

const ALL: string = '__all__';

function distinct(docs: IDocEntry[], pick: (d: IDocEntry) => string): string[] {
  const seen: { [k: string]: boolean } = {};
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

function fileTypeClass(ft: DocFileType): string {
  switch (ft) {
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
function fileTypeLabel(ft: DocFileType): string {
  return ft === 'other' ? 'FILE' : ft.toUpperCase();
}
function statusClass(status: string): string {
  const s: string = (status || '').toLowerCase();
  if (s.indexOf('publish') !== -1) {
    return `${styles.status} ${styles.stPublished}`;
  }
  if (s.indexOf('review') !== -1) {
    return `${styles.status} ${styles.stReview}`;
  }
  if (s.indexOf('archiv') !== -1) {
    return `${styles.status} ${styles.stArchived}`;
  }
  return `${styles.status} ${styles.stDraft}`;
}

const DocumentsTab: React.FunctionComponent<{ ctx: IDocCtx }> = ({ ctx }) => {
  const { docs, language } = ctx;
  const [search, setSearch] = React.useState<string>('');
  const [dept, setDept] = React.useState<string>(ALL);
  const [type, setType] = React.useState<string>(ALL);
  const [status, setStatus] = React.useState<string>(ALL);
  const [sens, setSens] = React.useState<string>(ALL);
  const [busyId, setBusyId] = React.useState<number>(0);

  const canApprove: boolean = ctx.role === 'Approver' || ctx.role === 'Administrator';

  const departments: string[] = distinct(docs, (d) => d.department);
  const types: string[] = distinct(docs, (d) => d.docType);
  const statuses: string[] = distinct(docs, (d) => d.status);
  const sensitivities: string[] = distinct(docs, (d) => d.sensitivity);

  const filtered: IDocEntry[] = docs.filter((d) => {
    const q: string = search.trim().toLowerCase();
    const matchesSearch: boolean =
      q.length === 0 ||
      d.title.toLowerCase().indexOf(q) !== -1 ||
      (d.documentNumber || '').toLowerCase().indexOf(q) !== -1 ||
      d.tags.some((tg) => tg.toLowerCase().indexOf(q) !== -1);
    return (
      matchesSearch &&
      (dept === ALL || d.department === dept) &&
      (type === ALL || d.docType === type) &&
      (status === ALL || d.status === status) &&
      (sens === ALL || d.sensitivity === sens)
    );
  });

  const changeStage = (id: number, stage: DocStage): void => {
    setBusyId(id);
    setStage(ctx.context, ctx.libraryTitle, id, stage, ctx.siteUrl)
      .then(() => ctx.reload())
      .catch(() => {
        /* ignore */
      })
      .then(() => setBusyId(0));
  };

  return (
    <div className={styles.wrap}>
      <div className={styles.toolbar}>
        <input
          className={styles.search}
          type="text"
          value={search}
          placeholder={t('searchDocuments', language)}
          aria-label={t('searchDocuments', language)}
          onChange={(e) => setSearch(e.target.value)}
        />
        <select className={styles.select} aria-label={t('department', language)} value={dept} onChange={(e) => setDept(e.target.value)}>
          <option value={ALL}>{t('allDepartments', language)}</option>
          {departments.map((d) => (
            <option key={d} value={d}>
              {localizeChoice(d, language)}
            </option>
          ))}
        </select>
        <select className={styles.select} aria-label={t('type', language)} value={type} onChange={(e) => setType(e.target.value)}>
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
        <select className={styles.select} aria-label={t('sensitivity', language)} value={sens} onChange={(e) => setSens(e.target.value)}>
          <option value={ALL}>{t('allSensitivities', language)}</option>
          {sensitivities.map((sv) => (
            <option key={sv} value={sv}>
              {localizeChoice(sv, language)}
            </option>
          ))}
        </select>
      </div>

      <div className={styles.count}>
        {filtered.length} {t('documentsCount', language)}
      </div>

      {filtered.length === 0 ? (
        <div className={styles.empty}>
          <div>{t('noDocumentsMatch', language)}</div>
          {docs.length === 0 ? (
            <div className={styles.diag}>
              <div>
                Library: &quot;{ctx.libraryTitle}&quot;
                {ctx.siteUrl ? <> · Site: {ctx.siteUrl}</> : null} · Role: {ctx.role} · Loaded: {docs.length} · v1.6.5
              </div>
              {ctx.loadError ? <div>Load error: {ctx.loadError}</div> : null}
            </div>
          ) : null}
        </div>
      ) : (
        <div className={styles.list}>
          {filtered.map((d) => {
            const overdue: boolean = isOverdue(d);
            return (
              <div className={styles.row} key={d.id}>
                <div className={`${styles.dic} ${fileTypeClass(d.fileType)}`}>{fileTypeLabel(d.fileType)}</div>
                <div className={styles.info}>
                  <div className={styles.nmRow}>
                    {d.documentNumber ? <span className={styles.docNum}>{d.documentNumber}</span> : null}
                    <span className={styles.nm}>{d.title}</span>
                  </div>
                  {d.description ? <div className={styles.desc}>{d.description}</div> : null}
                  <div className={styles.metaRow}>
                    {d.department ? <span className={styles.tag}>{localizeChoice(d.department, language)}</span> : null}
                    {d.docType ? <span className={styles.tag}>{localizeChoice(d.docType, language)}</span> : null}
                    {d.sensitivity ? <span className={styles.tag}>{localizeChoice(d.sensitivity, language)}</span> : null}
                    {d.tags.map((tg) => (
                      <span className={styles.hash} key={tg}>
                        #{tg}
                      </span>
                    ))}
                    {d.reviewDate ? (
                      <span className={overdue ? `${styles.metaBit} ${styles.overdue}` : styles.metaBit}>
                        {t('reviewDate', language)}: {formatLongDate(d.reviewDate)}
                      </span>
                    ) : null}
                  </div>
                </div>
                <div className={styles.right}>
                  {d.status ? <span className={statusClass(d.status)}>{localizeChoice(d.status, language)}</span> : null}
                  <div className={styles.actions}>
                    {d.url ? (
                      <a className={styles.act} href={d.url} target="_blank" rel="noopener noreferrer">
                        {t('open', language)}
                      </a>
                    ) : null}
                    <a className={styles.act} href={versionHistoryUrl(ctx.webUrl, ctx.listId, d.id)} target="_blank" rel="noopener noreferrer">
                      {t('versionHistory', language)}
                    </a>
                    {canApprove && d.status === 'Draft' ? (
                      <button className={styles.actBtn} disabled={busyId === d.id} onClick={() => changeStage(d.id, 'In Review')} type="button">
                        {t('submitForReview', language)}
                      </button>
                    ) : null}
                    {canApprove && d.status === 'In Review' ? (
                      <button className={styles.actBtn} disabled={busyId === d.id} onClick={() => changeStage(d.id, 'Published')} type="button">
                        {t('publish', language)}
                      </button>
                    ) : null}
                    {canApprove && d.status !== 'Archived' ? (
                      <button className={styles.actBtnGhost} disabled={busyId === d.id} onClick={() => changeStage(d.id, 'Archived')} type="button">
                        {t('archive', language)}
                      </button>
                    ) : null}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default DocumentsTab;
