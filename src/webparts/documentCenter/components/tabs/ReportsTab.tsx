import * as React from 'react';
import styles from './ReportsTab.module.scss';
import { IDocCtx } from '../IDocTabProps';
import { IDocStats, IDocEntry, isOverdue } from '../../services/DocumentCenterService';
import { t, localizeChoice } from '../../../../common/services/uiStrings';
import { formatLongDate } from '../../../../common/util/format';
import { Language } from '../../../../common/services/languageService';

interface IChartRow {
  key: string;
  value: number;
}

/** Turn an aggregate map into a sorted (desc) list of rows. */
function toRows(map: { [k: string]: number }): IChartRow[] {
  const rows: IChartRow[] = [];
  Object.keys(map).forEach((k) => {
    const v: number = map[k];
    if (v > 0) {
      rows.push({ key: k, value: v });
    }
  });
  return rows.sort((a, b) => b.value - a.value);
}

interface IChartProps {
  heading: string;
  map: { [k: string]: number };
  language: Language;
  tintClass: string;
}

const BarChart: React.FunctionComponent<IChartProps> = ({ heading, map, language, tintClass }) => {
  const rows: IChartRow[] = toRows(map);
  const max: number = rows.reduce((m, r) => (r.value > m ? r.value : m), 0);

  return (
    <div className={styles.chart}>
      <h3 className={styles.chartHeading}>{heading}</h3>
      {rows.length === 0 || max === 0 ? (
        <div className={styles.chartEmpty}>&mdash;</div>
      ) : (
        <div className={styles.bars}>
          {rows.map((r) => {
            const pct: number = Math.round((r.value / max) * 100);
            return (
              <div className={styles.barRow} key={r.key}>
                <div className={styles.barLabel}>{localizeChoice(r.key, language)}</div>
                <div className={styles.barTrack}>
                  <div className={`${styles.barFill} ${tintClass}`} style={{ width: pct + '%' }} />
                </div>
                <div className={styles.barValue}>{r.value}</div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

const ReportsTab: React.FunctionComponent<{ ctx: IDocCtx }> = ({ ctx }) => {
  const language: Language = ctx.language;
  const stats: IDocStats = ctx.stats;

  const overdueDocs: IDocEntry[] = ctx.docs.filter((d) => isOverdue(d));

  return (
    <div className={styles.wrap}>
      <div className={styles.grid}>
        <BarChart heading={t('byStatus', language)} map={stats.byStatus} language={language} tintClass={styles.tintA} />
        <BarChart heading={t('byDepartment', language)} map={stats.byDepartment} language={language} tintClass={styles.tintB} />
        <BarChart heading={t('byType', language)} map={stats.byType} language={language} tintClass={styles.tintC} />
        <BarChart heading={t('bySensitivity', language)} map={stats.bySensitivity} language={language} tintClass={styles.tintD} />
      </div>

      <div className={styles.panel}>
        <h3 className={styles.panelHeading}>{t('overdueReview', language)}</h3>
        {overdueDocs.length === 0 ? (
          <div className={styles.nothing}>{t('nothingOverdue', language)}</div>
        ) : (
          <div className={styles.overdueList}>
            {overdueDocs.map((d) => (
              <div className={styles.overdueRow} key={d.id}>
                <div className={styles.overdueMain}>
                  {d.documentNumber ? <span className={styles.overdueNum}>{d.documentNumber}</span> : null}
                  <span className={styles.overdueTitle}>{d.title}</span>
                </div>
                <div className={styles.overdueMeta}>
                  {d.department ? <span className={styles.overdueDept}>{localizeChoice(d.department, language)}</span> : null}
                  {d.reviewDate ? <span className={styles.overdueDate}>{formatLongDate(d.reviewDate)}</span> : null}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default ReportsTab;
