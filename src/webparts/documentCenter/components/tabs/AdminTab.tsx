import * as React from 'react';
import styles from './AdminTab.module.scss';
import { IDocCtx } from '../IDocTabProps';
import { getChoices, updateChoices } from '../../services/DocumentCenterService';
import { t, localizeChoice } from '../../../../common/services/uiStrings';

/** One editable taxonomy column. */
interface ITaxonomyDef {
  field: string;
  heading: string;
}

/** Per-section editing state. */
interface ISectionState {
  values: string[];
  draft: string;
  saving: boolean;
  saved: boolean;
}

const AdminTab: React.FunctionComponent<{ ctx: IDocCtx }> = ({ ctx }) => {
  const language = ctx.language;

  const columns: ITaxonomyDef[] = React.useMemo(
    () => [
      { field: 'Category', heading: t('department', language) },
      { field: 'DocumentType', heading: t('documentTypes', language) },
      { field: 'DocStatus', heading: t('statuses', language) },
      { field: 'Sensitivity', heading: t('sensitivities', language) }
    ],
    [language]
  );

  const [loading, setLoading] = React.useState<boolean>(true);
  const [sections, setSections] = React.useState<Record<string, ISectionState>>({});

  React.useEffect(() => {
    let active = true;
    setLoading(true);
    Promise.all(columns.map((c) => getChoices(ctx.context, ctx.libraryTitle, c.field)))
      .then((results) => {
        if (!active) {
          return;
        }
        const next: Record<string, ISectionState> = {};
        columns.forEach((c, i) => {
          next[c.field] = { values: results[i], draft: '', saving: false, saved: false };
        });
        setSections(next);
        setLoading(false);
      })
      .catch(() => {
        if (!active) {
          return;
        }
        const next: Record<string, ISectionState> = {};
        columns.forEach((c) => {
          next[c.field] = { values: [], draft: '', saving: false, saved: false };
        });
        setSections(next);
        setLoading(false);
      });
    return () => {
      active = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ctx.context, ctx.libraryTitle]);

  const patch = (field: string, change: Partial<ISectionState>): void => {
    setSections((prev) => {
      const cur = prev[field];
      if (!cur) {
        return prev;
      }
      return { ...prev, [field]: { ...cur, ...change } };
    });
  };

  const removeValue = (field: string, index: number): void => {
    const cur = sections[field];
    if (!cur) {
      return;
    }
    const values = cur.values.filter((_, i) => i !== index);
    patch(field, { values, saved: false });
  };

  const addValue = (field: string): void => {
    const cur = sections[field];
    if (!cur) {
      return;
    }
    const raw = cur.draft.trim();
    if (!raw || cur.values.indexOf(raw) !== -1) {
      patch(field, { draft: '' });
      return;
    }
    patch(field, { values: cur.values.concat(raw), draft: '', saved: false });
  };

  const save = (field: string): void => {
    const cur = sections[field];
    if (!cur) {
      return;
    }
    patch(field, { saving: true, saved: false });
    updateChoices(ctx.context, ctx.libraryTitle, field, cur.values)
      .then((ok) => {
        patch(field, { saving: false, saved: ok });
      })
      .catch(() => {
        patch(field, { saving: false, saved: false });
      });
  };

  if (loading) {
    return (
      <div className={styles.wrap}>
        <div className={styles.loading}>{t('loading', language)}</div>
      </div>
    );
  }

  return (
    <div className={styles.wrap}>
      <h2 className={styles.heading}>{t('manageTaxonomies', language)}</h2>

      <div className={styles.grid}>
        {columns.map((col) => {
          const section = sections[col.field];
          if (!section) {
            return null;
          }
          return (
            <section key={col.field} className={styles.card}>
              <h3 className={styles.cardTitle}>{col.heading}</h3>

              <ul className={styles.chips}>
                {section.values.length === 0 ? (
                  <li className={styles.emptyRow}>&mdash;</li>
                ) : (
                  section.values.map((value, index) => (
                    <li key={value} className={styles.chip}>
                      <span className={styles.chipLabel}>{localizeChoice(value, language)}</span>
                      <button
                        type="button"
                        className={styles.chipRemove}
                        onClick={() => removeValue(col.field, index)}
                        aria-label={t('remove', language)}
                        title={t('remove', language)}
                      >
                        &times;
                      </button>
                    </li>
                  ))
                )}
              </ul>

              <div className={styles.addRow}>
                <input
                  className={styles.input}
                  type="text"
                  value={section.draft}
                  onChange={(e) => patch(col.field, { draft: e.target.value })}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault();
                      addValue(col.field);
                    }
                  }}
                />
                <button type="button" className={styles.addBtn} onClick={() => addValue(col.field)}>
                  {t('addValue', language)}
                </button>
              </div>

              <div className={styles.cardActions}>
                <button
                  type="button"
                  className={styles.saveBtn}
                  disabled={section.saving}
                  onClick={() => save(col.field)}
                >
                  {t('save', language)}
                </button>
                {section.saved ? <span className={styles.savedMsg}>{t('saved', language)}</span> : null}
              </div>
            </section>
          );
        })}
      </div>

      <section className={styles.panel}>
        <h3 className={styles.panelTitle}>{t('numberingFormat', language)}</h3>
        <p className={styles.pattern}>DEPT-YYYY-0000</p>
        <p className={styles.patternHint}>HR-2026-0007</p>
      </section>

      <p className={styles.infoBanner}>{t('approvalRoutingNote', language)}</p>
    </div>
  );
};

export default AdminTab;
