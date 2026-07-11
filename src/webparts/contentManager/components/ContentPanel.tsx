import * as React from 'react';
import styles from './ContentManager.module.scss';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Icon
} from '@fluentui/react';
import { IContentTypeDef, IFieldDef } from '../services/contentTypes';
import { IContentRow, getItems, getChoices, addItem, updateItem, deleteItem } from '../services/ContentService';
import { formatLongDate } from '../../../common/util/format';
import RichTextField from './RichTextField';

interface IContentPanelProps {
  context: WebPartContext;
  listTitle: string;
  def: IContentTypeDef;
}

type FormValues = { [field: string]: string };

const pad = (n: number): string => n.toString().padStart(2, '0');

/** Converts a stored UTC ISO date to the viewer's local datetime-local value. */
function toInputDateTime(iso: string | undefined): string {
  if (!iso) {
    return '';
  }
  const d: Date = new Date(iso);
  if (isNaN(d.getTime())) {
    return '';
  }
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/** Converts a local datetime-local value to a UTC ISO string for SharePoint. */
function toISO(local: string): string | undefined {
  if (!local) {
    return undefined;
  }
  const d: Date = new Date(local); // datetime-local is interpreted in local time
  return isNaN(d.getTime()) ? undefined : d.toISOString();
}

const ContentPanel: React.FunctionComponent<IContentPanelProps> = (props) => {
  const { context, listTitle, def } = props;
  const [rows, setRows] = React.useState<IContentRow[]>([]);
  const [choices, setChoices] = React.useState<{ [field: string]: string[] }>({});
  const [loading, setLoading] = React.useState<boolean>(true);
  const [editingId, setEditingId] = React.useState<number | 'new' | undefined>(undefined);
  const [values, setValues] = React.useState<FormValues>({});
  const [saving, setSaving] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [notice, setNotice] = React.useState<string>('');

  const selectFields: string[] = React.useMemo(() => {
    const set: string[] = ['Id', 'Title'];
    def.fields.forEach((f) => {
      if (set.indexOf(f.name) === -1) {
        set.push(f.name);
      }
    });
    return set;
  }, [def]);

  const reload = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    const items: IContentRow[] = await getItems(context, listTitle, selectFields, def.orderBy);
    setRows(items);
    setLoading(false);
  }, [context, listTitle, selectFields, def.orderBy]);

  React.useEffect(() => {
    let active: boolean = true;
    (async (): Promise<void> => {
      // Fetch choices for any choice fields (works for main and department lists).
      const choiceFields: IFieldDef[] = def.fields.filter((f) => f.kind === 'choice');
      const map: { [field: string]: string[] } = {};
      for (const f of choiceFields) {
        map[f.name] = f.choices || (await getChoices(context, listTitle, f.name));
      }
      if (active) {
        setChoices(map);
      }
      await reload();
    })().catch(() => {
      if (active) {
        setLoading(false);
      }
    });
    return () => {
      active = false;
    };
  }, [context, listTitle, def, reload]);

  const startAdd = (): void => {
    setValues({});
    setEditingId('new');
    setError('');
    setNotice('');
  };

  const startEdit = (row: IContentRow): void => {
    const v: FormValues = {};
    def.fields.forEach((f) => {
      const raw: unknown = row[f.name];
      v[f.name] = f.kind === 'datetime' ? toInputDateTime(raw as string) : raw === undefined || raw === null ? '' : String(raw);
    });
    setValues(v);
    setEditingId(row.Id);
    setError('');
    setNotice('');
  };

  const cancel = (): void => {
    setEditingId(undefined);
    setError('');
  };

  const setField = (name: string, value: string): void => {
    setValues((prev) => ({ ...prev, [name]: value }));
  };

  const save = async (): Promise<void> => {
    // Basic required-field validation.
    const missing: IFieldDef | undefined = def.fields.filter((f) => f.required).filter((f) => !(values[f.name] || '').trim())[0];
    if (missing) {
      setError(`${missing.label} is required.`);
      return;
    }
    setSaving(true);
    setError('');
    try {
      const data: Record<string, unknown> = {};
      def.fields.forEach((f) => {
        const val: string = values[f.name] || '';
        if (f.kind === 'datetime') {
          data[f.name] = val ? toISO(val) : null;
        } else if (f.kind === 'choice') {
          // An empty choice must be null — '' is rejected as "not a valid choice".
          data[f.name] = val || null;
        } else {
          data[f.name] = val;
        }
      });
      if (editingId === 'new') {
        await addItem(context, listTitle, data);
        setNotice('Item added.');
      } else if (typeof editingId === 'number') {
        await updateItem(context, listTitle, editingId, data);
        setNotice('Item updated.');
      }
      setEditingId(undefined);
      await reload();
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Could not save. Check your permissions and try again.');
    } finally {
      setSaving(false);
    }
  };

  const remove = async (row: IContentRow): Promise<void> => {
    // eslint-disable-next-line no-alert
    if (typeof window !== 'undefined' && window.confirm && !window.confirm(`Delete "${row.Title}"? This cannot be undone.`)) {
      return;
    }
    try {
      await deleteItem(context, listTitle, row.Id);
      setNotice('Item deleted.');
      await reload();
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Could not delete the item.');
    }
  };

  const renderField = (f: IFieldDef): JSX.Element => {
    const val: string = values[f.name] || '';
    if (f.kind === 'richtext') {
      return (
        <RichTextField
          key={f.name}
          label={f.label}
          value={val}
          resetKey={`${editingId}-${f.name}`}
          onChange={(html) => setField(f.name, html)}
        />
      );
    }
    if (f.kind === 'choice') {
      const opts: IDropdownOption[] = (choices[f.name] || []).map((c) => ({ key: c, text: c }));
      return (
        <Dropdown
          key={f.name}
          label={f.label}
          selectedKey={val || undefined}
          options={opts}
          onChange={(_, o) => setField(f.name, o ? String(o.key) : '')}
        />
      );
    }
    if (f.kind === 'datetime') {
      return (
        <div key={f.name} className={styles.fieldRow}>
          <label className={styles.fieldLabel}>{f.label}</label>
          <input
            className={styles.dateInput}
            type="datetime-local"
            value={val}
            onChange={(e) => setField(f.name, e.target.value)}
          />
        </div>
      );
    }
    return (
      <TextField
        key={f.name}
        label={f.label}
        value={val}
        required={f.required}
        multiline={f.kind === 'multiline'}
        rows={f.kind === 'multiline' ? 5 : undefined}
        onChange={(_, v) => setField(f.name, v || '')}
      />
    );
  };

  return (
    <div className={styles.panel}>
      {notice ? (
        <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setNotice('')}>
          {notice}
        </MessageBar>
      ) : null}
      {error ? <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar> : null}

      {editingId !== undefined ? (
        <div className={styles.form}>
          <h3 className={styles.formTitle}>{editingId === 'new' ? `Add ${def.label.replace(/s$/, '')}` : 'Edit item'}</h3>
          {def.fields.map(renderField)}
          <div className={styles.formActions}>
            <PrimaryButton text={saving ? 'Saving…' : 'Save'} disabled={saving} onClick={save} />
            <DefaultButton text="Cancel" disabled={saving} onClick={cancel} />
          </div>
        </div>
      ) : (
        <div className={styles.toolbar}>
          <PrimaryButton iconProps={{ iconName: 'Add' }} text={`Add ${def.label.replace(/s$/, '')}`} onClick={startAdd} />
          <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={reload} />
        </div>
      )}

      {loading ? (
        <Spinner size={SpinnerSize.medium} label="Loading…" />
      ) : rows.length === 0 ? (
        <div className={styles.empty}>No {def.label.toLowerCase()} yet. Use “Add” to create the first item.</div>
      ) : (
        <div className={styles.list}>
          {rows.map((row) => {
            const sub: string = def.subtitleField ? String(row[def.subtitleField] || '') : '';
            const dateStr: string = row[def.orderBy] ? formatLongDate(String(row[def.orderBy])) : '';
            return (
              <div className={styles.rowItem} key={row.Id}>
                <div className={styles.rowMain}>
                  <div className={styles.rowTitle}>{row.Title || '(untitled)'}</div>
                  <div className={styles.rowMeta}>{[sub, dateStr].filter((v) => !!v).join(' · ')}</div>
                </div>
                <div className={styles.rowActions}>
                  <button className={styles.iconBtn} title="Edit" onClick={() => startEdit(row)}>
                    <Icon iconName="Edit" />
                  </button>
                  <button className={styles.iconBtn} title="Delete" onClick={() => remove(row)}>
                    <Icon iconName="Delete" />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

export default ContentPanel;
