import * as React from 'react';
import styles from './UploadTab.module.scss';
import { IDocCtx } from '../IDocTabProps';
import { uploadDocument, IUploadMeta } from '../../services/DocumentCenterService';
import { t, localizeChoice } from '../../../../common/services/uiStrings';

const DEPARTMENTS: string[] = ['HR', 'Finance', 'IT', 'Operations', 'Marketing', 'Legal', 'General'];
const TYPES: string[] = ['Policy', 'Procedure', 'Form', 'Template', 'Report', 'Guideline', 'Contract'];
const SENSITIVITIES: string[] = ['Public', 'Internal', 'Confidential', 'Restricted'];
const STAGES: string[] = ['Draft', 'In Review', 'Published'];

type State = 'idle' | 'busy' | 'ok' | 'error';

const UploadTab: React.FunctionComponent<{ ctx: IDocCtx }> = ({ ctx }) => {
  const { language } = ctx;
  const [file, setFile] = React.useState<File | undefined>(undefined);
  const [meta, setMeta] = React.useState<IUploadMeta>({
    title: '',
    titleAR: '',
    department: 'General',
    docType: 'Policy',
    sensitivity: 'Internal',
    status: 'Draft',
    tags: '',
    owner: '',
    reviewDate: '',
    description: ''
  });
  const [state, setState] = React.useState<State>('idle');
  const [message, setMessage] = React.useState<string>('');

  const set = <K extends keyof IUploadMeta>(key: K, value: IUploadMeta[K]): void => {
    setMeta((prev) => ({ ...prev, [key]: value }));
  };

  const submit = (): void => {
    if (!file) {
      return;
    }
    setState('busy');
    setMessage('');
    uploadDocument(ctx.context, ctx.libraryTitle, file, meta)
      .then((res) => {
        if (res.ok) {
          setState('ok');
          setMessage(`${t('uploadSuccess', language)}: ${res.documentNumber}`);
          setFile(undefined);
          setMeta((prev) => ({ ...prev, title: '', titleAR: '', tags: '', description: '' }));
          ctx.reload();
        } else {
          setState('error');
          setMessage(`${t('uploadFailed', language)}: ${res.error || ''}`);
        }
      })
      .catch((e) => {
        setState('error');
        setMessage(`${t('uploadFailed', language)}: ${e instanceof Error ? e.message : String(e)}`);
      });
  };

  const renderSelect = (
    label: string,
    key: keyof IUploadMeta,
    options: string[]
  ): JSX.Element => (
    <label className={styles.field}>
      <span className={styles.lbl}>{label}</span>
      <select className={styles.input} value={meta[key]} onChange={(e) => set(key, e.target.value)}>
        {options.map((o) => (
          <option key={o} value={o}>
            {localizeChoice(o, language)}
          </option>
        ))}
      </select>
    </label>
  );

  return (
    <div className={styles.wrap}>
      <p className={styles.note}>{t('autoNumberNote', language)}</p>

      <label className={styles.fileField}>
        <span className={styles.lbl}>
          {t('chooseFile', language)} <span className={styles.req}>*</span>
        </span>
        <input type="file" onChange={(e) => setFile(e.target.files && e.target.files.length > 0 ? e.target.files[0] : undefined)} />
      </label>

      <div className={styles.grid}>
        <label className={styles.field}>
          <span className={styles.lbl}>{t('titleField', language)}</span>
          <input className={styles.input} type="text" value={meta.title} onChange={(e) => set('title', e.target.value)} />
        </label>
        <label className={styles.field}>
          <span className={styles.lbl}>{t('titleARField', language)}</span>
          <input className={styles.input} type="text" dir="auto" value={meta.titleAR} onChange={(e) => set('titleAR', e.target.value)} />
        </label>
        {renderSelect(t('department', language), 'department', DEPARTMENTS)}
        {renderSelect(t('type', language), 'docType', TYPES)}
        {renderSelect(t('sensitivity', language), 'sensitivity', SENSITIVITIES)}
        {renderSelect(t('status', language), 'status', STAGES)}
        <label className={styles.field}>
          <span className={styles.lbl}>{t('ownerField', language)}</span>
          <input className={styles.input} type="text" value={meta.owner} onChange={(e) => set('owner', e.target.value)} />
        </label>
        <label className={styles.field}>
          <span className={styles.lbl}>{t('reviewDateField', language)}</span>
          <input className={styles.input} type="date" value={meta.reviewDate ? meta.reviewDate.substring(0, 10) : ''} onChange={(e) => set('reviewDate', e.target.value ? `${e.target.value}T00:00:00Z` : '')} />
        </label>
        <label className={styles.field}>
          <span className={styles.lbl}>{t('tagsLabel', language)}</span>
          <input className={styles.input} type="text" value={meta.tags} placeholder="tag1; tag2" onChange={(e) => set('tags', e.target.value)} />
        </label>
      </div>

      <label className={styles.field}>
        <span className={styles.lbl}>{t('descriptionField', language)}</span>
        <textarea className={styles.textarea} rows={3} dir="auto" value={meta.description} onChange={(e) => set('description', e.target.value)} />
      </label>

      <div className={styles.actions}>
        <button className={styles.btn} type="button" disabled={!file || state === 'busy'} onClick={submit}>
          {state === 'busy' ? t('uploading', language) : t('submit', language)}
        </button>
      </div>

      {message ? <div className={state === 'error' ? `${styles.msg} ${styles.err}` : `${styles.msg} ${styles.ok}`}>{message}</div> : null}
    </div>
  );
};

export default UploadTab;
