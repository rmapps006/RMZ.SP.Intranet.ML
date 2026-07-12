import * as React from 'react';
import styles from './DepartmentDocuments.module.scss';
import { IDepartmentDocumentsProps } from './IDepartmentDocumentsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import {
  getDepartmentDocuments,
  IDeptDocItem,
  IDeptDocPanel
} from '../services/DepartmentDocumentsService';
import { useDepartmentSettings } from '../../../common/services/useDepartmentSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';
import { getCurrentLanguage, pickLocalized, Language } from '../../../common/services/languageService';

const DepartmentDocuments: React.FunctionComponent<IDepartmentDocumentsProps> = (props) => {
  const [panels, setPanels] = React.useState<IDeptDocPanel[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);

  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const policiesLibrary: string = props.policiesLibrary || ds.policiesLibrary;
  const documentsLibrary: string = props.documentsLibrary || ds.documentsLibrary;
  const department: string = props.department || ds.departmentName;
  // The web part's own EN/AR property-pane override wins; otherwise fall back to
  // the central Department Admin settings. Both the override and the settings
  // fallback carry Arabic variants, so resolve the EN and AR sides through the
  // same fallback before localizing.
  const language: Language = getCurrentLanguage();
  const panel1Title: string = pickLocalized(props.panel1Title || ds.panel1Title, props.panel1TitleAR || ds.panel1TitleAR, language);
  const panel2Title: string = pickLocalized(props.panel2Title || ds.panel2Title, props.panel2TitleAR || ds.panel2TitleAR, language);
  const openLibraryLabel: string = pickLocalized('Open Library', 'فتح المكتبة', language);
  const documentHubUrl: string = props.documentHubUrl || ds.documentHubUrl;
  const newTab: boolean = ds.openLinksInNewTab;
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    getDepartmentDocuments(props.context, policiesLibrary, documentsLibrary, department, panel1Title, panel2Title)
      .then((result: IDeptDocPanel[]) => {
        if (active) {
          setPanels(result);
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
  }, [props.context, policiesLibrary, documentsLibrary, department, panel1Title, panel2Title, navKey]);

  const dicClass = (fileType: string): string => {
    if (fileType === 'pdf') {
      return `${styles.dic} ${styles.pdf}`;
    }
    if (fileType === 'xls') {
      return `${styles.dic} ${styles.xls}`;
    }
    return `${styles.dic} ${styles.doc}`;
  };

  return (
    <section className={styles.docs}>
      <SectionHeader title={props.title} linkText={props.linkText} linkUrl={documentHubUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.docsG}>
        {loading ? (
          <div className={styles.empty}>Loading…</div>
        ) : panels.length === 0 ? (
          <div className={styles.empty}>No documents to show.</div>
        ) : (
          panels.map((panel: IDeptDocPanel, pIndex: number) => (
            <div key={`${panel.title}-${pIndex}`} className={styles.dp}>
              <div className={styles.dpHd}>
                <div className={styles.dpTt}>{panel.title}</div>
                {/* Links to the real library (its root folder); disabled only when the library can't be resolved. */}
                {panel.libraryUrl ? (
                  <a
                    className={styles.lnk}
                    href={panel.libraryUrl}
                    aria-label={`${openLibraryLabel} — ${panel.title}`}
                    {...linkTarget(newTab)}
                  >
                    {openLibraryLabel}
                  </a>
                ) : (
                  <span className={`${styles.lnk} ${styles.lnkDisabled}`} aria-disabled="true">
                    {openLibraryLabel}
                  </span>
                )}
              </div>
              {panel.items.length === 0 ? (
                <div className={styles.empty}>No documents to show.</div>
              ) : (
                panel.items.map((item: IDeptDocItem, index: number) => {
                  const inner: JSX.Element = (
                    <>
                      <div className={dicClass(item.fileType)}>{item.fileType.toUpperCase()}</div>
                      <div className={styles.diInfo}>
                        <div className={styles.diNm}>{item.name}</div>
                        <div className={styles.diMeta}>{item.meta}</div>
                      </div>
                      {item.badge ? <span className={styles.diBadge}>{item.badge}</span> : null}
                    </>
                  );
                  const key: string = `${item.name}-${index}`;
                  return item.url ? (
                    <a
                      key={key}
                      className={`${styles.di} ${styles.diActive}`}
                      href={item.url}
                      {...linkTarget(newTab)}
                      style={{ textDecoration: 'none', color: 'inherit' }}
                    >
                      {inner}
                    </a>
                  ) : (
                    <div key={key} className={styles.di}>
                      {inner}
                    </div>
                  );
                })
              )}
            </div>
          ))
        )}
      </div>
    </section>
  );
};

export default DepartmentDocuments;
