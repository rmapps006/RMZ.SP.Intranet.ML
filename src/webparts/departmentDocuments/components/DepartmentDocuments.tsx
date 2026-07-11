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

const DepartmentDocuments: React.FunctionComponent<IDepartmentDocumentsProps> = (props) => {
  const [panels, setPanels] = React.useState<IDeptDocPanel[]>([]);

  // Property-pane values win; otherwise fall back to the central Department Admin settings.
  const ds = useDepartmentSettings(props.context);
  const policiesLibrary: string = props.policiesLibrary || ds.policiesLibrary;
  const documentsLibrary: string = props.documentsLibrary || ds.documentsLibrary;
  const department: string = props.department || ds.departmentName;
  const panel1Title: string = props.panel1Title || ds.panel1Title;
  const panel2Title: string = props.panel2Title || ds.panel2Title;
  const documentHubUrl: string = props.documentHubUrl || ds.documentHubUrl;
  const newTab: boolean = ds.openLinksInNewTab;
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;
    getDepartmentDocuments(props.context, policiesLibrary, documentsLibrary, department, panel1Title, panel2Title)
      .then((result: IDeptDocPanel[]) => {
        if (active) {
          setPanels(result);
        }
      })
      .catch(() => {
        /* service already returns fallback */
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
      <SectionHeader title="Forms & Documents" linkText="Document Hub" linkUrl={documentHubUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.docsG}>
        {panels.map((panel: IDeptDocPanel, pIndex: number) => (
          <div key={`${panel.title}-${pIndex}`} className={styles.dp}>
            <div className={styles.dpHd}>
              <div className={styles.dpTt}>{panel.title}</div>
              <a className={styles.lnk} href="#">
                Open Library
              </a>
            </div>
            {panel.items.map((item: IDeptDocItem, index: number) => {
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
                <a key={key} className={styles.di} href={item.url} {...linkTarget(newTab)} style={{ textDecoration: 'none', color: 'inherit' }}>
                  {inner}
                </a>
              ) : (
                <div key={key} className={styles.di}>
                  {inner}
                </div>
              );
            })}
          </div>
        ))}
      </div>
    </section>
  );
};

export default DepartmentDocuments;
