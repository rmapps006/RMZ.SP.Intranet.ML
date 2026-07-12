import * as React from 'react';
import styles from './SharedDocuments.module.scss';
import { ISharedDocumentsProps } from './ISharedDocumentsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { getLibraryDocuments } from '../services/SharedDocumentsService';
import { IDocumentItem, IDocumentPanel } from '../../../common/models';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';
import { getCurrentLanguage, isRtl, Language } from '../../../common/services/languageService';
import { t } from '../../../common/services/uiStrings';

const FALLBACK_PANEL1: IDocumentItem[] = [];

const FALLBACK_PANEL2: IDocumentItem[] = [];

function iconClass(fileType: IDocumentItem['fileType']): string {
  if (fileType === 'pdf') {
    return styles.pdf;
  }
  if (fileType === 'xls') {
    return styles.xls;
  }
  return styles.doc;
}

const DocumentPanel: React.FunctionComponent<IDocumentPanel & { newTab: boolean; loading: boolean; language: Language }> = (panel) => {
  const newTab: boolean = panel.newTab;
  const language: Language = panel.language;
  const hasLibraryUrl: boolean = !!panel.libraryUrl;
  return (
    <div className={styles.panel}>
      <div className={styles.panelHead}>
        <div className={styles.panelTitle}>{panel.title}</div>
        <a
          className={hasLibraryUrl ? styles.panelLink : `${styles.panelLink} ${styles.disabled}`}
          href={panel.libraryUrl || '#'}
          aria-label={t('openLibrary', language)}
          aria-disabled={!hasLibraryUrl}
          {...linkTarget(newTab)}
        >
          {t('openLibrary', language)}
        </a>
      </div>
      {panel.loading ? (
        <div className={styles.empty}>{t('loading', language)}</div>
      ) : panel.items.length === 0 ? (
        <div className={styles.empty}>{t('noDocuments', language)}</div>
      ) : (
        panel.items.map((doc: IDocumentItem, index: number) => {
          const hasUrl: boolean = !!doc.url;
          return (
            <a
              className={hasUrl ? styles.item : `${styles.item} ${styles.disabled}`}
              href={doc.url || '#'}
              aria-disabled={!hasUrl}
              key={index}
              {...linkTarget(newTab)}
            >
              <div className={`${styles.icon} ${iconClass(doc.fileType)}`}>
                {doc.fileType.toUpperCase()}
              </div>
              <div className={styles.name}>{doc.name}</div>
              <div className={styles.date}>{doc.modified}</div>
            </a>
          );
        })
      )}
    </div>
  );
};

const SharedDocuments: React.FunctionComponent<ISharedDocumentsProps> = (props) => {
  const [items1, setItems1] = React.useState<IDocumentItem[]>(FALLBACK_PANEL1);
  const [items2, setItems2] = React.useState<IDocumentItem[]>(FALLBACK_PANEL2);
  const [loading1, setLoading1] = React.useState<boolean>(true);
  const [loading2, setLoading2] = React.useState<boolean>(true);
  const language: Language = getCurrentLanguage();
  const newTab: boolean = useSettings(props.context).openLinksInNewTab;
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;
    setLoading1(true);
    setLoading2(true);

    getLibraryDocuments(props.context, props.panel1Library)
      .then((docs: IDocumentItem[]) => {
        if (active) {
          if (docs.length > 0) {
            setItems1(docs);
          }
          setLoading1(false);
        }
      })
      .catch(() => {
        /* keep design fallback */
        if (active) {
          setLoading1(false);
        }
      });

    getLibraryDocuments(props.context, props.panel2Library)
      .then((docs: IDocumentItem[]) => {
        if (active) {
          if (docs.length > 0) {
            setItems2(docs);
          }
          setLoading2(false);
        }
      })
      .catch(() => {
        /* keep design fallback */
        if (active) {
          setLoading2(false);
        }
      });

    return () => {
      active = false;
    };
  }, [props.context, props.panel1Library, props.panel2Library, navKey]);

  return (
    <section className={styles.root} dir={isRtl(language) ? 'rtl' : 'ltr'}>
      <SectionHeader title={props.title} linkText={props.linkText} linkUrl={props.documentHubUrl} showTitle={props.showTitle} showLink={props.showViewAll} />
      <div className={styles.grid}>
        <DocumentPanel title={props.panel1Title} libraryUrl={props.panel1Url} items={items1} newTab={newTab} loading={loading1} language={language} />
        <DocumentPanel title={props.panel2Title} libraryUrl={props.panel2Url} items={items2} newTab={newTab} loading={loading2} language={language} />
      </div>
    </section>
  );
};

export default SharedDocuments;
