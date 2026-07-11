import * as React from 'react';
import styles from './SharedDocuments.module.scss';
import { ISharedDocumentsProps } from './ISharedDocumentsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { getLibraryDocuments } from '../services/SharedDocumentsService';
import { IDocumentItem, IDocumentPanel } from '../../../common/models';
import { useSettings } from '../../../common/services/useSettings';
import { useNavKey } from '../../../common/services/useNavKey';
import { linkTarget } from '../../../common/util/format';

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

const DocumentPanel: React.FunctionComponent<IDocumentPanel & { newTab: boolean }> = (panel) => {
  const newTab: boolean = panel.newTab;
  return (
    <div className={styles.panel}>
      <div className={styles.panelHead}>
        <div className={styles.panelTitle}>{panel.title}</div>
        <a className={styles.panelLink} href={panel.libraryUrl || '#'} {...linkTarget(newTab)}>
          Open Library
        </a>
      </div>
      {panel.items.map((doc: IDocumentItem, index: number) => (
        <a className={styles.item} href={doc.url || '#'} key={index} {...linkTarget(newTab)}>
          <div className={`${styles.icon} ${iconClass(doc.fileType)}`}>
            {doc.fileType.toUpperCase()}
          </div>
          <div className={styles.name}>{doc.name}</div>
          <div className={styles.date}>{doc.modified}</div>
        </a>
      ))}
    </div>
  );
};

const SharedDocuments: React.FunctionComponent<ISharedDocumentsProps> = (props) => {
  const [items1, setItems1] = React.useState<IDocumentItem[]>(FALLBACK_PANEL1);
  const [items2, setItems2] = React.useState<IDocumentItem[]>(FALLBACK_PANEL2);
  const newTab: boolean = useSettings(props.context).openLinksInNewTab;
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;

    getLibraryDocuments(props.context, props.panel1Library)
      .then((docs: IDocumentItem[]) => {
        if (active && docs.length > 0) {
          setItems1(docs);
        }
      })
      .catch(() => {
        /* keep design fallback */
      });

    getLibraryDocuments(props.context, props.panel2Library)
      .then((docs: IDocumentItem[]) => {
        if (active && docs.length > 0) {
          setItems2(docs);
        }
      })
      .catch(() => {
        /* keep design fallback */
      });

    return () => {
      active = false;
    };
  }, [props.context, props.panel1Library, props.panel2Library, navKey]);

  return (
    <section className={styles.root}>
      <SectionHeader title="Shared Documents" linkText="Document Hub" linkUrl={props.documentHubUrl} showTitle={props.showTitle} showLink={props.showViewAll} />
      <div className={styles.grid}>
        <DocumentPanel title={props.panel1Title} libraryUrl={props.panel1Url} items={items1} newTab={newTab} />
        <DocumentPanel title={props.panel2Title} libraryUrl={props.panel2Url} items={items2} newTab={newTab} />
      </div>
    </section>
  );
};

export default SharedDocuments;
