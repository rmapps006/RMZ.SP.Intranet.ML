import * as React from 'react';
import styles from './DocumentCenter.module.scss';
import { IDocumentCenterProps } from './IDocumentCenterProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { getDocuments, computeStats, IDocCenterResult, IDocEntry, IDocStats, DocRole } from '../services/DocumentCenterService';
import { useNavKey } from '../../../common/services/useNavKey';
import { getCurrentLanguage, isRtl, Language } from '../../../common/services/languageService';
import { t } from '../../../common/services/uiStrings';
import { IDocCtx } from './IDocTabProps';
import { Pivot, PivotItem } from '@fluentui/react';
import DashboardTab from './tabs/DashboardTab';
import DocumentsTab from './tabs/DocumentsTab';
import UploadTab from './tabs/UploadTab';
import MyDocumentsTab from './tabs/MyDocumentsTab';
import ReportsTab from './tabs/ReportsTab';
import AdminTab from './tabs/AdminTab';

const EMPTY_STATS: IDocStats = {
  total: 0,
  byStatus: {},
  byDepartment: {},
  byType: {},
  bySensitivity: {},
  overdue: 0,
  myCount: 0,
  draft: 0,
  inReview: 0,
  published: 0
};

function roleLabel(role: DocRole, language: Language): string {
  switch (role) {
    case 'Administrator':
      return t('roleAdministrator', language);
    case 'Approver':
      return t('roleApprover', language);
    case 'Contributor':
      return t('roleContributor', language);
    default:
      return t('roleReader', language);
  }
}

const DocumentCenter: React.FunctionComponent<IDocumentCenterProps> = (props) => {
  const [docs, setDocs] = React.useState<IDocEntry[]>([]);
  const [libraryUrl, setLibraryUrl] = React.useState<string | undefined>(undefined);
  const [webUrl, setWebUrl] = React.useState<string>(props.context.pageContext.web.absoluteUrl);
  const [listId, setListId] = React.useState<string>('');
  const [role, setRole] = React.useState<DocRole>('Reader');
  const [currentUserId, setCurrentUserId] = React.useState<number>(0);
  const [loadError, setLoadError] = React.useState<string | undefined>(undefined);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [reloadKey, setReloadKey] = React.useState<number>(0);

  const language: Language = getCurrentLanguage();
  const rtl: boolean = isRtl(language);
  const navKey: string = useNavKey();

  const reload = React.useCallback((): void => setReloadKey((k) => k + 1), []);

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    getDocuments(props.context, props.libraryTitle, props.pageSize, props.siteUrl)
      .then((result: IDocCenterResult) => {
        if (!active) {
          return;
        }
        setDocs(result.documents);
        setLibraryUrl(result.libraryUrl);
        setWebUrl(result.webUrl);
        setListId(result.listId);
        setRole(result.role);
        setCurrentUserId(result.currentUserId);
        setLoadError(result.loadError);
      })
      .catch(() => {
        /* service returns a safe default */
      })
      .finally(() => {
        if (active) {
          setLoading(false);
        }
      });
    return () => {
      active = false;
    };
  }, [props.context, props.libraryTitle, props.siteUrl, props.pageSize, navKey, reloadKey]);

  const stats: IDocStats = React.useMemo(() => (docs.length ? computeStats(docs, currentUserId) : EMPTY_STATS), [docs, currentUserId]);

  const ctx: IDocCtx = {
    context: props.context,
    libraryTitle: props.libraryTitle,
    siteUrl: props.siteUrl,
    language,
    rtl,
    role,
    docs,
    stats,
    webUrl,
    listId,
    currentUserId,
    libraryUrl,
    loadError,
    reload
  };

  const canContribute: boolean = role === 'Contributor' || role === 'Approver' || role === 'Administrator';
  const isAdmin: boolean = role === 'Administrator';

  const tabs: JSX.Element[] = [];
  tabs.push(
    <PivotItem key="dashboard" headerText={t('tabDashboard', language)}>
      <DashboardTab ctx={ctx} />
    </PivotItem>
  );
  tabs.push(
    <PivotItem key="documents" headerText={t('tabDocuments', language)}>
      <DocumentsTab ctx={ctx} />
    </PivotItem>
  );
  if (canContribute) {
    tabs.push(
      <PivotItem key="upload" headerText={t('tabUpload', language)}>
        <UploadTab ctx={ctx} />
      </PivotItem>
    );
    tabs.push(
      <PivotItem key="mydocs" headerText={t('tabMyDocuments', language)}>
        <MyDocumentsTab ctx={ctx} />
      </PivotItem>
    );
  }
  tabs.push(
    <PivotItem key="reports" headerText={t('tabReports', language)}>
      <ReportsTab ctx={ctx} />
    </PivotItem>
  );
  if (isAdmin) {
    tabs.push(
      <PivotItem key="admin" headerText={t('tabAdmin', language)}>
        <AdminTab ctx={ctx} />
      </PivotItem>
    );
  }

  return (
    <section className={styles.dc} dir={rtl ? 'rtl' : 'ltr'}>
      <SectionHeader
        title={props.title}
        linkText={props.linkText}
        linkUrl={libraryUrl || ''}
        showTitle={props.showTitle}
        showLink={props.showViewAll && !!libraryUrl}
      />

      <div className={styles.roleBar}>
        <span className={styles.roleChip}>{roleLabel(role, language)}</span>
      </div>

      {loading ? <div className={styles.empty}>{t('loading', language)}</div> : <Pivot className={styles.pivot}>{tabs}</Pivot>}
    </section>
  );
};

export default DocumentCenter;
