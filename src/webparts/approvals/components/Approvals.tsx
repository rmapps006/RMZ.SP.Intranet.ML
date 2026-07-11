import * as React from 'react';
import styles from './Approvals.module.scss';
import { IApprovalsProps } from './IApprovalsProps';
import { SectionHeader } from '../../../common/components/SectionHeader';
import { getApprovals, IApprovalsData, IApprovalItem, ApprovalStatus } from '../services/ApprovalsService';
import { useNavKey } from '../../../common/services/useNavKey';

const EMPTY_DATA: IApprovalsData = { myPending: [], awaitingMe: [] };

const Approvals: React.FunctionComponent<IApprovalsProps> = (props) => {
  const [data, setData] = React.useState<IApprovalsData>(EMPTY_DATA);
  const [loading, setLoading] = React.useState<boolean>(true);
  const navKey: string = useNavKey();

  React.useEffect(() => {
    let active: boolean = true;
    setLoading(true);
    getApprovals(props.context, props.requestsList, props.department)
      .then((result: IApprovalsData) => {
        if (active) {
          setData(result);
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
  }, [props.context, props.requestsList, props.department, navKey]);

  const statusClass = (status: ApprovalStatus): string => {
    if (status === 'approved') {
      return styles.approved;
    }
    if (status === 'review') {
      return styles.review;
    }
    return styles.pending;
  };

  const renderPanel = (
    panelTitle: string,
    linkText: string,
    linkUrl: string,
    items: IApprovalItem[],
    emptyText: string
  ): React.ReactElement => {
    return (
      <div className={styles.appPanel}>
        <div className={styles.appPanelHd}>
          <div className={styles.appPanelTt}>{panelTitle}</div>
          <a className={styles.appPanelLnk} href={linkUrl || '#'}>
            {linkText}
          </a>
        </div>
        {loading ? (
          <div className={styles.appEmpty}>Loading…</div>
        ) : items.length === 0 ? (
          <div className={styles.appEmpty}>{emptyText}</div>
        ) : (
          items.map((item: IApprovalItem, index: number) => (
            <div key={`${item.title}-${index}`} className={styles.appItem}>
              <div className={`${styles.appStatus} ${statusClass(item.status)}`} />
              <div className={styles.appInfo}>
                <div className={styles.appTitle}>{item.title}</div>
                <div className={styles.appSub}>{item.sub}</div>
              </div>
              <span className={`${styles.appChip} ${statusClass(item.status)}`}>{item.chipLabel}</span>
            </div>
          ))
        )}
      </div>
    );
  };

  return (
    <section className={styles.approvals}>
      <SectionHeader title="Applications & Approvals" linkText="View All" linkUrl={props.viewAllUrl} showTitle={props.showTitle} showLink={props.showViewAll} />

      <div className={styles.appsG}>
        {renderPanel('My Pending Requests', 'New Request', props.newRequestUrl, data.myPending, 'No pending requests.')}
        {renderPanel('Awaiting My Approval', 'Review All', props.reviewAllUrl, data.awaitingMe, 'Nothing awaiting your approval.')}
      </div>
    </section>
  );
};

export default Approvals;
