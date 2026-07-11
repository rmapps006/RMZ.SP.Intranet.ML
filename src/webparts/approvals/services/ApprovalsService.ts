import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import '@pnp/sp/site-users/web';

export type ApprovalStatus = 'pending' | 'approved' | 'review';

export interface IApprovalItem {
  title: string;
  sub: string;
  status: ApprovalStatus;
  chipLabel: string;
}

export interface IApprovalsData {
  myPending: IApprovalItem[];
  awaitingMe: IApprovalItem[];
}

// Design fallback used on error or when the list returns nothing.
const FALLBACK_MY_PENDING: IApprovalItem[] = [];

const FALLBACK_AWAITING_ME: IApprovalItem[] = [];

const FALLBACK_DATA: IApprovalsData = {
  myPending: FALLBACK_MY_PENDING,
  awaitingMe: FALLBACK_AWAITING_ME
};

interface IRawRequestItem {
  Title?: string;
  Description?: string;
  Status?: string;
  RequestedBy?: { Title?: string; Email?: string };
  AuthorEmail?: string;
  ApproverEmail?: string;
  Department?: string;
}

interface ICurrentUser {
  Email?: string;
  LoginName?: string;
}

function normaliseStatus(raw: string | undefined): ApprovalStatus {
  const v: string = (raw || '').toLowerCase();
  if (v.indexOf('approve') !== -1) {
    return 'approved';
  }
  if (v.indexOf('review') !== -1) {
    return 'review';
  }
  return 'pending';
}

function chipFor(status: ApprovalStatus, awaiting: boolean): string {
  if (status === 'approved') {
    return 'Approved';
  }
  if (status === 'review') {
    return 'In Review';
  }
  return awaiting ? 'Action Needed' : 'Pending';
}

function mapItem(item: IRawRequestItem, awaiting: boolean): IApprovalItem {
  const status: ApprovalStatus = normaliseStatus(item.Status);
  return {
    title: item.Title || '',
    sub: item.Description || '',
    status: status,
    chipLabel: chipFor(status, awaiting)
  };
}

/**
 * Reads the requests/approvals list for the current user and splits the rows
 * into "my pending requests" and "awaiting my approval". On error/empty
 * returns the design fallback data.
 */
export async function getApprovals(
  context: WebPartContext,
  requestsList: string,
  department: string
): Promise<IApprovalsData> {
  try {
    const sp = getSP(context);

    const user: ICurrentUser = await sp.web.currentUser();
    const email: string = (user.Email || '').replace(/'/g, "''");
    const deptFilter: string = `Department eq '${department.replace(/'/g, "''")}'`;

    const list = sp.web.lists.getByTitle(requestsList);
    const select: string[] = [
      'Title',
      'Description',
      'Status',
      'AuthorEmail',
      'ApproverEmail',
      'Department'
    ];

    const mineRaw: IRawRequestItem[] = await list.items
      .select(...select)
      .filter(`AuthorEmail eq '${email}' and ${deptFilter}`)();

    const awaitingRaw: IRawRequestItem[] = await list.items
      .select(...select)
      .filter(`ApproverEmail eq '${email}' and ${deptFilter}`)();

    const myPending: IApprovalItem[] = (mineRaw || []).map((i: IRawRequestItem) => mapItem(i, false));
    const awaitingMe: IApprovalItem[] = (awaitingRaw || []).map((i: IRawRequestItem) => mapItem(i, true));

    if (myPending.length === 0 && awaitingMe.length === 0) {
      return FALLBACK_DATA;
    }

    return {
      myPending: myPending.length > 0 ? myPending : FALLBACK_MY_PENDING,
      awaitingMe: awaitingMe.length > 0 ? awaitingMe : FALLBACK_AWAITING_ME
    };
  } catch {
    return FALLBACK_DATA;
  }
}
