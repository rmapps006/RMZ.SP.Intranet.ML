import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { getGraphClient } from '../../../common/services/graphService';
import { imageFieldUrl } from '../../../common/util/format';

export type DetailType = 'news' | 'event' | 'policy' | 'employee' | 'benefit';

export interface IDetailParams {
  type: DetailType | '';
  id: string;
}

export interface INewsDetail {
  title: string;
  category: string;
  date: string;
  source: string;
  linkUrl?: string;
  imageUrl?: string;
  body?: string;
}

export interface IEventDetail {
  title: string;
  category: string;
  start?: string;
  end?: string;
  location: string;
  description?: string;
}

export interface IPolicyDetail {
  title: string;
  department: string;
  version: string;
  fileType: string;
  modified?: string;
  documentUrl?: string; // the policy file to preview/open, if attached or linked
}

export interface IEmployeeDetail {
  displayName: string;
  jobTitle?: string;
  department?: string;
  email?: string;
  phone?: string;
  mobile?: string;
  office?: string;
}

export interface IBenefitDetail {
  title: string;
  category: string;
  summary?: string;
  eligibility?: string;
  coverage?: string;
  details?: string;
}

/**
 * Reads the ?type= and ?recId= parameters from the current page URL. `recId` is
 * used instead of `id` because SharePoint reserves the `id` query parameter
 * (older links using `id` are still accepted as a fallback).
 */
export function readDetailParams(): IDetailParams {
  try {
    const params: URLSearchParams = new URLSearchParams(window.location.search);
    const type: string = (params.get('type') || '').toLowerCase();
    const id: string = params.get('recId') || params.get('id') || '';
    const valid: boolean =
      type === 'news' || type === 'event' || type === 'policy' || type === 'employee' || type === 'benefit';
    return { type: valid ? (type as DetailType) : '', id };
  } catch {
    return { type: '', id: '' };
  }
}

interface IRawNews {
  Title?: string;
  Category?: string;
  NewsDate?: string;
  Source?: string;
  LinkUrl?: string;
  ImageUrl?: string;
  Body?: string;
  Created?: string;
}

export async function getNewsDetail(context: WebPartContext, listTitle: string, id: number): Promise<INewsDetail | undefined> {
  try {
    const listItem = getSP(context).web.lists.getByTitle(listTitle).items.getById(id);
    // Normal fields via the items endpoint; tolerate a legacy text ImageUrl.
    const base: string[] = ['Title', 'Category', 'NewsDate', 'Source', 'LinkUrl', 'Body', 'Created'];
    const attempts: string[][] = [[...base, 'ImageUrl'], base];
    let item: IRawNews | undefined;
    for (let i = 0; i < attempts.length; i++) {
      try {
        item = await listItem.select(...attempts[i])();
        break;
      } catch (e) {
        if (i === attempts.length - 1) {
          throw e;
        }
      }
    }
    if (!item) {
      return undefined;
    }
    return {
      title: item.Title || '',
      category: item.Category || '',
      date: item.NewsDate || item.Created || '',
      source: item.Source || '',
      linkUrl: item.LinkUrl || undefined,
      imageUrl: imageFieldUrl(item.ImageUrl),
      body: item.Body || undefined
    };
  } catch {
    return undefined;
  }
}

interface IRawEvent {
  Title?: string;
  Category?: string;
  EventDate?: string;
  EndDate?: string;
  Location?: string;
  Description?: string;
}

export async function getEventDetail(context: WebPartContext, listTitle: string, id: number): Promise<IEventDetail | undefined> {
  try {
    const item: IRawEvent = await getSP(context)
      .web.lists.getByTitle(listTitle)
      .items.getById(id)
      .select('Title', 'Category', 'EventDate', 'EndDate', 'Location', 'Description')();
    return {
      title: item.Title || '',
      category: item.Category || '',
      start: item.EventDate,
      end: item.EndDate,
      location: item.Location || '',
      description: item.Description || undefined
    };
  } catch {
    return undefined;
  }
}

interface IRawPolicy {
  Title?: string;
  Department?: string;
  PolicyVersion?: string;
  Modified?: string;
  FileLeafRef?: string;
  FileRef?: string;
  FileSystemObjectType?: number;
  DocumentUrl?: string;
  AttachmentFiles?: Array<{ ServerRelativeUrl?: string }>;
}

function policyFileType(fileName: string | undefined): string {
  const v: string = (fileName || '').toLowerCase();
  if (v.indexOf('.xls') !== -1 || v.indexOf('.csv') !== -1) {
    return 'xls';
  }
  if (v.indexOf('.doc') !== -1 || v.indexOf('.ppt') !== -1) {
    return 'doc';
  }
  if (v.indexOf('.pdf') !== -1) {
    return 'pdf';
  }
  return '';
}

export async function getPolicyDetail(context: WebPartContext, listTitle: string, id: number): Promise<IPolicyDetail | undefined> {
  // Policies is a document library, so the item is the file itself (FileRef).
  // Still tolerate a legacy custom list that used a DocumentUrl field or a list
  // attachment — try the richest shape first and drop fields that don't exist.
  const fileFields: string[] = ['Title', 'Department', 'PolicyVersion', 'Modified', 'FileLeafRef', 'FileRef', 'FileSystemObjectType'];
  const attempts: Array<{ select: string[]; expand: string[] }> = [
    // Document-library case first (the current shape) — the file is FileRef.
    { select: fileFields, expand: [] },
    // Legacy custom list that used a DocumentUrl field or a list attachment.
    { select: [...fileFields, 'DocumentUrl', 'AttachmentFiles/ServerRelativeUrl'], expand: ['AttachmentFiles'] },
    { select: ['Title', 'Department', 'PolicyVersion', 'Modified'], expand: [] }
  ];
  for (let i = 0; i < attempts.length; i++) {
    try {
      const q = getSP(context).web.lists.getByTitle(listTitle).items.getById(id).select(...attempts[i].select);
      const item: IRawPolicy = await (attempts[i].expand.length > 0 ? q.expand(...attempts[i].expand)() : q());
      const attachment: string | undefined =
        item.AttachmentFiles && item.AttachmentFiles.length > 0 ? item.AttachmentFiles[0].ServerRelativeUrl : undefined;
      const leaf: string = item.FileLeafRef || '';
      const title: string = item.Title || (leaf ? leaf.replace(/\.[^.]+$/, '') : '');
      return {
        title: title,
        department: item.Department || '',
        version: item.PolicyVersion || '',
        fileType: policyFileType(leaf),
        modified: item.Modified,
        documentUrl: (item.DocumentUrl || '').trim() || item.FileRef || attachment || undefined
      };
    } catch {
      if (i === attempts.length - 1) {
        return undefined;
      }
    }
  }
  return undefined;
}

interface IRawBenefit {
  Title?: string;
  Category?: string;
  Summary?: string;
  Eligibility?: string;
  Coverage?: string;
  Details?: string;
}

export async function getBenefitDetail(context: WebPartContext, listTitle: string, id: number): Promise<IBenefitDetail | undefined> {
  try {
    const item: IRawBenefit = await getSP(context)
      .web.lists.getByTitle(listTitle)
      .items.getById(id)
      .select('Title', 'Category', 'Summary', 'Eligibility', 'Coverage', 'Details')();
    return {
      title: item.Title || '',
      category: item.Category || '',
      summary: item.Summary || undefined,
      eligibility: item.Eligibility || undefined,
      coverage: item.Coverage || undefined,
      details: item.Details || undefined
    };
  } catch {
    return undefined;
  }
}

interface IGraphUser {
  displayName?: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  officeLocation?: string;
}

/** id may be a user object id, UPN or email. */
export async function getEmployeeDetail(context: WebPartContext, id: string): Promise<IEmployeeDetail | undefined> {
  try {
    const client = await getGraphClient(context);
    const u: IGraphUser = await client
      .api(`/users/${encodeURIComponent(id)}`)
      .select('displayName,jobTitle,department,mail,userPrincipalName,businessPhones,mobilePhone,officeLocation')
      .get();
    const phones: string[] = u.businessPhones || [];
    return {
      displayName: u.displayName || '',
      jobTitle: u.jobTitle,
      department: u.department,
      email: u.mail || u.userPrincipalName,
      phone: phones.length > 0 ? phones[0] : undefined,
      mobile: u.mobilePhone,
      office: u.officeLocation
    };
  } catch {
    return undefined;
  }
}
