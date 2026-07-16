import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPPermission } from '@microsoft/sp-page-context';
import { getSP } from '../../../common/services/pnpService';
import { getCurrentLanguage, pickLocalized } from '../../../common/services/languageService';
import '@pnp/sp/security';
import '@pnp/sp/files';
import '@pnp/sp/folders';
import '@pnp/sp/fields';

export type DocFileType = 'pdf' | 'doc' | 'xls' | 'ppt' | 'other';
export type DocStage = 'Draft' | 'In Review' | 'Published' | 'Archived';
export type DocRole = 'Administrator' | 'Approver' | 'Contributor' | 'Reader';

/** A single managed document. */
export interface IDocEntry {
  id: number;
  title: string;
  description: string;
  documentNumber: string;
  department: string; // raw (English) stored value — stored in the Category column
  docType: string; // raw (English)
  status: string; // raw (English) stage/status
  sensitivity: string; // raw (English)
  tags: string[];
  owner: string;
  reviewDate: string;
  modified: string;
  modifiedBy: string;
  authorId: number;
  fileType: DocFileType;
  url?: string;
}

/** Aggregate counts for the dashboard and reports. */
export interface IDocStats {
  total: number;
  byStatus: { [key: string]: number };
  byDepartment: { [key: string]: number };
  byType: { [key: string]: number };
  bySensitivity: { [key: string]: number };
  overdue: number;
  myCount: number;
  draft: number;
  inReview: number;
  published: number;
}

/** Everything the web part needs for one render: docs, library context and the viewer's role. */
export interface IDocCenterResult {
  documents: IDocEntry[];
  libraryUrl?: string;
  webUrl: string;
  listId: string;
  role: DocRole;
  currentUserId: number;
}

interface IDocFileRow {
  Id?: number;
  Title?: string;
  TitleAR?: string;
  DocDescription?: string;
  DescriptionAR?: string;
  DocumentNumber?: string;
  Category?: string;
  DocumentType?: string;
  DocStatus?: string;
  DocSensitivity?: string;
  DocTags?: string;
  DocOwner?: string;
  ReviewDate?: string;
  Modified?: string;
  Author?: { Id?: number; Title?: string };
  Editor?: { Title?: string };
  FileLeafRef?: string;
  FileRef?: string;
  FileSystemObjectType?: number;
}

/** Metadata supplied by the Upload form (department/type/etc. are raw English stored values). */
export interface IUploadMeta {
  title: string;
  titleAR: string;
  department: string;
  docType: string;
  sensitivity: string;
  status: string;
  tags: string;
  owner: string;
  reviewDate: string;
  description: string;
}

/** Short codes used to build a document number per department. */
const DEPT_CODE: { [key: string]: string } = {
  HR: 'HR',
  Finance: 'FIN',
  IT: 'IT',
  Operations: 'OPS',
  Marketing: 'MKT',
  Legal: 'LEG',
  General: 'GEN'
};

function deriveFileType(fileName: string | undefined): DocFileType {
  const v: string = (fileName || '').toLowerCase();
  if (v.indexOf('.xls') !== -1 || v.indexOf('.csv') !== -1) {
    return 'xls';
  }
  if (v.indexOf('.ppt') !== -1) {
    return 'ppt';
  }
  if (v.indexOf('.doc') !== -1) {
    return 'doc';
  }
  if (v.indexOf('.pdf') !== -1) {
    return 'pdf';
  }
  return 'other';
}

function stripExtension(fileName: string): string {
  const idx: number = fileName.lastIndexOf('.');
  return idx > 0 ? fileName.substring(0, idx) : fileName;
}

function splitTags(value: string | undefined): string[] {
  return (value || '')
    .split(/[;,]/)
    .map((t) => t.trim())
    .filter((t) => !!t);
}

/**
 * Resolves the viewer's DMS role from their web-level base permissions. Uses the
 * synchronous page context (no API call, always available) so role detection is
 * reliable: Manage Web → Administrator, Approve Items → Approver, Add List Items
 * → Contributor, otherwise Reader.
 */
export function getUserRole(context: WebPartContext): DocRole {
  try {
    const perms = context.pageContext.web.permissions;
    if (perms.hasPermission(SPPermission.manageWeb)) {
      return 'Administrator';
    }
    if (perms.hasPermission(SPPermission.approveItems)) {
      return 'Approver';
    }
    if (perms.hasPermission(SPPermission.addListItems) || perms.hasPermission(SPPermission.editListItems)) {
      return 'Contributor';
    }
    return 'Reader';
  } catch {
    return 'Reader';
  }
}

/** Returns the library's server-relative URL (its root folder) for the "Open Library" link. */
async function getLibraryUrl(context: WebPartContext, libraryTitle: string): Promise<string | undefined> {
  try {
    const info: { ServerRelativeUrl?: string } = await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .rootFolder.select('ServerRelativeUrl')();
    return info && info.ServerRelativeUrl ? info.ServerRelativeUrl : undefined;
  } catch {
    return undefined;
  }
}

function mapRows(rows: IDocFileRow[], language: 'en' | 'ar'): IDocEntry[] {
  return rows.map((it: IDocFileRow): IDocEntry => {
    const leaf: string = it.FileLeafRef || '';
    return {
      id: it.Id || 0,
      title: pickLocalized(it.Title || stripExtension(leaf), it.TitleAR, language),
      description: pickLocalized(it.DocDescription || '', it.DescriptionAR, language),
      documentNumber: it.DocumentNumber || '',
      department: it.Category || '',
      docType: it.DocumentType || '',
      status: it.DocStatus || '',
      sensitivity: it.DocSensitivity || '',
      tags: splitTags(it.DocTags),
      owner: it.DocOwner || (it.Author && it.Author.Title) || '',
      reviewDate: it.ReviewDate || '',
      modified: it.Modified || '',
      modifiedBy: (it.Editor && it.Editor.Title) || '',
      authorId: (it.Author && it.Author.Id) || 0,
      fileType: deriveFileType(leaf),
      url: it.FileRef || undefined
    };
  });
}

/**
 * Loads the full working set: documents (bilingual title/description resolved),
 * the library URL and web/list identifiers (for version-history links), the
 * viewer's role and their user id (for "My Documents"). The metadata select is
 * resilient — if the DMS columns aren't present yet, it falls back to built-in
 * fields so documents still list.
 */
export async function getDocuments(context: WebPartContext, libraryTitle: string, pageSize: number): Promise<IDocCenterResult> {
  const sp = getSP(context);
  const webUrl: string = context.pageContext.web.absoluteUrl;
  const role: DocRole = getUserRole(context);

  const [currentUser, listInfo, libraryUrl] = await Promise.all([
    sp.web.currentUser.select('Id')().catch(() => ({ Id: 0 })),
    sp.web.lists
      .getByTitle(libraryTitle)
      .select('Id')()
      .catch(() => ({ Id: '' })),
    getLibraryUrl(context, libraryTitle)
  ]);

  const base: IDocCenterResult = {
    documents: [],
    libraryUrl,
    webUrl,
    listId: (listInfo as { Id?: string }).Id || '',
    role,
    currentUserId: (currentUser as { Id?: number }).Id || 0
  };

  const list = sp.web.lists.getByTitle(libraryTitle);
  const top: number = pageSize > 0 ? pageSize : 200;
  const CORE: string[] = ['Id', 'Title', 'Modified', 'FileLeafRef', 'FileRef', 'FileSystemObjectType'];
  const WITH_PEOPLE: string[] = [...CORE, 'Author/Id', 'Author/Title', 'Editor/Title'];
  const FULL: string[] = [
    ...WITH_PEOPLE,
    'TitleAR',
    'DocDescription',
    'DescriptionAR',
    'DocumentNumber',
    'Category',
    'DocumentType',
    'DocStatus',
    'DocSensitivity',
    'DocTags',
    'DocOwner',
    'ReviewDate'
  ];
  const language = getCurrentLanguage();

  // Progressive fallback: full metadata + people, then people-only, then the
  // guaranteed built-in fields with no expand. This ensures documents list even
  // when the DMS columns don't exist yet or an expand isn't permitted.
  const attempts: (() => Promise<IDocFileRow[]>)[] = [
    () => list.items.select(...FULL).expand('Author', 'Editor').filter('FSObjType eq 0').orderBy('Modified', false).top(top)(),
    () => list.items.select(...WITH_PEOPLE).expand('Author', 'Editor').filter('FSObjType eq 0').orderBy('Modified', false).top(top)(),
    () => list.items.select(...CORE).filter('FSObjType eq 0').top(top)()
  ];

  for (const attempt of attempts) {
    try {
      const rows: IDocFileRow[] = await attempt();
      base.documents = mapRows(rows || [], language);
      return base;
    } catch {
      /* try the next, less-demanding query shape */
    }
  }
  return base;
}

/** Computes dashboard/report aggregates from a document set. */
export function computeStats(docs: IDocEntry[], currentUserId: number): IDocStats {
  const stats: IDocStats = {
    total: docs.length,
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
  const now: number = new Date().getTime();
  docs.forEach((d) => {
    const bump = (map: { [key: string]: number }, key: string): void => {
      const k: string = key || '—';
      map[k] = (map[k] || 0) + 1;
    };
    bump(stats.byStatus, d.status);
    bump(stats.byDepartment, d.department);
    bump(stats.byType, d.docType);
    bump(stats.bySensitivity, d.sensitivity);
    if (d.status === 'Draft') {
      stats.draft += 1;
    } else if (d.status === 'In Review') {
      stats.inReview += 1;
    } else if (d.status === 'Published') {
      stats.published += 1;
    }
    if (d.authorId && d.authorId === currentUserId) {
      stats.myCount += 1;
    }
    if (d.reviewDate && d.status !== 'Archived') {
      const rd: number = new Date(d.reviewDate).getTime();
      if (!isNaN(rd) && rd < now) {
        stats.overdue += 1;
      }
    }
  });
  return stats;
}

/** True when a review date is in the past (and the doc isn't archived). */
export function isOverdue(doc: IDocEntry): boolean {
  if (!doc.reviewDate || doc.status === 'Archived') {
    return false;
  }
  const rd: number = new Date(doc.reviewDate).getTime();
  return !isNaN(rd) && rd < new Date().getTime();
}

/** Builds the next document number for a department, e.g. HR-2026-0007. */
async function nextDocumentNumber(context: WebPartContext, libraryTitle: string, department: string): Promise<string> {
  const code: string = DEPT_CODE[department] || (department ? department.substring(0, 3).toUpperCase() : 'DOC');
  const year: number = new Date().getFullYear();
  const prefix: string = `${code}-${year}-`;
  let max: number = 0;
  try {
    const rows: { DocumentNumber?: string }[] = await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .items.select('DocumentNumber')
      .filter(`startswith(DocumentNumber,'${prefix}')`)
      .top(500)();
    rows.forEach((r) => {
      const parts: string[] = (r.DocumentNumber || '').split('-');
      const seq: number = parseInt(parts[parts.length - 1], 10);
      if (!isNaN(seq) && seq > max) {
        max = seq;
      }
    });
  } catch {
    /* fall back to 1 */
  }
  return `${prefix}${String(max + 1).padStart(4, '0')}`;
}

/** Uploads a file and sets its metadata, assigning an automatic document number. */
export async function uploadDocument(
  context: WebPartContext,
  libraryTitle: string,
  file: File,
  meta: IUploadMeta
): Promise<{ ok: boolean; documentNumber?: string; error?: string }> {
  try {
    const sp = getSP(context);
    const list = sp.web.lists.getByTitle(libraryTitle);
    const documentNumber: string = await nextDocumentNumber(context, libraryTitle, meta.department);
    const info: { ServerRelativeUrl?: string } = await list.rootFolder.files.addUsingPath(file.name, file, { Overwrite: false });
    const item = await sp.web.getFileByServerRelativePath(info.ServerRelativeUrl || '').getItem();
    const fields: Record<string, unknown> = {
      Title: meta.title || stripExtension(file.name),
      TitleAR: meta.titleAR,
      DocumentNumber: documentNumber,
      Category: meta.department,
      DocumentType: meta.docType,
      DocStatus: meta.status || 'Draft',
      DocSensitivity: meta.sensitivity,
      DocTags: meta.tags,
      DocOwner: meta.owner,
      ReviewDate: meta.reviewDate || null,
      DocDescription: meta.description
    };
    try {
      await item.update(fields);
    } catch {
      for (const key of Object.keys(fields)) {
        try {
          await item.update({ [key]: fields[key] });
        } catch {
          /* skip a field that doesn't accept the value */
        }
      }
    }
    return { ok: true, documentNumber };
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) };
  }
}

/** Moves a document to a new stage (Draft / In Review / Published / Archived). */
export async function setStage(context: WebPartContext, libraryTitle: string, itemId: number, stage: DocStage): Promise<boolean> {
  try {
    await getSP(context).web.lists.getByTitle(libraryTitle).items.getById(itemId).update({ DocStatus: stage });
    return true;
  } catch {
    return false;
  }
}

/** Server-relative link to a document's native SharePoint version history. */
export function versionHistoryUrl(webUrl: string, listId: string, itemId: number): string {
  return `${webUrl}/_layouts/15/Versions.aspx?list=${encodeURIComponent(listId)}&ID=${itemId}`;
}

/** Reads the current choices for a choice column (Admin taxonomy editor). */
export async function getChoices(context: WebPartContext, libraryTitle: string, fieldInternalName: string): Promise<string[]> {
  try {
    const field: { Choices?: string[] } = await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .fields.getByInternalNameOrTitle(fieldInternalName)
      .select('Choices')();
    return field && field.Choices ? field.Choices : [];
  } catch {
    return [];
  }
}

/** Replaces the choices of a choice column (Admin taxonomy editor, Administrators only). */
export async function updateChoices(context: WebPartContext, libraryTitle: string, fieldInternalName: string, choices: string[]): Promise<boolean> {
  try {
    await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .fields.getByInternalNameOrTitle(fieldInternalName)
      .update({ Choices: choices }, 'SP.FieldChoice');
    return true;
  } catch {
    return false;
  }
}
