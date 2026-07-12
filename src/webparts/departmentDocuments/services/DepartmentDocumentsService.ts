import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { formatLongDate, browserViewUrl } from '../../../common/util/format';

export type DocFileType = 'pdf' | 'doc' | 'xls';

/** A single document row in a panel. */
export interface IDeptDocItem {
  name: string;
  meta: string; // "Category · Date"
  fileType: DocFileType;
  url?: string; // opens the file in the browser (web=1) rather than downloading
  badge?: string; // e.g. "New" or "Updated" (shown on the right of the row)
}

/** A document panel (Department Policies / Shared Documents). */
export interface IDeptDocPanel {
  title: string;
  items: IDeptDocItem[];
  libraryUrl?: string; // server-relative URL of the library — the "Open Library" link target
}

// Design fallbacks (from design/department-page.html) — used on error/empty.
const FALLBACK_POLICIES: IDeptDocItem[] = [];

const FALLBACK_DOCUMENTS: IDeptDocItem[] = [];

interface IRawFileItem {
  Title?: string;
  Category?: string;
  Department?: string;
  Modified?: string;
  File?: { Name?: string; TimeLastModified?: string; ServerRelativeUrl?: string };
}

/** Derive a file-type badge from a file name; defaults to doc. */
function deriveFileType(fileName: string | undefined): DocFileType {
  const lower: string = (fileName || '').toLowerCase();
  if (lower.indexOf('.pdf') >= 0) {
    return 'pdf';
  }
  if (lower.indexOf('.xls') >= 0 || lower.indexOf('.csv') >= 0) {
    return 'xls';
  }
  return 'doc';
}

function buildMeta(category: string | undefined, modified: string | undefined): string {
  const cat: string = (category || '').trim();
  const date: string = formatLongDate(modified);
  if (cat && date) {
    return `${cat} · ${date}`;
  }
  return cat || date;
}

async function readLibrary(
  context: WebPartContext,
  libraryTitle: string,
  department: string,
  firstBadge: string
): Promise<IDeptDocItem[]> {
  const sp = getSP(context);
  let query = sp.web.lists
    .getByTitle(libraryTitle)
    .items.select('Title', 'Category', 'Department', 'Modified', 'File/Name', 'File/TimeLastModified', 'File/ServerRelativeUrl')
    .expand('File');

  if (department) {
    query = query.filter(`Department eq '${department.replace(/'/g, "''")}'`);
  }

  const items: IRawFileItem[] = await query.orderBy('Modified', false).top(5)();

  return items.map((item: IRawFileItem, index: number): IDeptDocItem => {
    const fileName: string = (item.File && item.File.Name) || item.Title || '';
    return {
      name: item.Title || fileName,
      meta: buildMeta(item.Category, (item.File && item.File.TimeLastModified) || item.Modified),
      fileType: deriveFileType(fileName),
      url: browserViewUrl(item.File && item.File.ServerRelativeUrl),
      badge: index === 0 ? firstBadge : undefined
    };
  });
}

/**
 * Returns a library's server-relative URL (its root folder) so the panel's
 * "Open Library" link can point at the real library. Undefined when the library
 * doesn't exist yet or can't be read.
 */
async function getLibraryUrl(context: WebPartContext, libraryTitle: string): Promise<string | undefined> {
  if (!libraryTitle) {
    return undefined;
  }
  try {
    const info: { ServerRelativeUrl?: string } = await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .rootFolder.select('ServerRelativeUrl')();
    return info && info.ServerRelativeUrl ? info.ServerRelativeUrl : undefined;
  } catch {
    return undefined;
  }
}

/**
 * Reads the configured Policies and Documents libraries (filtered by department
 * where applicable) and maps to two typed panels, each carrying the library's
 * URL for its "Open Library" link. On error/empty per library, the design
 * fallbacks are returned.
 */
export async function getDepartmentDocuments(
  context: WebPartContext,
  policiesLibrary: string,
  documentsLibrary: string,
  department: string,
  panel1Title: string,
  panel2Title: string
): Promise<IDeptDocPanel[]> {
  let policies: IDeptDocItem[];
  try {
    const result: IDeptDocItem[] = await readLibrary(context, policiesLibrary, department, 'New');
    policies = result && result.length > 0 ? result : FALLBACK_POLICIES;
  } catch {
    policies = FALLBACK_POLICIES;
  }

  let documents: IDeptDocItem[];
  try {
    const result: IDeptDocItem[] = await readLibrary(context, documentsLibrary, department, 'Updated');
    documents = result && result.length > 0 ? result : FALLBACK_DOCUMENTS;
  } catch {
    documents = FALLBACK_DOCUMENTS;
  }

  const [policiesUrl, documentsUrl] = await Promise.all([
    getLibraryUrl(context, policiesLibrary),
    getLibraryUrl(context, documentsLibrary)
  ]);

  return [
    { title: panel1Title, items: policies, libraryUrl: policiesUrl },
    { title: panel2Title, items: documents, libraryUrl: documentsUrl }
  ];
}
