import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { getCurrentLanguage, pickLocalized } from '../../../common/services/languageService';

export type DocFileType = 'pdf' | 'doc' | 'xls' | 'ppt' | 'other';

/** A single document, mapped from a file in the Document Center library. */
export interface IDocEntry {
  id?: number;
  title: string;
  description: string;
  category: string; // raw (English) stored value — filtering/badges localise for display
  docType: string; // raw (English) stored value
  status: string; // raw (English) stored value
  owner: string;
  reviewDate: string;
  modified: string;
  fileType: DocFileType;
  url?: string;
}

/** The document set plus the library's own URL (for the "Open Library" link). */
export interface IDocCenterResult {
  documents: IDocEntry[];
  libraryUrl?: string;
}

interface IDocFileRow {
  Id?: number;
  Title?: string;
  TitleAR?: string;
  Description?: string;
  DescriptionAR?: string;
  Category?: string;
  DocumentType?: string;
  DocStatus?: string;
  DocOwner?: string;
  ReviewDate?: string;
  Modified?: string;
  FileLeafRef?: string;
  FileRef?: string;
  FileSystemObjectType?: number;
}

/** Derives a coarse file-type badge from a file name extension. */
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

/** Strips the extension from a file name for display. */
function stripExtension(fileName: string): string {
  const idx: number = fileName.lastIndexOf('.');
  return idx > 0 ? fileName.substring(0, idx) : fileName;
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

/**
 * Reads documents (files only) from the Document Center library with their
 * metadata, resolving bilingual title/description to the visitor's language.
 * Category/type/status are returned raw (English) so filtering stays stable;
 * the component localises them for display. Returns an empty set on error.
 */
export async function getDocuments(context: WebPartContext, libraryTitle: string, pageSize: number): Promise<IDocCenterResult> {
  try {
    const rows: IDocFileRow[] = await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .items.select(
        'Id',
        'Title',
        'TitleAR',
        'Description',
        'DescriptionAR',
        'Category',
        'DocumentType',
        'DocStatus',
        'DocOwner',
        'ReviewDate',
        'Modified',
        'FileLeafRef',
        'FileRef',
        'FileSystemObjectType'
      )
      .filter('FSObjType eq 0')
      .orderBy('Modified', false)
      .top(pageSize > 0 ? pageSize : 50)();

    if (!rows || rows.length === 0) {
      const emptyUrl: string | undefined = await getLibraryUrl(context, libraryTitle);
      return { documents: [], libraryUrl: emptyUrl };
    }

    const language = getCurrentLanguage();
    const documents: IDocEntry[] = rows.map((it: IDocFileRow): IDocEntry => {
      const leaf: string = it.FileLeafRef || '';
      return {
        id: it.Id,
        title: pickLocalized(it.Title || stripExtension(leaf), it.TitleAR, language),
        description: pickLocalized(it.Description || '', it.DescriptionAR, language),
        category: it.Category || '',
        docType: it.DocumentType || '',
        status: it.DocStatus || '',
        owner: it.DocOwner || '',
        reviewDate: it.ReviewDate || '',
        modified: it.Modified || '',
        fileType: deriveFileType(leaf),
        url: it.FileRef || undefined
      };
    });

    const libraryUrl: string | undefined = await getLibraryUrl(context, libraryTitle);
    return { documents, libraryUrl };
  } catch {
    return { documents: [] };
  }
}
