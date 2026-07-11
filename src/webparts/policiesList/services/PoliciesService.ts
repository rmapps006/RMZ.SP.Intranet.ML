import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { IPolicy } from '../../../common/models';

// Raw shape returned from the Policies document library.
interface IPolicyFileRow {
  Id?: number;
  Title?: string;
  Department?: string;
  Modified?: string;
  PolicyVersion?: string;
  FileLeafRef?: string;
  FileRef?: string;
  FileSystemObjectType?: number;
}

/** Derives the badge type from a file name extension; defaults to pdf. */
function deriveFileType(fileName: string | undefined): 'pdf' | 'doc' | 'xls' {
  const v: string = (fileName || '').toLowerCase();
  if (v.indexOf('.xls') !== -1 || v.indexOf('.csv') !== -1) {
    return 'xls';
  }
  if (v.indexOf('.doc') !== -1 || v.indexOf('.ppt') !== -1) {
    return 'doc';
  }
  return 'pdf';
}

/** Strips the extension from a file name for display. */
function stripExtension(fileName: string): string {
  const idx: number = fileName.lastIndexOf('.');
  return idx > 0 ? fileName.substring(0, idx) : fileName;
}

export const FALLBACK_POLICIES: IPolicy[] = [];

/**
 * Reads policy files from the Policies document library. Each file is a policy;
 * the title falls back to the file name and the type is derived from the
 * extension. Returns an empty array on error/empty.
 */
export async function getPolicies(context: WebPartContext, listTitle: string): Promise<IPolicy[]> {
  try {
    const rows: IPolicyFileRow[] = await getSP(context)
      .web.lists.getByTitle(listTitle)
      .items.select('Id', 'Title', 'Department', 'Modified', 'PolicyVersion', 'FileLeafRef', 'FileRef', 'FileSystemObjectType')
      .filter('FSObjType eq 0')
      .top(200)();

    if (!rows || rows.length === 0) {
      return FALLBACK_POLICIES;
    }

    return rows.map((it: IPolicyFileRow): IPolicy => {
      const leaf: string = it.FileLeafRef || '';
      return {
        id: it.Id,
        title: it.Title || stripExtension(leaf),
        department: it.Department || '',
        updated: it.Modified || '',
        version: it.PolicyVersion || '',
        fileType: deriveFileType(leaf),
        url: it.FileRef || undefined
      };
    });
  } catch {
    return FALLBACK_POLICIES;
  }
}
