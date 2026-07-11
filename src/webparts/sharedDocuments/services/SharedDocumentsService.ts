import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { formatLongDate, browserViewUrl } from '../../../common/util/format';
import { IDocumentItem } from '../../../common/models';

interface ILibraryFileRow {
  FileLeafRef?: string;
  FileRef?: string;
  Modified?: string;
  File_x0020_Type?: string;
}

/** Maps a file extension / SP file type to a design badge type. */
function deriveFileType(ext: string | undefined): 'pdf' | 'doc' | 'xls' {
  const e: string = (ext || '').toLowerCase();
  if (e === 'pdf') {
    return 'pdf';
  }
  if (e === 'xls' || e === 'xlsx' || e === 'xlsm' || e === 'csv') {
    return 'xls';
  }
  return 'doc';
}

/** Strips the extension from a file name for display. */
function stripExtension(fileName: string): string {
  const idx: number = fileName.lastIndexOf('.');
  return idx > 0 ? fileName.substring(0, idx) : fileName;
}

/**
 * Reads the most recent files from a SharePoint document library by title.
 * Returns an empty array on error so the caller can fall back to design data.
 */
export async function getLibraryDocuments(
  context: WebPartContext,
  libraryTitle: string
): Promise<IDocumentItem[]> {
  try {
    const rows: ILibraryFileRow[] = await getSP(context)
      .web.lists.getByTitle(libraryTitle)
      .items.select('FileLeafRef', 'FileRef', 'Modified', 'File_x0020_Type')
      .top(8)();

    const items: IDocumentItem[] = (rows || [])
      .filter((r: ILibraryFileRow) => !!r.FileLeafRef)
      .map((r: ILibraryFileRow): IDocumentItem => {
        const leaf: string = r.FileLeafRef as string;
        const ext: string =
          r.File_x0020_Type || (leaf.indexOf('.') > -1 ? leaf.substring(leaf.lastIndexOf('.') + 1) : '');
        return {
          name: stripExtension(leaf),
          modified: formatLongDate(r.Modified),
          fileType: deriveFileType(ext),
          url: browserViewUrl(r.FileRef)
        };
      });

    return items;
  } catch {
    return [];
  }
}
