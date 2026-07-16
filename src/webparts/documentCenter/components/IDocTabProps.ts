import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Language } from '../../../common/services/languageService';
import { IDocEntry, IDocStats, DocRole } from '../services/DocumentCenterService';

/** Shared context handed to every Document Center tab. */
export interface IDocCtx {
  context: WebPartContext;
  libraryTitle: string;
  language: Language;
  rtl: boolean;
  role: DocRole;
  docs: IDocEntry[];
  stats: IDocStats;
  webUrl: string;
  listId: string;
  currentUserId: number;
  libraryUrl?: string;
  /** Re-fetches documents from the library and re-renders all tabs. */
  reload: () => void;
}
