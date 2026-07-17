import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Language } from '../../../common/services/languageService';
import { IDocEntry, IDocStats, DocRole } from '../services/DocumentCenterService';

/** Shared context handed to every Document Center tab. */
export interface IDocCtx {
  context: WebPartContext;
  libraryTitle: string;
  /** Optional URL of the site hosting the library; blank = the page's own site. */
  siteUrl: string;
  language: Language;
  rtl: boolean;
  role: DocRole;
  docs: IDocEntry[];
  stats: IDocStats;
  webUrl: string;
  listId: string;
  currentUserId: number;
  libraryUrl?: string;
  /** Populated when the document query failed outright (library missing / no access). */
  loadError?: string;
  /** Re-fetches documents from the library and re-renders all tabs. */
  reload: () => void;
}
