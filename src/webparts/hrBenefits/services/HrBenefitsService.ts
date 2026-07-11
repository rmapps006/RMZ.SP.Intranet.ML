import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { IBenefit } from '../../../common/models';
import { getCurrentLanguage, pickLocalized } from '../../../common/services/languageService';

interface IRawBenefit {
  Id?: number;
  Title?: string;
  TitleAR?: string;
  Category?: string;
  Summary?: string;
  SummaryAR?: string;
}

/**
 * Reads the HR Benefits list (Title, Category, Summary) and maps to IBenefit.
 * Returns an empty array on error/empty so the web part can fall back.
 */
export async function getBenefits(context: WebPartContext, listTitle: string): Promise<IBenefit[]> {
  try {
    const items: IRawBenefit[] = await getSP(context)
      .web.lists.getByTitle(listTitle)
      .items.select('Id', 'Title', 'TitleAR', 'Category', 'Summary', 'SummaryAR')
      .top(24)();

    if (!items || items.length === 0) {
      return [];
    }

    const language = getCurrentLanguage();
    return items.map(
      (it: IRawBenefit): IBenefit => ({
        id: it.Id,
        category: it.Category || '',
        title: pickLocalized(it.Title || '', it.TitleAR, language),
        description: pickLocalized(it.Summary || '', it.SummaryAR, language),
        linkText: 'View Details'
      })
    );
  } catch {
    return [];
  }
}
