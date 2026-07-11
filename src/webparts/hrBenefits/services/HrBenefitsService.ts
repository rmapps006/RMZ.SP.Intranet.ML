import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { IBenefit } from '../../../common/models';

interface IRawBenefit {
  Id?: number;
  Title?: string;
  Category?: string;
  Summary?: string;
}

/**
 * Reads the HR Benefits list (Title, Category, Summary) and maps to IBenefit.
 * Returns an empty array on error/empty so the web part can fall back.
 */
export async function getBenefits(context: WebPartContext, listTitle: string): Promise<IBenefit[]> {
  try {
    const items: IRawBenefit[] = await getSP(context)
      .web.lists.getByTitle(listTitle)
      .items.select('Id', 'Title', 'Category', 'Summary')
      .top(24)();

    if (!items || items.length === 0) {
      return [];
    }

    return items.map(
      (it: IRawBenefit): IBenefit => ({
        id: it.Id,
        category: it.Category || '',
        title: it.Title || '',
        description: it.Summary || '',
        linkText: 'View Details'
      })
    );
  } catch {
    return [];
  }
}
