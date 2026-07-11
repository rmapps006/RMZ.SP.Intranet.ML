import { WebPartContext } from '@microsoft/sp-webpart-base';
import { INewsItem } from '../../../common/models';
import { readNewsRaw, newsImageUrl, IRawNewsItem } from '../../../common/services/newsRead';
import { getCurrentLanguage, pickLocalized } from '../../../common/services/languageService';

// News card image gradients, cycled per card (from design/department-page.html).
export const NEWS_GRADIENTS: string[] = [
  'linear-gradient(145deg,#c4b49a 0%,#a89882 40%,#cfc4b6 100%)',
  'linear-gradient(145deg,#d8c8b4 0%,#e8ddd0 60%,#c8bab0 100%)',
  'linear-gradient(145deg,#b8c4b4 0%,#d0d8cc 60%,#a8b4a8 100%)'
];

// Faint photography captions overlaid on each card image, cycled per card.
export const NEWS_CAPTIONS: string[] = ['People Photography', 'Office Photography', 'Wellness Photography'];

// Design fallback used on error or when the source returns nothing.
const FALLBACK_NEWS: INewsItem[] = [];

/**
 * Reads the department site's local News list (a custom list — separate from
 * Site Pages) and maps to INewsItem. Returns an empty array on error/empty.
 */
export async function getDepartmentNews(context: WebPartContext, newsList: string): Promise<INewsItem[]> {
  try {
    const items: IRawNewsItem[] = await readNewsRaw(context, newsList, 3);

    if (!items || items.length === 0) {
      return FALLBACK_NEWS;
    }

    const language = getCurrentLanguage();
    return items.map((item: IRawNewsItem, index: number): INewsItem => ({
      id: item.Id,
      title: pickLocalized(item.Title || '', item.TitleAR, language),
      category: item.Category || '',
      date: item.NewsDate || item.Created || '',
      author: pickLocalized(item.Source || '', item.SourceAR, language),
      url: item.LinkUrl || undefined,
      imageUrl: newsImageUrl(item),
      imageGradient: NEWS_GRADIENTS[index % NEWS_GRADIENTS.length],
      imageCaption: NEWS_CAPTIONS[index % NEWS_CAPTIONS.length]
    }));
  } catch {
    return FALLBACK_NEWS;
  }
}
