import { WebPartContext } from '@microsoft/sp-webpart-base';
import { INewsItem } from '../../../common/models';
import { formatLongDate } from '../../../common/util/format';
import { readNewsRaw, newsImageUrl, IRawNewsItem } from '../../../common/services/newsRead';
import { getCurrentLanguage, pickLocalized } from '../../../common/services/languageService';

// Gradient backgrounds for the news image blocks (from design/v1-homepage.html).
const NEWS_GRADIENTS: string[] = [
  'linear-gradient(145deg,#c4b49a,#a89882 40%,#cfc4b6)',
  'linear-gradient(145deg,#9aaa9e,#b5c4b8 60%,#8a9a8e)',
  'linear-gradient(145deg,#d8c8b4,#e8ddd0 60%,#c8bab0)'
];

// Category pills shown above the grid (first is the implicit "All").
export const NEWS_CATEGORIES: string[] = [
  'All',
  'General Announcements',
  'Construction Updates',
  'HR Updates',
  'Finance'
];

// Design fallback used on error or when the library returns nothing.
const FALLBACK_NEWS: INewsItem[] = [];

/**
 * Reads the News list (custom list — separate from Site Pages) and maps to
 * INewsItem. Returns an empty array on error/empty.
 */
export async function getNews(context: WebPartContext, newsList: string, top: number = 9): Promise<INewsItem[]> {
  try {
    const items: IRawNewsItem[] = await readNewsRaw(context, newsList, top);

    if (!items || items.length === 0) {
      return FALLBACK_NEWS;
    }

    const language = getCurrentLanguage();
    return items.map((item: IRawNewsItem, index: number): INewsItem => ({
      id: item.Id,
      title: pickLocalized(item.Title || '', item.TitleAR, language),
      category: item.Category || 'General',
      date: formatLongDate(item.NewsDate || item.Created),
      author: pickLocalized(item.Source || '', item.SourceAR, language),
      url: item.LinkUrl || undefined,
      imageUrl: newsImageUrl(item),
      imageGradient: NEWS_GRADIENTS[index % NEWS_GRADIENTS.length]
    }));
  } catch {
    return FALLBACK_NEWS;
  }
}
