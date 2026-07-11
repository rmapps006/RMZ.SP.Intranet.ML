import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { getSP } from './pnpService';
import { imageFieldUrl } from '../util/format';

/** Raw News list row. The picture is a plain image URL in the ImageUrl field. */
export interface IRawNewsItem {
  Id?: number;
  Title?: string;
  Category?: string;
  NewsDate?: string;
  Source?: string;
  LinkUrl?: string;
  ImageUrl?: string;
  Created?: string;
}

/**
 * Reads News rows. The picture comes from a plain text `ImageUrl` field (paste a
 * hosted image URL). Selecting a non-existent field errors, so a list without
 * ImageUrl falls back to the base fields.
 */
export async function readNewsRaw(context: WebPartContext, listTitle: string, top: number): Promise<IRawNewsItem[]> {
  const list = getSP(context).web.lists.getByTitle(listTitle);
  const base: string[] = ['Id', 'Title', 'Category', 'NewsDate', 'Source', 'LinkUrl', 'Created'];
  const attempts: string[][] = [[...base, 'ImageUrl'], base];
  for (let i = 0; i < attempts.length; i++) {
    try {
      return await list.items.select(...attempts[i]).orderBy('NewsDate', false).top(top)();
    } catch (e) {
      if (i === attempts.length - 1) {
        throw e;
      }
    }
  }
  return [];
}

/** Resolves the display image URL from the row's ImageUrl field. */
export function newsImageUrl(item: IRawNewsItem): string | undefined {
  return imageFieldUrl(item.ImageUrl);
}
