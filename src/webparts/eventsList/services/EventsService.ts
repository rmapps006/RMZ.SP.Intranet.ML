import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { IEventItem } from '../../../common/models';
import { formatClock } from '../../../common/util/format';

// Date-chip background colours, cycled per card (from design/v1-homepage.html).
export const EVENT_ACCENTS: string[] = [
  '#e3cba8',
  'rgba(164,224,230,.28)',
  'rgba(38,38,38,.06)',
  'rgba(234,139,110,.15)'
];

// Design fallback used on error or when the list returns nothing.
const FALLBACK_EVENTS: IEventItem[] = [];

interface IRawEventItem {
  Id?: number;
  Title?: string;
  EventDate?: string;
  EndDate?: string;
  Location?: string;
  Category?: string;
}

// Event times are shown in the viewer's own timezone (stored UTC → local).
function buildTimeRange(start: string | undefined, end: string | undefined): string {
  const s: string = formatClock(start);
  const e: string = formatClock(end);
  if (s && e) {
    return `${s} – ${e}`;
  }
  return s || '';
}

/**
 * Reads the Events list and maps to IEventItem. On error/empty returns the
 * four design fallback events.
 */
export async function getEvents(
  context: WebPartContext,
  eventsList: string,
  maxItems: number
): Promise<IEventItem[]> {
  try {
    const sp = getSP(context);
    const list = sp.web.lists.getByTitle(eventsList);
    const [items, rootFolder] = await Promise.all([
      list.items.select('Id', 'Title', 'EventDate', 'EndDate', 'Location', 'Category').orderBy('EventDate', true).top(maxItems)(),
      list.rootFolder.select('ServerRelativeUrl')()
    ]);

    if (!items || items.length === 0) {
      return FALLBACK_EVENTS;
    }

    const listUrl: string = (rootFolder && rootFolder.ServerRelativeUrl) || '';

    return (items as IRawEventItem[]).map((item: IRawEventItem, index: number): IEventItem => {
      return {
        id: item.Id,
        title: item.Title || '',
        type: item.Category || '',
        date: item.EventDate || '',
        time: buildTimeRange(item.EventDate, item.EndDate),
        location: item.Location || '',
        accent: EVENT_ACCENTS[index % EVENT_ACCENTS.length],
        url: listUrl && item.Id !== undefined ? `${listUrl}/DispForm.aspx?ID=${item.Id}` : undefined
      };
    });
  } catch {
    return FALLBACK_EVENTS;
  }
}
