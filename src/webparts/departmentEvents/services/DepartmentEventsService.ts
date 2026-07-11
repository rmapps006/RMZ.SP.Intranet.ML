import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from '../../../common/services/pnpService';
import { IEventItem } from '../../../common/models';
import { formatClock } from '../../../common/util/format';
import { getCurrentLanguage, pickLocalized } from '../../../common/services/languageService';

// Date-chip background colours, cycled per card (from design/department-page.html).
export const EVENT_ACCENTS: string[] = ['#e3cba8', 'rgba(164,224,230,.28)', 'rgba(38,38,38,.06)'];

// Design fallback used on error or when the list returns nothing.
const FALLBACK_EVENTS: IEventItem[] = [];

interface IRawEventItem {
  Id?: number;
  Title?: string;
  TitleAR?: string;
  EventDate?: string;
  EndDate?: string;
  Location?: string;
  LocationAR?: string;
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
 * Reads the Events list filtered by department and maps to IEventItem. On
 * error/empty returns the three design fallback events.
 */
export async function getDepartmentEvents(
  context: WebPartContext,
  eventsList: string,
  maxItems: number
): Promise<IEventItem[]> {
  try {
    const sp = getSP(context);
    // The department site's Events list is already scoped to that department, so
    // no Department column/filter is required.
    const items: IRawEventItem[] = await sp.web.lists
      .getByTitle(eventsList)
      .items.select('Id', 'Title', 'TitleAR', 'EventDate', 'EndDate', 'Location', 'LocationAR', 'Category')
      .orderBy('EventDate', true)
      .top(maxItems)();

    if (!items || items.length === 0) {
      return FALLBACK_EVENTS;
    }

    const language = getCurrentLanguage();

    return items.map((item: IRawEventItem, index: number): IEventItem => {
      return {
        id: item.Id,
        title: pickLocalized(item.Title || '', item.TitleAR, language),
        type: item.Category || '',
        date: item.EventDate || '',
        time: buildTimeRange(item.EventDate, item.EndDate),
        location: pickLocalized(item.Location || '', item.LocationAR, language),
        accent: EVENT_ACCENTS[index % EVENT_ACCENTS.length]
      };
    });
  } catch {
    return FALLBACK_EVENTS;
  }
}
