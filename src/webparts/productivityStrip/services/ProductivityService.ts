import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getGraphClient } from '../../../common/services/graphService';
import { deviceTimeZone } from '../../../common/util/format';

/** A single schedule entry shown in the "Today's Schedule" card. */
export interface IScheduleItem {
  time: string;
  subject: string;
}

interface IGraphDateTime {
  dateTime?: string;
}

interface IGraphEvent {
  subject?: string;
  start?: IGraphDateTime;
  end?: IGraphDateTime;
}

interface IGraphEventsResponse {
  value?: IGraphEvent[];
}

/**
 * Today's date range expressed in the viewer's own timezone, as naive
 * (offset-less) date-times. Combined with the `Prefer: outlook.timezone` header,
 * Graph treats these boundaries — and returns event times — in that timezone.
 */
function localTodayRange(tz: string): { start: string; end: string } {
  const now: Date = new Date();
  const ymd: string = new Intl.DateTimeFormat('en-CA', {
    timeZone: tz,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(now); // en-CA → "YYYY-MM-DD"

  const parts: number[] = ymd.split('-').map((v) => parseInt(v, 10));
  const next: Date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  next.setUTCDate(next.getUTCDate() + 1);
  const ny: string = next.getUTCFullYear().toString();
  const nm: string = (next.getUTCMonth() + 1).toString().padStart(2, '0');
  const nd: string = next.getUTCDate().toString().padStart(2, '0');

  return { start: `${ymd}T00:00:00`, end: `${ny}-${nm}-${nd}T00:00:00` };
}

/**
 * Reads HH:mm from a Graph date-time. With the `Prefer: outlook.timezone` header
 * the value is already in the requested timezone and carries no offset, so we
 * read HH:mm straight from the string.
 */
function formatEventTime(value: string | undefined): string {
  if (!value) {
    return '';
  }
  const match: RegExpExecArray | null = /T(\d{2}):(\d{2})/.exec(value);
  return match ? `${match[1]}:${match[2]}` : '';
}

/**
 * Returns today's calendar events from Graph in the viewer's own timezone.
 * Returns an empty array when there are no events (or on error) — no sample data.
 */
export async function getTodaySchedule(
  context: WebPartContext,
  maxEvents: number
): Promise<IScheduleItem[]> {
  try {
    const tz: string = deviceTimeZone();
    const range: { start: string; end: string } = localTodayRange(tz);

    const client = await getGraphClient(context);
    const response: IGraphEventsResponse = await client
      .api('/me/calendarView')
      .header('Prefer', `outlook.timezone="${tz}"`)
      .query({
        startDateTime: range.start,
        endDateTime: range.end
      })
      .select('subject,start,end')
      .orderby('start/dateTime')
      .top(maxEvents)
      .get();

    const events: IGraphEvent[] = (response && response.value) || [];
    return events.slice(0, maxEvents).map((ev: IGraphEvent) => ({
      time: formatEventTime(ev.start ? ev.start.dateTime : undefined),
      subject: ev.subject || '(No subject)'
    }));
  } catch {
    return [];
  }
}
