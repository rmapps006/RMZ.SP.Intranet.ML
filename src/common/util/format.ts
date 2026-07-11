// Shared formatting helpers for the Intranet web parts.

function pad2(n: number): string {
  return n.toString().padStart(2, '0');
}

/**
 * Formats a date/time as HH:mm (24-hour) in the viewer's own timezone. Times are
 * stored in UTC by SharePoint and shown in local time, so each user sees the
 * times in their own zone.
 */
export function formatClock(value: string | Date | undefined): string {
  if (!value) {
    return '';
  }
  const d: Date = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) {
    return '';
  }
  return `${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

/** The viewer's IANA timezone (e.g. 'Asia/Dubai'), used for Graph calendar reads. */
export function deviceTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
  } catch {
    return 'UTC';
  }
}

const GRADIENTS: string[] = [
  'linear-gradient(135deg,#e3cba8,#c8a880)',
  'linear-gradient(135deg,#6ab8c0,#a4e0e6)',
  'linear-gradient(135deg,#9aaa9e,#c4cec8)',
  'linear-gradient(135deg,#c8a880,#e3cba8)',
  'linear-gradient(135deg,#b09878,#d4b898)',
  'linear-gradient(135deg,#d4846c,#ea8b6e)',
  'linear-gradient(135deg,#262626,#4a3c30)',
  'linear-gradient(135deg,#8a8682,#b5b5b5)',
  'linear-gradient(135deg,#7a7470,#a8a4a0)'
];

/** Encodes spaces so a URL is safe in CSS `url("…")` and `<img src>`. */
function encodeSpaces(url: string): string {
  return url.replace(/ /g, '%20');
}

/**
 * Resolves the display URL from a modern SharePoint "Image" (Thumbnail) column.
 * The column stores a JSON blob (e.g. `{ "serverRelativeUrl": "…", … }`) when an
 * image is uploaded; a plain string is treated as a direct URL for backward
 * compatibility. Returns undefined when empty or unresolvable.
 *
 * The blob can arrive HTML-encoded (e.g. `&quot;`) depending on how the field is
 * returned, which would break a naive JSON.parse — so we decode common entities
 * first and, if parsing still fails, fall back to a regex extraction of the URL.
 */
export function imageFieldUrl(value: unknown): string | undefined {
  if (!value) {
    return undefined;
  }
  // Some layers may hand back the already-parsed object rather than the raw JSON.
  if (typeof value === 'object') {
    const obj: { serverRelativeUrl?: string; url?: string; Url?: string } = value as never;
    const u: string | undefined = obj.serverRelativeUrl || obj.url || obj.Url;
    return u ? encodeSpaces(u) : undefined;
  }
  if (typeof value !== 'string') {
    return undefined;
  }
  let raw: string = value.trim();
  if (!raw) {
    return undefined;
  }

  // Not a JSON blob → treat as a direct URL (legacy text ImageUrl field).
  if (raw.indexOf('{') === -1) {
    return encodeSpaces(raw);
  }

  // Decode the HTML entities SharePoint may wrap the blob in before parsing.
  if (raw.indexOf('&') !== -1) {
    raw = raw
      .replace(/&quot;/g, '"')
      .replace(/&#34;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');
  }

  try {
    const parsed: { serverRelativeUrl?: string; url?: string; Url?: string } = JSON.parse(raw);
    const url: string | undefined = parsed.serverRelativeUrl || parsed.url || parsed.Url;
    if (url) {
      return encodeSpaces(url);
    }
  } catch {
    /* fall through to regex extraction */
  }

  // Last resort: pull the URL straight out of the blob even if JSON.parse choked.
  const match: RegExpExecArray | null = /(?:serverRelativeUrl|url)"\s*:\s*"([^"]+)"/i.exec(raw);
  return match && match[1] ? encodeSpaces(match[1]) : undefined;
}

/**
 * Anchor attributes for the central "open links in a new tab" setting. Spread
 * onto an `<a>` (e.g. `{...linkTarget(newTab)}`): when newTab is true it returns
 * `target="_blank"` with a safe `rel`; otherwise an empty object (same tab).
 */
export function linkTarget(newTab: boolean): { target?: string; rel?: string; 'data-interception'?: string } {
  // data-interception="off" stops SharePoint's SPA router from hijacking the
  // click, so target="_blank" actually opens a new tab for internal links too.
  return newTab ? { target: '_blank', rel: 'noopener noreferrer', 'data-interception': 'off' } : {};
}

/**
 * Ensures a document URL opens in the browser (Office for the web / built-in
 * viewer) instead of downloading, by adding SharePoint's `web=1` hint. Leaves
 * empty/absent URLs untouched and won't duplicate the param.
 */
export function browserViewUrl(url: string | undefined): string | undefined {
  if (!url) {
    return undefined;
  }
  if (/[?&]web=1(&|$)/.test(url)) {
    return url;
  }
  return url + (url.indexOf('?') === -1 ? '?web=1' : '&web=1');
}

/** Initials from a display name (e.g. "Ahmed Al-Karimi" -> "AK"). */
export function getInitials(name: string | undefined): string {
  const source: string = (name || '').trim();
  if (!source) {
    return '?';
  }
  const parts: string[] = source.split(/\s+/).filter((p) => p.length > 0);
  if (parts.length === 1) {
    return parts[0].substring(0, 2).toUpperCase();
  }
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
}

/** Deterministic avatar gradient picked from a seed string. */
export function pickGradient(seed: string | undefined): string {
  const s: string = seed || '';
  let hash: number = 0;
  for (let i: number = 0; i < s.length; i++) {
    hash = (hash * 31 + s.charCodeAt(i)) | 0;
  }
  return GRADIENTS[Math.abs(hash) % GRADIENTS.length];
}

/** Short date e.g. "June 10, 2026". */
export function formatLongDate(value: string | Date | undefined): string {
  if (!value) {
    return '';
  }
  const d: Date = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) {
    return '';
  }
  return d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
}

/** Month abbreviation + day for event date chips (e.g. { mo: "JUN", dy: "18" }). */
export function splitDateChip(value: string | Date | undefined): { mo: string; dy: string } {
  if (!value) {
    return { mo: '', dy: '' };
  }
  const d: Date = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) {
    return { mo: '', dy: '' };
  }
  return {
    mo: d.toLocaleDateString('en-US', { month: 'short' }).toUpperCase(),
    dy: d.getDate().toString()
  };
}
