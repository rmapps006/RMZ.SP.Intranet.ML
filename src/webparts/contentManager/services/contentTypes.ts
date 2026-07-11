/** The kind of editor to render for a field. */
export type FieldKind = 'text' | 'multiline' | 'richtext' | 'url' | 'choice' | 'datetime';

/** Describes one editable field on a content type. */
export interface IFieldDef {
  name: string; // SharePoint internal field name
  label: string;
  kind: FieldKind;
  required?: boolean;
  choices?: string[]; // for 'choice' when not fetched from the list
}

/** Describes a content type managed by the dashboard (maps to one list). */
export interface IContentTypeDef {
  key: 'news' | 'events' | 'benefits';
  label: string;
  orderBy: string; // field to sort the list by (newest first)
  // Field shown as the row's secondary line in the list.
  subtitleField?: string;
  fields: IFieldDef[];
}

/** Field definitions per content type. Choice values are fetched live from the
 *  list where possible so the same web part works on main and department sites. */
export const NEWS_TYPE: IContentTypeDef = {
  key: 'news',
  label: 'News',
  orderBy: 'NewsDate',
  subtitleField: 'Category',
  fields: [
    { name: 'Title', label: 'Headline', kind: 'text', required: true },
    { name: 'TitleAR', label: 'Headline (Arabic)', kind: 'text' },
    { name: 'Category', label: 'Category', kind: 'choice' },
    { name: 'NewsDate', label: 'News date', kind: 'datetime' },
    { name: 'Source', label: 'Source / author', kind: 'text' },
    { name: 'SourceAR', label: 'Source / author (Arabic)', kind: 'text' },
    { name: 'LinkUrl', label: 'External link (optional)', kind: 'url' },
    { name: 'ImageUrl', label: 'Image URL (optional)', kind: 'url' },
    { name: 'Body', label: 'Body', kind: 'richtext' },
    { name: 'BodyAR', label: 'Body (Arabic)', kind: 'richtext' }
  ]
};

export const EVENTS_TYPE: IContentTypeDef = {
  key: 'events',
  label: 'Events',
  orderBy: 'EventDate',
  subtitleField: 'Location',
  fields: [
    { name: 'Title', label: 'Event title', kind: 'text', required: true },
    { name: 'TitleAR', label: 'Event title (Arabic)', kind: 'text' },
    { name: 'EventDate', label: 'Starts', kind: 'datetime' },
    { name: 'EndDate', label: 'Ends', kind: 'datetime' },
    { name: 'Location', label: 'Location', kind: 'text' },
    { name: 'LocationAR', label: 'Location (Arabic)', kind: 'text' },
    { name: 'Category', label: 'Category', kind: 'choice' },
    { name: 'Description', label: 'Description', kind: 'multiline' },
    { name: 'DescriptionAR', label: 'Description (Arabic)', kind: 'multiline' }
  ]
};

export const BENEFITS_TYPE: IContentTypeDef = {
  key: 'benefits',
  label: 'HR Benefits',
  orderBy: 'Modified',
  subtitleField: 'Category',
  fields: [
    { name: 'Title', label: 'Benefit name', kind: 'text', required: true },
    { name: 'TitleAR', label: 'Benefit name (Arabic)', kind: 'text' },
    { name: 'Category', label: 'Category', kind: 'choice' },
    { name: 'Summary', label: 'Summary', kind: 'text' },
    { name: 'SummaryAR', label: 'Summary (Arabic)', kind: 'text' },
    { name: 'Eligibility', label: 'Eligibility', kind: 'text' },
    { name: 'EligibilityAR', label: 'Eligibility (Arabic)', kind: 'text' },
    { name: 'Coverage', label: 'Coverage', kind: 'text' },
    { name: 'CoverageAR', label: 'Coverage (Arabic)', kind: 'text' },
    { name: 'Details', label: 'Details', kind: 'richtext' },
    { name: 'DetailsAR', label: 'Details (Arabic)', kind: 'richtext' }
  ]
};
