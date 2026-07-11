// Shared data models for the Intranet homepage web parts.

export interface IPerson {
  displayName: string;
  jobTitle?: string;
  department?: string;
  email?: string;
  phone?: string;
  location?: string;
  initials?: string;
}

export interface ICelebration {
  name: string;
  type: 'birthday' | 'anniversary';
  date: string; // display string e.g. "June 11"
  detail?: string; // e.g. "3 Years"
  initials?: string;
  gradient?: string;
}

export interface INewsItem {
  id?: number; // list item id (for linking to the detail page)
  title: string;
  category: string;
  excerpt?: string;
  date: string;
  author?: string;
  imageUrl?: string;
  imageGradient?: string;
  imageCaption?: string; // faint overlay caption on the image (e.g. "People Photography")
  url?: string;
}

export interface IEventItem {
  id?: number; // list item id (for linking to the detail page)
  title: string;
  type: string;
  date: string; // ISO or display
  time?: string;
  location?: string;
  accent?: string; // date-chip background
  url?: string; // link to the event's list item (detail)
}

export interface IBenefit {
  id?: number; // list item id (for linking to the detail page)
  category: string;
  title: string;
  description: string;
  linkText?: string;
  url?: string;
}

export interface IPolicy {
  id?: number; // list item id (for linking to the detail page)
  title: string;
  department: string;
  updated: string;
  version: string;
  fileType: 'pdf' | 'doc' | 'xls';
  url?: string;
}

export interface IDocumentItem {
  name: string;
  modified: string;
  fileType: 'pdf' | 'doc' | 'xls';
  url?: string;
}

export interface IDocumentPanel {
  title: string;
  libraryUrl?: string;
  items: IDocumentItem[];
}

export interface IOrgNode {
  name: string;
  role: string;
  isRoot?: boolean;
}
