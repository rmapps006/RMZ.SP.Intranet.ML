/**
 * Fixed UI-chrome translations (English/Arabic) for text that is baked into the
 * web parts rather than coming from a SharePoint list or an admin setting —
 * section link labels, table headers, filter chips, button captions, search
 * placeholders and empty/loading states.
 *
 * List content (titles, bodies) and admin-entered copy are localised separately
 * via `pickLocalized`; this module is only for the fixed interface strings, so
 * the Arabic view reads fully in Arabic without every label needing its own
 * admin field.
 */

import { Language } from './languageService';

export type UiStringKey =
  | 'todaysSchedule'
  | 'quickLinks'
  | 'noEventsToday'
  | 'noQuickLinks'
  | 'searchPolicies'
  | 'policyTitle'
  | 'department'
  | 'lastUpdated'
  | 'version'
  | 'loadingPolicies'
  | 'noPoliciesMatch'
  | 'searchDirectory'
  | 'call'
  | 'email'
  | 'teams'
  | 'searchingDirectory'
  | 'noMatchingPeople'
  | 'loadingNews'
  | 'noNews'
  | 'loadingEvents'
  | 'noEvents'
  | 'openLibrary'
  | 'loading'
  | 'noDocuments'
  // Section "view all" link defaults (used when no admin override is set)
  | 'viewArchive'
  | 'viewCalendar'
  | 'allPolicies'
  | 'documentHub'
  | 'fullDirectory';

const STRINGS: Record<UiStringKey, { en: string; ar: string }> = {
  todaysSchedule: { en: "Today's Schedule", ar: 'جدول اليوم' },
  quickLinks: { en: 'Quick Links', ar: 'روابط سريعة' },
  noEventsToday: { en: 'No events scheduled today.', ar: 'لا توجد فعاليات مجدولة اليوم.' },
  noQuickLinks: { en: 'No quick links configured.', ar: 'لم تتم تهيئة روابط سريعة.' },
  searchPolicies: { en: 'Search policies…', ar: 'ابحث في السياسات…' },
  policyTitle: { en: 'Policy Title', ar: 'عنوان السياسة' },
  department: { en: 'Department', ar: 'الإدارة' },
  lastUpdated: { en: 'Last Updated', ar: 'آخر تحديث' },
  version: { en: 'Version', ar: 'الإصدار' },
  loadingPolicies: { en: 'Loading policies…', ar: 'جارٍ تحميل السياسات…' },
  noPoliciesMatch: { en: 'No policies match your search.', ar: 'لا توجد سياسات مطابقة لبحثك.' },
  searchDirectory: {
    en: 'Search the full directory by name, email, department, position…',
    ar: 'ابحث في الدليل بالاسم أو البريد أو الإدارة أو المنصب…'
  },
  call: { en: 'Call', ar: 'اتصال' },
  email: { en: 'Email', ar: 'بريد' },
  teams: { en: 'Teams', ar: 'تيمز' },
  searchingDirectory: { en: 'Searching the directory…', ar: 'جارٍ البحث في الدليل…' },
  noMatchingPeople: { en: 'No matching people found.', ar: 'لا يوجد أشخاص مطابقون.' },
  loadingNews: { en: 'Loading news…', ar: 'جارٍ تحميل الأخبار…' },
  noNews: { en: 'No news to show.', ar: 'لا توجد أخبار لعرضها.' },
  loadingEvents: { en: 'Loading events…', ar: 'جارٍ تحميل الفعاليات…' },
  noEvents: { en: 'No events to show.', ar: 'لا توجد فعاليات لعرضها.' },
  openLibrary: { en: 'Open Library', ar: 'فتح المكتبة' },
  loading: { en: 'Loading…', ar: 'جارٍ التحميل…' },
  noDocuments: { en: 'No documents yet.', ar: 'لا توجد مستندات بعد.' },
  viewArchive: { en: 'View Archive', ar: 'عرض الأرشيف' },
  viewCalendar: { en: 'View Calendar', ar: 'عرض التقويم' },
  allPolicies: { en: 'All Policies', ar: 'كل السياسات' },
  documentHub: { en: 'Document Hub', ar: 'مركز المستندات' },
  fullDirectory: { en: 'Full Directory', ar: 'الدليل الكامل' }
};

/** Returns the fixed UI string for the given key in the active language. */
export function t(key: UiStringKey, language: Language): string {
  const entry = STRINGS[key];
  return language === 'ar' ? entry.ar : entry.en;
}

/**
 * Arabic labels for the fixed SharePoint choice values the solution ships with
 * (News/Events categories, Policy departments, Benefit categories, and the
 * "show everything" filter options). These are stored in the lists in English;
 * this maps a stored value to its Arabic label for display only — the English
 * value is still what filtering compares against. Unknown/custom values fall
 * back to the value itself.
 */
const CHOICE_AR: Record<string, string> = {
  All: 'الكل',
  'All Departments': 'كل الإدارات',
  // News categories
  'General Announcements': 'الإعلانات العامة',
  'Construction Updates': 'تحديثات الإنشاءات',
  'HR Updates': 'تحديثات الموارد البشرية',
  // Events categories
  'All-Company': 'على مستوى الشركة',
  'Learning & Development': 'التعلّم والتطوير',
  // Departments / shared
  HR: 'الموارد البشرية',
  Finance: 'الشؤون المالية',
  IT: 'تقنية المعلومات',
  Operations: 'العمليات',
  Marketing: 'التسويق',
  Legal: 'الشؤون القانونية',
  // Benefit categories
  Health: 'الصحة',
  Wellness: 'العافية',
  Leave: 'الإجازات',
  Financial: 'الأمور المالية',
  Learning: 'التعلّم',
  Family: 'العائلة'
};

/** Localises a known SharePoint choice value; returns the raw value when unmapped or not Arabic. */
export function localizeChoice(value: string, language: Language): string {
  if (language !== 'ar') {
    return value;
  }
  const key: string = (value || '').trim();
  return CHOICE_AR[key] || value;
}
