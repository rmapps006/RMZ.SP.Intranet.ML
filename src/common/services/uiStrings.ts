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
  | 'searchDocuments'
  | 'allTypes'
  | 'allStatuses'
  | 'type'
  | 'status'
  | 'owner'
  | 'reviewDate'
  | 'noDocumentsMatch'
  | 'documentsCount'
  // Section "view all" link defaults (used when no admin override is set)
  | 'viewArchive'
  | 'viewCalendar'
  | 'allPolicies'
  | 'documentHub'
  | 'fullDirectory'
  | 'tabDashboard'
  | 'tabDocuments'
  | 'tabUpload'
  | 'tabMyDocuments'
  | 'tabReports'
  | 'tabAdmin'
  | 'totalDocuments'
  | 'publishedCount'
  | 'inReviewCount'
  | 'draftCount'
  | 'overdueReview'
  | 'myDocuments'
  | 'needsAttention'
  | 'recentDocuments'
  | 'byStatus'
  | 'byDepartment'
  | 'byType'
  | 'bySensitivity'
  | 'nothingOverdue'
  | 'sensitivity'
  | 'allDepartments'
  | 'allSensitivities'
  | 'documentNumber'
  | 'tagsLabel'
  | 'chooseFile'
  | 'titleField'
  | 'titleARField'
  | 'ownerField'
  | 'reviewDateField'
  | 'descriptionField'
  | 'submit'
  | 'uploading'
  | 'uploadSuccess'
  | 'uploadFailed'
  | 'autoNumberNote'
  | 'requiredField'
  | 'publish'
  | 'submitForReview'
  | 'archive'
  | 'versionHistory'
  | 'open'
  | 'stageDraft'
  | 'stageInReview'
  | 'stagePublished'
  | 'noMyDocuments'
  | 'manageTaxonomies'
  | 'documentTypes'
  | 'statuses'
  | 'sensitivities'
  | 'numberingFormat'
  | 'save'
  | 'saved'
  | 'addValue'
  | 'remove'
  | 'approvalRoutingNote'
  | 'roleAdministrator'
  | 'roleApprover'
  | 'roleContributor'
  | 'roleReader';

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
  searchDocuments: { en: 'Search documents…', ar: 'ابحث في المستندات…' },
  allTypes: { en: 'All Types', ar: 'كل الأنواع' },
  allStatuses: { en: 'All Statuses', ar: 'كل الحالات' },
  type: { en: 'Type', ar: 'النوع' },
  status: { en: 'Status', ar: 'الحالة' },
  owner: { en: 'Owner', ar: 'المسؤول' },
  reviewDate: { en: 'Review by', ar: 'المراجعة بحلول' },
  noDocumentsMatch: { en: 'No documents match your filters.', ar: 'لا توجد مستندات مطابقة لعوامل التصفية.' },
  documentsCount: { en: 'documents', ar: 'مستند' },
  viewArchive: { en: 'View Archive', ar: 'عرض الأرشيف' },
  viewCalendar: { en: 'View Calendar', ar: 'عرض التقويم' },
  allPolicies: { en: 'All Policies', ar: 'كل السياسات' },
  documentHub: { en: 'Document Hub', ar: 'مركز المستندات' },
  fullDirectory: { en: 'Full Directory', ar: 'الدليل الكامل' },
  tabDashboard: { en: 'Dashboard', ar: 'لوحة المعلومات' },
  tabDocuments: { en: 'Documents', ar: 'المستندات' },
  tabUpload: { en: 'Upload', ar: 'رفع' },
  tabMyDocuments: { en: 'My Documents', ar: 'مستنداتي' },
  tabReports: { en: 'Reports', ar: 'التقارير' },
  tabAdmin: { en: 'Administration', ar: 'الإدارة' },
  totalDocuments: { en: 'Total documents', ar: 'إجمالي المستندات' },
  publishedCount: { en: 'Published', ar: 'منشور' },
  inReviewCount: { en: 'In review', ar: 'قيد المراجعة' },
  draftCount: { en: 'Draft', ar: 'مسودة' },
  overdueReview: { en: 'Overdue for review', ar: 'متأخر للمراجعة' },
  myDocuments: { en: 'My documents', ar: 'مستنداتي' },
  needsAttention: { en: 'Needs attention', ar: 'يتطلب انتباهاً' },
  recentDocuments: { en: 'Recent documents', ar: 'أحدث المستندات' },
  byStatus: { en: 'By status', ar: 'حسب الحالة' },
  byDepartment: { en: 'By department', ar: 'حسب الإدارة' },
  byType: { en: 'By type', ar: 'حسب النوع' },
  bySensitivity: { en: 'By sensitivity', ar: 'حسب الحساسية' },
  nothingOverdue: { en: 'Nothing overdue for review.', ar: 'لا شيء متأخر للمراجعة.' },
  sensitivity: { en: 'Sensitivity', ar: 'الحساسية' },
  allDepartments: { en: 'All Departments', ar: 'كل الإدارات' },
  allSensitivities: { en: 'All Sensitivities', ar: 'كل مستويات الحساسية' },
  documentNumber: { en: 'Document number', ar: 'رقم المستند' },
  tagsLabel: { en: 'Tags', ar: 'الوسوم' },
  chooseFile: { en: 'Choose a file', ar: 'اختر ملفاً' },
  titleField: { en: 'Title', ar: 'العنوان' },
  titleARField: { en: 'Title (Arabic)', ar: 'العنوان (بالعربية)' },
  ownerField: { en: 'Owner', ar: 'المسؤول' },
  reviewDateField: { en: 'Review date', ar: 'تاريخ المراجعة' },
  descriptionField: { en: 'Description', ar: 'الوصف' },
  submit: { en: 'Upload document', ar: 'رفع المستند' },
  uploading: { en: 'Uploading…', ar: 'جارٍ الرفع…' },
  uploadSuccess: { en: 'Uploaded — document number', ar: 'تم الرفع — رقم المستند' },
  uploadFailed: { en: 'Upload failed', ar: 'فشل الرفع' },
  autoNumberNote: { en: 'A document number is assigned automatically.', ar: 'يتم تعيين رقم المستند تلقائياً.' },
  requiredField: { en: 'Required', ar: 'مطلوب' },
  publish: { en: 'Publish', ar: 'نشر' },
  submitForReview: { en: 'Submit for review', ar: 'إرسال للمراجعة' },
  archive: { en: 'Archive', ar: 'أرشفة' },
  versionHistory: { en: 'Version history', ar: 'سجل الإصدارات' },
  open: { en: 'Open', ar: 'فتح' },
  stageDraft: { en: 'Draft', ar: 'مسودة' },
  stageInReview: { en: 'In Review', ar: 'قيد المراجعة' },
  stagePublished: { en: 'Published', ar: 'منشور' },
  noMyDocuments: { en: 'You have no documents yet.', ar: 'ليس لديك مستندات بعد.' },
  manageTaxonomies: { en: 'Manage taxonomies', ar: 'إدارة التصنيفات' },
  documentTypes: { en: 'Document types', ar: 'أنواع المستندات' },
  statuses: { en: 'Statuses', ar: 'الحالات' },
  sensitivities: { en: 'Sensitivity levels', ar: 'مستويات الحساسية' },
  numberingFormat: { en: 'Numbering format', ar: 'صيغة الترقيم' },
  save: { en: 'Save', ar: 'حفظ' },
  saved: { en: 'Saved.', ar: 'تم الحفظ.' },
  addValue: { en: 'Add', ar: 'إضافة' },
  remove: { en: 'Remove', ar: 'إزالة' },
  approvalRoutingNote: {
    en: 'Approval routing and email notifications are driven by a Power Automate flow (configured separately).',
    ar: 'يتم تنفيذ مسار الموافقات وإشعارات البريد عبر تدفق Power Automate (يُهيّأ بشكل منفصل).'
  },
  roleAdministrator: { en: 'Administrator', ar: 'مسؤول' },
  roleApprover: { en: 'Approver', ar: 'معتمِد' },
  roleContributor: { en: 'Contributor', ar: 'مساهم' },
  roleReader: { en: 'Reader', ar: 'قارئ' }
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
  'All Types': 'كل الأنواع',
  'All Statuses': 'كل الحالات',
  General: 'عام',
  // Document Center document types
  Policy: 'سياسة',
  Procedure: 'إجراء',
  Form: 'نموذج',
  Template: 'قالب',
  Report: 'تقرير',
  Guideline: 'إرشادات',
  Contract: 'عقد',
  // Document Center statuses / stages
  Draft: 'مسودة',
  'In Review': 'قيد المراجعة',
  Approved: 'معتمد',
  Published: 'منشور',
  Archived: 'مؤرشف',
  // Document Center sensitivity levels
  Public: 'عام',
  Internal: 'داخلي',
  Confidential: 'سري',
  Restricted: 'مقيّد',
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
