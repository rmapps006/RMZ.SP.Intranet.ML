import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentDocumentsProps {
  context: WebPartContext;
  title: string;
  linkText: string;
  department: string;
  policiesLibrary: string;
  documentsLibrary: string;
  documentHubUrl: string;
  panel1Title: string;
  panel1TitleAR: string;
  panel2Title: string;
  panel2TitleAR: string;
  showTitle: boolean;
  showViewAll: boolean;
}
