import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentDocumentsProps {
  context: WebPartContext;
  department: string;
  policiesLibrary: string;
  documentsLibrary: string;
  documentHubUrl: string;
  panel1Title: string;
  panel2Title: string;
  showTitle: boolean;
  showViewAll: boolean;
}
