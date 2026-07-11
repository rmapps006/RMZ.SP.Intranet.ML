import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISharedDocumentsProps {
  context: WebPartContext;
  panel1Library: string;
  panel1Title: string;
  panel1Url: string;
  panel2Library: string;
  panel2Title: string;
  panel2Url: string;
  documentHubUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
  title: string;
  linkText: string;
}
