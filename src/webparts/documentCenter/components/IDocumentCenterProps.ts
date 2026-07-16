import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDocumentCenterProps {
  context: WebPartContext;
  libraryTitle: string;
  pageSize: number;
  showTitle: boolean;
  showViewAll: boolean;
  title: string;
  linkText: string;
}
