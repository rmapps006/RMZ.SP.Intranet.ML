import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeDirectoryProps {
  context: WebPartContext;
  title: string;
  linkText: string;
  fullDirectoryUrl: string;
  pageSize: number;
  showTitle: boolean;
  showViewAll: boolean;
}
