import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmployeeDirectoryProps {
  context: WebPartContext;
  fullDirectoryUrl: string;
  pageSize: number;
  showTitle: boolean;
  showViewAll: boolean;
}
