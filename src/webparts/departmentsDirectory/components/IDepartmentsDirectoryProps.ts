import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentsDirectoryProps {
  context: WebPartContext;
  title: string;
  viewAllUrl: string;
  departmentsJson: string;
  showTitle: boolean;
  showViewAll: boolean;
}
