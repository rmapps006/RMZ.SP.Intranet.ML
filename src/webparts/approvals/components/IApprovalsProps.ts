import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IApprovalsProps {
  context: WebPartContext;
  requestsList: string;
  department: string;
  viewAllUrl: string;
  newRequestUrl: string;
  reviewAllUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}
