import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentNewsProps {
  context: WebPartContext;
  newsList: string;
  allNewsUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}
