import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentNewsProps {
  context: WebPartContext;
  title: string;
  linkText: string;
  newsList: string;
  allNewsUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}
