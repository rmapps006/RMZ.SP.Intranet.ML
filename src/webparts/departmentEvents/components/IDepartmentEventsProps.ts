import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentEventsProps {
  context: WebPartContext;
  title: string;
  linkText: string;
  eventsList: string;
  calendarUrl: string;
  maxItems: string;
  showTitle: boolean;
  showViewAll: boolean;
}
