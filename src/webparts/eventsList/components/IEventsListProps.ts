import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEventsListProps {
  context: WebPartContext;
  eventsList: string;
  calendarUrl: string;
  maxItems: string;
  showTitle: boolean;
  showViewAll: boolean;
  title: string;
  linkText: string;
}
