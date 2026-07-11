import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface INewsCarouselProps {
  context: WebPartContext;
  newsList: string;
  archiveUrl: string;
  maxItems: number;
  showTitle: boolean;
  showViewAll: boolean;
}
