import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IContentManagerProps {
  context: WebPartContext;
  title: string;
  newsList: string;
  eventsList: string;
  benefitsList: string;
  showNews: boolean;
  showEvents: boolean;
  showBenefits: boolean;
}
