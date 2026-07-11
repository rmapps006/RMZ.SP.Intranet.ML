import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDetailViewProps {
  context: WebPartContext;
  newsList: string;
  eventsList: string;
  policiesList: string;
  benefitsList: string;
  backUrl: string;
}
