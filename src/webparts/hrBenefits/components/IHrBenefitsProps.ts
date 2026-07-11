import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHrBenefitsProps {
  context: WebPartContext;
  title: string;
  allBenefitsUrl: string;
  benefitsList: string;
  benefitsJson: string;
  showTitle: boolean;
  showViewAll: boolean;
}
