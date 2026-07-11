import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPoliciesListProps {
  context: WebPartContext;
  policiesList: string;
  allPoliciesUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}
