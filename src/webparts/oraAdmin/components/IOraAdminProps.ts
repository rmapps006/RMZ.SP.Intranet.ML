import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IOraAdminProps {
  context: WebPartContext;
  eventsList: string;
  policiesList: string;
  newsList: string;
  benefitsList: string;
  formsLibrary: string;
  templatesLibrary: string;
  pagesLibrary: string;
  registerHeaderFooter: boolean;
  seedSampleData: boolean;
  createViewAllPages: boolean;
}
