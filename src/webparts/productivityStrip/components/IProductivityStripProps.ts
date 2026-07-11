import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IProductivityStripProps {
  context: WebPartContext;
  quickLinksJson: string;
  maxEvents: string;
}
