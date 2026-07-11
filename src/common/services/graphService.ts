import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/** Returns the Microsoft Graph v3 client for the current SPFx context. */
export async function getGraphClient(context: WebPartContext): Promise<MSGraphClientV3> {
  return context.msGraphClientFactory.getClient('3');
}
