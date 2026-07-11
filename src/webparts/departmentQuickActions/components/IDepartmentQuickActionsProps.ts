import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentQuickActionsProps {
  context: WebPartContext;
  actionsJson: string;
}
