import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDepartmentHeroProps {
  context: WebPartContext;
  eyebrow: string;
  departmentName: string;
  description: string;
  ownerName: string;
  ownerRole: string;
}
