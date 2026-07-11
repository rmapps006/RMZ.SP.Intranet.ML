import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHeroGreetingProps {
  context: WebPartContext;
  eyebrow: string;
  titleLine1: string;
  titleEmphasis: string;
  subtitle: string;
  buttonText: string;
  buttonUrl: string;
  showGreeting: boolean;
}
