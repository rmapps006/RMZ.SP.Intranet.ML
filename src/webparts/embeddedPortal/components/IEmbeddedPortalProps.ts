import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEmbeddedPortalProps {
  context: WebPartContext;
  portalUrl: string;
  frameHeight: string;
  sectionTitle: string;
  linkText: string;
  /** Hide the URL pill in the browser chrome of the embedded frame. */
  hideUrl: boolean;
  /**
   * Relax the iframe sandbox so a trusted embedded app can run its own sign-in
   * (keep its session cookie, open the IdP popup, request storage access).
   * Off by default — only enable for first-party apps you control.
   */
  appSignInMode: boolean;
}
