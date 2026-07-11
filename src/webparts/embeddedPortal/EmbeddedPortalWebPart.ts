import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmbeddedPortalWebPartStrings';
import EmbeddedPortal from './components/EmbeddedPortal';
import { IEmbeddedPortalProps } from './components/IEmbeddedPortalProps';

export interface IEmbeddedPortalWebPartProps {
  portalUrl: string;
  frameHeight: string;
  sectionTitle: string;
  linkText: string;
  hideUrl: boolean;
  appSignInMode: boolean;
}

export default class EmbeddedPortalWebPart extends BaseClientSideWebPart<IEmbeddedPortalWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IEmbeddedPortalProps> = React.createElement(EmbeddedPortal, {
      context: this.context,
      portalUrl: this.properties.portalUrl,
      frameHeight: this.properties.frameHeight,
      sectionTitle: this.properties.sectionTitle,
      linkText: this.properties.linkText,
      hideUrl: this.properties.hideUrl === true,
      appSignInMode: this.properties.appSignInMode === true
    });
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('sectionTitle', { label: strings.SectionTitleLabel }),
                PropertyPaneTextField('linkText', { label: strings.LinkTextLabel }),
                PropertyPaneTextField('portalUrl', { label: strings.PortalUrlLabel }),
                PropertyPaneTextField('frameHeight', { label: strings.FrameHeightLabel }),
                PropertyPaneToggle('hideUrl', { label: strings.HideUrlLabel }),
                PropertyPaneToggle('appSignInMode', { label: strings.AppSignInModeLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
