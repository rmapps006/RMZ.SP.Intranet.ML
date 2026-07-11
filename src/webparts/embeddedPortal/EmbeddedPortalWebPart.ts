import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmbeddedPortalWebPartStrings';
import EmbeddedPortal from './components/EmbeddedPortal';
import { IEmbeddedPortalProps } from './components/IEmbeddedPortalProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IEmbeddedPortalWebPartProps {
  portalUrl: string;
  frameHeight: string;
  sectionTitle: string;
  sectionTitleAR: string;
  linkText: string;
  linkTextAR: string;
  hideUrl: boolean;
  appSignInMode: boolean;
}

export default class EmbeddedPortalWebPart extends BaseClientSideWebPart<IEmbeddedPortalWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IEmbeddedPortalProps> = React.createElement(EmbeddedPortal, {
      context: this.context,
      portalUrl: this.properties.portalUrl,
      frameHeight: this.properties.frameHeight,
      sectionTitle: pickLocalized(this.properties.sectionTitle || 'Embedded Portal', this.properties.sectionTitleAR, language),
      linkText: pickLocalized(this.properties.linkText || 'Open in new tab', this.properties.linkTextAR, language),
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
                PropertyPaneTextField('sectionTitleAR', { label: strings.SectionTitleARLabel }),
                PropertyPaneTextField('linkText', { label: strings.LinkTextLabel }),
                PropertyPaneTextField('linkTextAR', { label: strings.LinkTextARLabel }),
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
