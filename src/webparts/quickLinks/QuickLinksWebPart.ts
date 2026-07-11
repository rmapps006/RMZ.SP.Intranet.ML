import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'QuickLinksWebPartStrings';
import QuickLinks from './components/QuickLinks';
import { IQuickLinksProps } from './components/IQuickLinksProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IQuickLinksWebPartProps {
  title: string;
  titleAR: string;
  viewAllUrl: string;
  linksJson: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart<IQuickLinksWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(QuickLinks, {
      title: pickLocalized(this.properties.title || 'Quick Links', this.properties.titleAR, language),
      viewAllUrl: this.properties.viewAllUrl || '',
      linksJson: this.properties.linksJson || '',
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false
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
                PropertyPaneToggle('showTitle', { label: 'Show section title' }),
                PropertyPaneToggle('showViewAll', { label: 'Show the "View All" link' }),
                PropertyPaneTextField('title', { label: strings.TitleLabel }),
                PropertyPaneTextField('titleAR', { label: strings.TitleARLabel }),
                PropertyPaneTextField('viewAllUrl', { label: strings.ViewAllUrlLabel }),
                PropertyPaneTextField('linksJson', {
                  label: strings.LinksJsonLabel,
                  multiline: true,
                  rows: 8
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
