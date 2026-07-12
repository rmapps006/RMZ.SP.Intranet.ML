import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmployeeDirectoryWebPartStrings';
import EmployeeDirectory from './components/EmployeeDirectory';
import { IEmployeeDirectoryProps } from './components/IEmployeeDirectoryProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IEmployeeDirectoryWebPartProps {
  title: string;
  titleAR: string;
  linkText: string;
  linkTextAR: string;
  fullDirectoryUrl: string;
  pageSize: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class EmployeeDirectoryWebPart extends BaseClientSideWebPart<IEmployeeDirectoryWebPartProps> {
  public render(): void {
    const parsed: number = parseInt(this.properties.pageSize, 10);
    const pageSize: number = isNaN(parsed) || parsed < 1 ? 24 : parsed;
    const language = getCurrentLanguage();

    const element: React.ReactElement<IEmployeeDirectoryProps> = React.createElement(EmployeeDirectory, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      title: pickLocalized(this.properties.title || 'Employee Directory', this.properties.titleAR, language),
      linkText: pickLocalized(this.properties.linkText || 'Full Directory', this.properties.linkTextAR || 'الدليل الكامل', language),
      fullDirectoryUrl: this.properties.fullDirectoryUrl,
      pageSize: pageSize
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
                PropertyPaneTextField('linkText', { label: strings.LinkTextLabel }),
                PropertyPaneTextField('linkTextAR', { label: strings.LinkTextARLabel }),
                PropertyPaneTextField('fullDirectoryUrl', { label: strings.FullDirectoryUrlLabel }),
                PropertyPaneTextField('pageSize', { label: strings.PageSizeLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
