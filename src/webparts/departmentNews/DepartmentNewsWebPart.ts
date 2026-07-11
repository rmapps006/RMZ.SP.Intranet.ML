import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentNewsWebPartStrings';
import DepartmentNews from './components/DepartmentNews';
import { IDepartmentNewsProps } from './components/IDepartmentNewsProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IDepartmentNewsWebPartProps {
  title: string;
  titleAR: string;
  linkText: string;
  linkTextAR: string;
  newsList: string;
  allNewsUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class DepartmentNewsWebPart extends BaseClientSideWebPart<IDepartmentNewsWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IDepartmentNewsProps> = React.createElement(DepartmentNews, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      title: pickLocalized(this.properties.title, this.properties.titleAR, language),
      linkText: pickLocalized(this.properties.linkText, this.properties.linkTextAR, language),
      newsList: this.properties.newsList,
      allNewsUrl: this.properties.allNewsUrl
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
                PropertyPaneTextField('newsList', { label: strings.NewsListLabel }),
                PropertyPaneTextField('allNewsUrl', { label: strings.AllNewsUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
