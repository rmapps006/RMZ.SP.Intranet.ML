import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentCenterWebPartStrings';
import DocumentCenter from './components/DocumentCenter';
import { IDocumentCenterProps } from './components/IDocumentCenterProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IDocumentCenterWebPartProps {
  libraryTitle: string;
  siteUrl: string;
  pageSize: string;
  showTitle: boolean;
  showViewAll: boolean;
  title: string;
  titleAR: string;
  linkText: string;
  linkTextAR: string;
}

export default class DocumentCenterWebPart extends BaseClientSideWebPart<IDocumentCenterWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const parsed: number = parseInt(this.properties.pageSize, 10);
    const pageSize: number = isNaN(parsed) || parsed < 1 ? 50 : parsed;

    const element: React.ReactElement<IDocumentCenterProps> = React.createElement(DocumentCenter, {
      context: this.context,
      libraryTitle: this.properties.libraryTitle || 'Document Center',
      siteUrl: (this.properties.siteUrl || '').trim(),
      pageSize: pageSize,
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      title: pickLocalized(this.properties.title || 'Document Center', this.properties.titleAR, language),
      linkText: pickLocalized(this.properties.linkText || 'Open Library', this.properties.linkTextAR || 'فتح المكتبة', language)
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
                PropertyPaneToggle('showViewAll', { label: 'Show the "Open Library" link' }),
                PropertyPaneTextField('title', { label: strings.TitleLabel }),
                PropertyPaneTextField('titleAR', { label: strings.TitleARLabel }),
                PropertyPaneTextField('linkText', { label: strings.LinkTextLabel }),
                PropertyPaneTextField('linkTextAR', { label: strings.LinkTextARLabel }),
                PropertyPaneTextField('libraryTitle', { label: strings.LibraryTitleLabel }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlLabel,
                  description: strings.SiteUrlDescription
                }),
                PropertyPaneTextField('pageSize', { label: strings.PageSizeLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
