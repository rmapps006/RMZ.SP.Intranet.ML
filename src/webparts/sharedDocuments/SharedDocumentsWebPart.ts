import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SharedDocumentsWebPartStrings';
import SharedDocuments from './components/SharedDocuments';
import { ISharedDocumentsProps } from './components/ISharedDocumentsProps';

export interface ISharedDocumentsWebPartProps {
  panel1Library: string;
  panel1Title: string;
  panel1Url: string;
  panel2Library: string;
  panel2Title: string;
  panel2Url: string;
  documentHubUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class SharedDocumentsWebPart extends BaseClientSideWebPart<ISharedDocumentsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ISharedDocumentsProps> = React.createElement(SharedDocuments, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      panel1Library: this.properties.panel1Library,
      panel1Title: this.properties.panel1Title,
      panel1Url: this.properties.panel1Url,
      panel2Library: this.properties.panel2Library,
      panel2Title: this.properties.panel2Title,
      panel2Url: this.properties.panel2Url,
      documentHubUrl: this.properties.documentHubUrl
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
                PropertyPaneTextField('panel1Library', { label: strings.Panel1LibraryLabel }),
                PropertyPaneTextField('panel1Title', { label: strings.Panel1TitleLabel }),
                PropertyPaneTextField('panel1Url', { label: strings.Panel1UrlLabel }),
                PropertyPaneTextField('panel2Library', { label: strings.Panel2LibraryLabel }),
                PropertyPaneTextField('panel2Title', { label: strings.Panel2TitleLabel }),
                PropertyPaneTextField('panel2Url', { label: strings.Panel2UrlLabel }),
                PropertyPaneTextField('documentHubUrl', { label: strings.DocumentHubUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
