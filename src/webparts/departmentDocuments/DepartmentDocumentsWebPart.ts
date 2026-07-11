import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentDocumentsWebPartStrings';
import DepartmentDocuments from './components/DepartmentDocuments';
import { IDepartmentDocumentsProps } from './components/IDepartmentDocumentsProps';

export interface IDepartmentDocumentsWebPartProps {
  department: string;
  policiesLibrary: string;
  documentsLibrary: string;
  documentHubUrl: string;
  panel1Title: string;
  panel2Title: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class DepartmentDocumentsWebPart extends BaseClientSideWebPart<IDepartmentDocumentsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDepartmentDocumentsProps> = React.createElement(DepartmentDocuments, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      department: this.properties.department,
      policiesLibrary: this.properties.policiesLibrary,
      documentsLibrary: this.properties.documentsLibrary,
      documentHubUrl: this.properties.documentHubUrl,
      panel1Title: this.properties.panel1Title,
      panel2Title: this.properties.panel2Title
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
                PropertyPaneTextField('department', { label: strings.DepartmentLabel }),
                PropertyPaneTextField('policiesLibrary', { label: strings.PoliciesLibraryLabel }),
                PropertyPaneTextField('documentsLibrary', { label: strings.DocumentsLibraryLabel }),
                PropertyPaneTextField('documentHubUrl', { label: strings.DocumentHubUrlLabel }),
                PropertyPaneTextField('panel1Title', { label: strings.Panel1TitleLabel }),
                PropertyPaneTextField('panel2Title', { label: strings.Panel2TitleLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
