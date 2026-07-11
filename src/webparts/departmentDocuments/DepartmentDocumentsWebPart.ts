import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentDocumentsWebPartStrings';
import DepartmentDocuments from './components/DepartmentDocuments';
import { IDepartmentDocumentsProps } from './components/IDepartmentDocumentsProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IDepartmentDocumentsWebPartProps {
  title: string;
  titleAR: string;
  linkText: string;
  linkTextAR: string;
  department: string;
  policiesLibrary: string;
  documentsLibrary: string;
  documentHubUrl: string;
  panel1Title: string;
  panel1TitleAR: string;
  panel2Title: string;
  panel2TitleAR: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class DepartmentDocumentsWebPart extends BaseClientSideWebPart<IDepartmentDocumentsWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IDepartmentDocumentsProps> = React.createElement(DepartmentDocuments, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      title: pickLocalized(this.properties.title, this.properties.titleAR, language),
      linkText: pickLocalized(this.properties.linkText, this.properties.linkTextAR, language),
      department: this.properties.department,
      policiesLibrary: this.properties.policiesLibrary,
      documentsLibrary: this.properties.documentsLibrary,
      documentHubUrl: this.properties.documentHubUrl,
      // Resolves this web part's own property-pane override only. The
      // settings-backed fallback (Department Admin's ds.panel1Title/panel2Title,
      // applied in the component when this resolves empty) still has no Arabic
      // variant — known follow-up, out of scope here.
      panel1Title: pickLocalized(this.properties.panel1Title, this.properties.panel1TitleAR, language),
      panel2Title: pickLocalized(this.properties.panel2Title, this.properties.panel2TitleAR, language)
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
                PropertyPaneTextField('department', { label: strings.DepartmentLabel }),
                PropertyPaneTextField('policiesLibrary', { label: strings.PoliciesLibraryLabel }),
                PropertyPaneTextField('documentsLibrary', { label: strings.DocumentsLibraryLabel }),
                PropertyPaneTextField('documentHubUrl', { label: strings.DocumentHubUrlLabel }),
                PropertyPaneTextField('panel1Title', { label: strings.Panel1TitleLabel }),
                PropertyPaneTextField('panel1TitleAR', { label: strings.Panel1TitleARLabel }),
                PropertyPaneTextField('panel2Title', { label: strings.Panel2TitleLabel }),
                PropertyPaneTextField('panel2TitleAR', { label: strings.Panel2TitleARLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
