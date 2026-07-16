import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'IntranetAdminWebPartStrings';
import IntranetAdmin from './components/IntranetAdmin';
import { IIntranetAdminProps } from './components/IIntranetAdminProps';

export interface IIntranetAdminWebPartProps {
  eventsList: string;
  policiesList: string;
  newsList: string;
  benefitsList: string;
  docCenterLibrary: string;
  formsLibrary: string;
  templatesLibrary: string;
  pagesLibrary: string;
  registerHeaderFooter: boolean;
  seedSampleData: boolean;
  createViewAllPages: boolean;
}

export default class IntranetAdminWebPart extends BaseClientSideWebPart<IIntranetAdminWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IIntranetAdminProps> = React.createElement(IntranetAdmin, {
      context: this.context,
      eventsList: this.properties.eventsList || 'Events',
      policiesList: this.properties.policiesList || 'Policies',
      newsList: this.properties.newsList || 'News',
      benefitsList: this.properties.benefitsList || 'HR Benefits',
      docCenterLibrary: this.properties.docCenterLibrary || 'Document Center',
      formsLibrary: this.properties.formsLibrary || 'Forms',
      templatesLibrary: this.properties.templatesLibrary || 'Templates',
      pagesLibrary: this.properties.pagesLibrary || 'Site Pages',
      registerHeaderFooter: this.properties.registerHeaderFooter !== false,
      seedSampleData: this.properties.seedSampleData === true,
      createViewAllPages: this.properties.createViewAllPages !== false
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
                PropertyPaneTextField('eventsList', { label: strings.EventsListLabel }),
                PropertyPaneTextField('policiesList', { label: strings.PoliciesListLabel }),
                PropertyPaneTextField('newsList', { label: strings.NewsListLabel }),
                PropertyPaneTextField('benefitsList', { label: strings.BenefitsListLabel }),
                PropertyPaneTextField('docCenterLibrary', { label: strings.DocCenterLibraryLabel }),
                PropertyPaneTextField('formsLibrary', { label: strings.FormsLibraryLabel }),
                PropertyPaneTextField('templatesLibrary', { label: strings.TemplatesLibraryLabel }),
                PropertyPaneTextField('pagesLibrary', { label: strings.PagesLibraryLabel }),
                PropertyPaneToggle('registerHeaderFooter', { label: strings.RegisterHeaderFooterLabel }),
                PropertyPaneToggle('seedSampleData', { label: strings.SeedSampleDataLabel }),
                PropertyPaneToggle('createViewAllPages', { label: 'Create View All pages on setup' })
              ]
            }
          ]
        }
      ]
    };
  }
}
