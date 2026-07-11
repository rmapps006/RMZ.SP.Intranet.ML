import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DetailViewWebPartStrings';
import DetailView from './components/DetailView';
import { IDetailViewProps } from './components/IDetailViewProps';

export interface IDetailViewWebPartProps {
  newsList: string;
  eventsList: string;
  policiesList: string;
  benefitsList: string;
  backUrl: string;
}

export default class DetailViewWebPart extends BaseClientSideWebPart<IDetailViewWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDetailViewProps> = React.createElement(DetailView, {
      context: this.context,
      newsList: this.properties.newsList || 'News',
      eventsList: this.properties.eventsList || 'Events',
      policiesList: this.properties.policiesList || 'Policies',
      benefitsList: this.properties.benefitsList || 'HR Benefits',
      backUrl: this.properties.backUrl || ''
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
                PropertyPaneTextField('newsList', { label: strings.NewsListLabel }),
                PropertyPaneTextField('eventsList', { label: strings.EventsListLabel }),
                PropertyPaneTextField('policiesList', { label: strings.PoliciesListLabel }),
                PropertyPaneTextField('benefitsList', { label: 'HR Benefits list name' }),
                PropertyPaneTextField('backUrl', { label: strings.BackUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
