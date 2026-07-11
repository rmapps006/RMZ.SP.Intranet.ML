import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentEventsWebPartStrings';
import DepartmentEvents from './components/DepartmentEvents';
import { IDepartmentEventsProps } from './components/IDepartmentEventsProps';

export interface IDepartmentEventsWebPartProps {
  eventsList: string;
  calendarUrl: string;
  maxItems: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class DepartmentEventsWebPart extends BaseClientSideWebPart<IDepartmentEventsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDepartmentEventsProps> = React.createElement(DepartmentEvents, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      eventsList: this.properties.eventsList,
      calendarUrl: this.properties.calendarUrl,
      maxItems: this.properties.maxItems
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
                PropertyPaneTextField('eventsList', { label: strings.EventsListLabel }),
                PropertyPaneTextField('calendarUrl', { label: strings.CalendarUrlLabel }),
                PropertyPaneTextField('maxItems', { label: strings.MaxItemsLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
