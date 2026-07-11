import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EventsListWebPartStrings';
import EventsList from './components/EventsList';
import { IEventsListProps } from './components/IEventsListProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IEventsListWebPartProps {
  eventsList: string;
  calendarUrl: string;
  maxItems: string;
  showTitle: boolean;
  showViewAll: boolean;
  title: string;
  titleAR: string;
  linkText: string;
  linkTextAR: string;
}

export default class EventsListWebPart extends BaseClientSideWebPart<IEventsListWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IEventsListProps> = React.createElement(EventsList, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      eventsList: this.properties.eventsList,
      calendarUrl: this.properties.calendarUrl,
      maxItems: this.properties.maxItems,
      title: pickLocalized(this.properties.title || 'Company Events', this.properties.titleAR, language),
      linkText: pickLocalized(this.properties.linkText || 'View Calendar', this.properties.linkTextAR, language)
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
