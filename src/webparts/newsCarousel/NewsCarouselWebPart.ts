import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'NewsCarouselWebPartStrings';
import NewsCarousel from './components/NewsCarousel';
import { INewsCarouselProps } from './components/INewsCarouselProps';

export interface INewsCarouselWebPartProps {
  newsList: string;
  archiveUrl: string;
  maxItems: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class NewsCarouselWebPart extends BaseClientSideWebPart<INewsCarouselWebPartProps> {
  public render(): void {
    const element: React.ReactElement<INewsCarouselProps> = React.createElement(NewsCarousel, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      newsList: this.properties.newsList,
      archiveUrl: this.properties.archiveUrl,
      maxItems: parseInt(this.properties.maxItems, 10) || 3
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
                PropertyPaneTextField('newsList', { label: strings.NewsListLabel }),
                PropertyPaneTextField('archiveUrl', { label: strings.ArchiveUrlLabel }),
                PropertyPaneTextField('maxItems', { label: strings.MaxItemsLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
