import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ContentManagerWebPartStrings';
import ContentManager from './components/ContentManager';
import { IContentManagerProps } from './components/IContentManagerProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IContentManagerWebPartProps {
  title: string;
  titleAR: string;
  newsList: string;
  eventsList: string;
  benefitsList: string;
  showNews: boolean;
  showEvents: boolean;
  showBenefits: boolean;
}

export default class ContentManagerWebPart extends BaseClientSideWebPart<IContentManagerWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IContentManagerProps> = React.createElement(ContentManager, {
      context: this.context,
      title: pickLocalized(this.properties.title || 'Content Manager', this.properties.titleAR, language),
      newsList: this.properties.newsList || 'News',
      eventsList: this.properties.eventsList || 'Events',
      benefitsList: this.properties.benefitsList || 'HR Benefits',
      showNews: this.properties.showNews !== false,
      showEvents: this.properties.showEvents !== false,
      // Benefits off by default so a department site shows only News + Events.
      showBenefits: this.properties.showBenefits === true
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
                PropertyPaneTextField('title', { label: strings.TitleLabel }),
                PropertyPaneTextField('titleAR', { label: strings.TitleARLabel }),
                PropertyPaneToggle('showNews', { label: strings.ShowNewsLabel }),
                PropertyPaneToggle('showEvents', { label: strings.ShowEventsLabel }),
                PropertyPaneToggle('showBenefits', { label: strings.ShowBenefitsLabel }),
                PropertyPaneTextField('newsList', { label: strings.NewsListLabel }),
                PropertyPaneTextField('eventsList', { label: strings.EventsListLabel }),
                PropertyPaneTextField('benefitsList', { label: strings.BenefitsListLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
