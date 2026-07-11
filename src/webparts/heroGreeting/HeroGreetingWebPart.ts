import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HeroGreetingWebPartStrings';
import HeroGreeting from './components/HeroGreeting';
import { IHeroGreetingProps } from './components/IHeroGreetingProps';

export interface IHeroGreetingWebPartProps {
  eyebrow: string;
  titleLine1: string;
  titleEmphasis: string;
  subtitle: string;
  buttonText: string;
  buttonUrl: string;
  showGreeting: boolean;
}

export default class HeroGreetingWebPart extends BaseClientSideWebPart<IHeroGreetingWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IHeroGreetingProps> = React.createElement(HeroGreeting, {
      context: this.context,
      eyebrow: this.properties.eyebrow,
      titleLine1: this.properties.titleLine1,
      titleEmphasis: this.properties.titleEmphasis,
      subtitle: this.properties.subtitle,
      buttonText: this.properties.buttonText,
      buttonUrl: this.properties.buttonUrl,
      showGreeting: this.properties.showGreeting !== false
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
                PropertyPaneToggle('showGreeting', { label: strings.ShowGreetingLabel }),
                PropertyPaneTextField('eyebrow', { label: strings.EyebrowLabel }),
                PropertyPaneTextField('titleLine1', { label: strings.TitleLine1Label }),
                PropertyPaneTextField('titleEmphasis', { label: strings.TitleEmphasisLabel }),
                PropertyPaneTextField('subtitle', { label: strings.SubtitleLabel, multiline: true }),
                PropertyPaneTextField('buttonText', { label: strings.ButtonTextLabel }),
                PropertyPaneTextField('buttonUrl', { label: strings.ButtonUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
