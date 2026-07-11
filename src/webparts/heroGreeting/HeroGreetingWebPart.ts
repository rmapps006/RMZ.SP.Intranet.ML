import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HeroGreetingWebPartStrings';
import HeroGreeting from './components/HeroGreeting';
import { IHeroGreetingProps } from './components/IHeroGreetingProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IHeroGreetingWebPartProps {
  eyebrow: string;
  eyebrowAR: string;
  titleLine1: string;
  titleLine1AR: string;
  titleEmphasis: string;
  titleEmphasisAR: string;
  subtitle: string;
  subtitleAR: string;
  buttonText: string;
  buttonTextAR: string;
  buttonUrl: string;
  showGreeting: boolean;
}

export default class HeroGreetingWebPart extends BaseClientSideWebPart<IHeroGreetingWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IHeroGreetingProps> = React.createElement(HeroGreeting, {
      context: this.context,
      eyebrow: pickLocalized(this.properties.eyebrow, this.properties.eyebrowAR, language),
      titleLine1: pickLocalized(this.properties.titleLine1, this.properties.titleLine1AR, language),
      titleEmphasis: pickLocalized(this.properties.titleEmphasis, this.properties.titleEmphasisAR, language),
      subtitle: pickLocalized(this.properties.subtitle, this.properties.subtitleAR, language),
      buttonText: pickLocalized(this.properties.buttonText, this.properties.buttonTextAR, language),
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
                PropertyPaneTextField('eyebrowAR', { label: strings.EyebrowARLabel }),
                PropertyPaneTextField('titleLine1', { label: strings.TitleLine1Label }),
                PropertyPaneTextField('titleLine1AR', { label: strings.TitleLine1ARLabel }),
                PropertyPaneTextField('titleEmphasis', { label: strings.TitleEmphasisLabel }),
                PropertyPaneTextField('titleEmphasisAR', { label: strings.TitleEmphasisARLabel }),
                PropertyPaneTextField('subtitle', { label: strings.SubtitleLabel, multiline: true }),
                PropertyPaneTextField('subtitleAR', { label: strings.SubtitleARLabel, multiline: true }),
                PropertyPaneTextField('buttonText', { label: strings.ButtonTextLabel }),
                PropertyPaneTextField('buttonTextAR', { label: strings.ButtonTextARLabel }),
                PropertyPaneTextField('buttonUrl', { label: strings.ButtonUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
