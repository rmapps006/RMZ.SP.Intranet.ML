import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'HrBenefitsWebPartStrings';
import HrBenefits from './components/HrBenefits';
import { IHrBenefitsProps } from './components/IHrBenefitsProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IHrBenefitsWebPartProps {
  title: string;
  titleAR: string;
  allBenefitsUrl: string;
  benefitsList: string;
  benefitsJson: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class HrBenefitsWebPart extends BaseClientSideWebPart<IHrBenefitsWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IHrBenefitsProps> = React.createElement(HrBenefits, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      title: pickLocalized(this.properties.title || 'HR Benefits', this.properties.titleAR, language),
      allBenefitsUrl: this.properties.allBenefitsUrl,
      benefitsList: this.properties.benefitsList || 'HR Benefits',
      benefitsJson: this.properties.benefitsJson
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
                PropertyPaneTextField('allBenefitsUrl', { label: strings.AllBenefitsUrlLabel }),
                PropertyPaneTextField('benefitsList', { label: 'HR Benefits list name' }),
                PropertyPaneTextField('benefitsJson', { label: strings.BenefitsJsonLabel, multiline: true })
              ]
            }
          ]
        }
      ]
    };
  }
}
