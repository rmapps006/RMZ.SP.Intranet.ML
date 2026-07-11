import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PoliciesListWebPartStrings';
import PoliciesList from './components/PoliciesList';
import { IPoliciesListProps } from './components/IPoliciesListProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IPoliciesListWebPartProps {
  policiesList: string;
  allPoliciesUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
  title: string;
  titleAR: string;
  linkText: string;
  linkTextAR: string;
}

export default class PoliciesListWebPart extends BaseClientSideWebPart<IPoliciesListWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IPoliciesListProps> = React.createElement(PoliciesList, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      policiesList: this.properties.policiesList || 'Policies',
      allPoliciesUrl: this.properties.allPoliciesUrl,
      title: pickLocalized(this.properties.title || 'Policies & Procedures', this.properties.titleAR, language),
      linkText: pickLocalized(this.properties.linkText || 'All Policies', this.properties.linkTextAR, language)
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
                PropertyPaneTextField('policiesList', { label: strings.PoliciesListLabel }),
                PropertyPaneTextField('allPoliciesUrl', { label: strings.AllPoliciesUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
