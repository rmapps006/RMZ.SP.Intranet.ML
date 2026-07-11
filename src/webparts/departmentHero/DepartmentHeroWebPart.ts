import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentHeroWebPartStrings';
import DepartmentHero from './components/DepartmentHero';
import { IDepartmentHeroProps } from './components/IDepartmentHeroProps';
import { getCurrentLanguage, pickLocalized } from '../../common/services/languageService';

export interface IDepartmentHeroWebPartProps {
  eyebrow: string;
  eyebrowAR: string;
  departmentName: string;
  departmentNameAR: string;
  description: string;
  descriptionAR: string;
  ownerName: string;
  ownerRole: string;
}

export default class DepartmentHeroWebPart extends BaseClientSideWebPart<IDepartmentHeroWebPartProps> {
  public render(): void {
    const language = getCurrentLanguage();
    const element: React.ReactElement<IDepartmentHeroProps> = React.createElement(DepartmentHero, {
      context: this.context,
      eyebrow: pickLocalized(this.properties.eyebrow, this.properties.eyebrowAR, language),
      departmentName: pickLocalized(this.properties.departmentName, this.properties.departmentNameAR, language),
      description: pickLocalized(this.properties.description, this.properties.descriptionAR, language),
      ownerName: this.properties.ownerName,
      ownerRole: this.properties.ownerRole
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
                PropertyPaneTextField('eyebrow', { label: strings.EyebrowLabel }),
                PropertyPaneTextField('eyebrowAR', { label: strings.EyebrowARLabel }),
                PropertyPaneTextField('departmentName', { label: strings.DepartmentNameLabel }),
                PropertyPaneTextField('departmentNameAR', { label: strings.DepartmentNameARLabel }),
                PropertyPaneTextField('description', { label: strings.DescriptionLabel, multiline: true }),
                PropertyPaneTextField('descriptionAR', { label: strings.DescriptionARLabel, multiline: true }),
                PropertyPaneTextField('ownerName', { label: strings.OwnerNameLabel }),
                PropertyPaneTextField('ownerRole', { label: strings.OwnerRoleLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
