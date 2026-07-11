import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentHeroWebPartStrings';
import DepartmentHero from './components/DepartmentHero';
import { IDepartmentHeroProps } from './components/IDepartmentHeroProps';

export interface IDepartmentHeroWebPartProps {
  eyebrow: string;
  departmentName: string;
  description: string;
  ownerName: string;
  ownerRole: string;
}

export default class DepartmentHeroWebPart extends BaseClientSideWebPart<IDepartmentHeroWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDepartmentHeroProps> = React.createElement(DepartmentHero, {
      context: this.context,
      eyebrow: this.properties.eyebrow,
      departmentName: this.properties.departmentName,
      description: this.properties.description,
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
                PropertyPaneTextField('departmentName', { label: strings.DepartmentNameLabel }),
                PropertyPaneTextField('description', { label: strings.DescriptionLabel, multiline: true }),
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
