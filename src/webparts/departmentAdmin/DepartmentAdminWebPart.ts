import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DepartmentAdminWebPartStrings';
import DepartmentAdmin from './components/DepartmentAdmin';
import { IDepartmentAdminProps } from './components/IDepartmentAdminProps';

export interface IDepartmentAdminWebPartProps {
  // No configurable properties — everything is edited on the web part surface.
  description?: string;
}

export default class DepartmentAdminWebPart extends BaseClientSideWebPart<IDepartmentAdminWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IDepartmentAdminProps> = React.createElement(DepartmentAdmin, {
      context: this.context
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
          groups: []
        }
      ]
    };
  }
}
