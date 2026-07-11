import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ApprovalsWebPartStrings';
import Approvals from './components/Approvals';
import { IApprovalsProps } from './components/IApprovalsProps';

export interface IApprovalsWebPartProps {
  requestsList: string;
  department: string;
  viewAllUrl: string;
  newRequestUrl: string;
  reviewAllUrl: string;
  showTitle: boolean;
  showViewAll: boolean;
}

export default class ApprovalsWebPart extends BaseClientSideWebPart<IApprovalsWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IApprovalsProps> = React.createElement(Approvals, {
      showTitle: this.properties.showTitle !== false,
      showViewAll: this.properties.showViewAll !== false,
      context: this.context,
      requestsList: this.properties.requestsList,
      department: this.properties.department,
      viewAllUrl: this.properties.viewAllUrl,
      newRequestUrl: this.properties.newRequestUrl,
      reviewAllUrl: this.properties.reviewAllUrl
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
                PropertyPaneTextField('requestsList', { label: strings.RequestsListLabel }),
                PropertyPaneTextField('department', { label: strings.DepartmentLabel }),
                PropertyPaneTextField('viewAllUrl', { label: strings.ViewAllUrlLabel }),
                PropertyPaneTextField('newRequestUrl', { label: strings.NewRequestUrlLabel }),
                PropertyPaneTextField('reviewAllUrl', { label: strings.ReviewAllUrlLabel })
              ]
            }
          ]
        }
      ]
    };
  }
}
