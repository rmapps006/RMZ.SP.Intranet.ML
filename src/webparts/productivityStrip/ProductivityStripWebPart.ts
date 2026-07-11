import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ProductivityStripWebPartStrings';
import ProductivityStrip from './components/ProductivityStrip';
import { IProductivityStripProps } from './components/IProductivityStripProps';

export interface IProductivityStripWebPartProps {
  quickLinksJson: string;
  maxEvents: string;
}

export default class ProductivityStripWebPart extends BaseClientSideWebPart<IProductivityStripWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IProductivityStripProps> = React.createElement(ProductivityStrip, {
      context: this.context,
      quickLinksJson: this.properties.quickLinksJson,
      maxEvents: this.properties.maxEvents
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
                PropertyPaneTextField('maxEvents', { label: strings.MaxEventsLabel }),
                PropertyPaneTextField('quickLinksJson', {
                  label: strings.QuickLinksJsonLabel,
                  multiline: true,
                  rows: 8
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
