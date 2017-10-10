import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'receiverWpStrings';
import ReceiverWp from './components/ReceiverWp';
import { IReceiverWpProps } from './components/IReceiverWpProps';
import { IReceiverWpWebPartProps } from './IReceiverWpWebPartProps';

export default class ReceiverWpWebPart extends BaseClientSideWebPart<IReceiverWpWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReceiverWpProps > = React.createElement(
      ReceiverWp,
      {
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
