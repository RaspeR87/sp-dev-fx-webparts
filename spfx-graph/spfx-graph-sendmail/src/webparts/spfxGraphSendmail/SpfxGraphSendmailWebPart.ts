import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxGraphSendmailWebPartStrings';
import SpfxGraphSendmail from './components/SpfxGraphSendmail';
import { ISpfxGraphSendmailProps } from './components/ISpfxGraphSendmailProps';

export interface ISpfxGraphSendmailWebPartProps {
  description: string;
}

export default class SpfxGraphSendmailWebPart extends BaseClientSideWebPart<ISpfxGraphSendmailWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxGraphSendmailProps > = React.createElement(
      SpfxGraphSendmail,
      {
        msGraphClientFactory: this.context.msGraphClientFactory
      }
    );

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
