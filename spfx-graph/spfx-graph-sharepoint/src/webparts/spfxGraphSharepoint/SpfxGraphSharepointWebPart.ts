import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxGraphSharepointWebPartStrings';
import SpfxGraphSharepoint from './components/SpfxGraphSharepoint';
import { ISpfxGraphSharepointProps } from './components/ISpfxGraphSharepointProps';

import { HttpClient } from '@microsoft/sp-http';

export interface ISpfxGraphSharepointWebPartProps {
  description: string;
}

export default class SpfxGraphSharepointWebPart extends BaseClientSideWebPart<ISpfxGraphSharepointWebPartProps> {

  public render(): void {
    
    const element: React.ReactElement<ISpfxGraphSharepointProps > = React.createElement(
      SpfxGraphSharepoint,
      {
        msGraphClientFactory: this.context.msGraphClientFactory,
        currWebUrl: this.context.pageContext.web.serverRelativeUrl
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
