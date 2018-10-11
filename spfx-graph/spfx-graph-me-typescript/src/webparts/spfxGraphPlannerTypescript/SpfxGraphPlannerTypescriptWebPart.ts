import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxGraphPlannerTypescriptWebPartStrings';
import SpfxGraphPlannerTypescript from './components/SpfxGraphPlannerTypescript';
import { ISpfxGraphPlannerTypescriptProps } from './components/ISpfxGraphPlannerTypescriptProps';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface ISpfxGraphPlannerTypescriptWebPartProps {
  description: string;
}

export default class SpfxGraphPlannerTypescriptWebPart extends BaseClientSideWebPart<ISpfxGraphPlannerTypescriptWebPartProps> {

  private myInformations: MicrosoftGraph.User;

  @override
  public async onInit(): Promise<void> {
    await this.context.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      await client.api('/me').get().then((value:MicrosoftGraph.User) => {
        this.myInformations = value;
      }).catch((error: any) => {
        console.log("Error: " + error);
      });
    }).catch((error :any) => {
      console.log("Error: " + error);
    });
  }

  public render(): void {
    const element: React.ReactElement<ISpfxGraphPlannerTypescriptProps > = React.createElement(
      SpfxGraphPlannerTypescript,
      {
        myInformations: this.myInformations
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
