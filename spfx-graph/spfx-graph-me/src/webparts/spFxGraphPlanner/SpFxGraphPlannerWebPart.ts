import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpFxGraphPlannerWebPartStrings';
import SpFxGraphPlanner from './components/SpFxGraphPlanner';
import { ISpFxGraphPlannerProps } from './components/ISpFxGraphPlannerProps';

import { MSGraphClient } from '@microsoft/sp-http';

export interface ISpFxGraphPlannerWebPartProps {
  description: string;
}

export default class SpFxGraphPlannerWebPart extends BaseClientSideWebPart<ISpFxGraphPlannerWebPartProps> {

  private myInformations: any;

  @override
  public async onInit(): Promise<void> {
    await this.context.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      await client.api('/me').get().then((value:any) => {
        this.myInformations = value;
      }).catch((error: any) => {
        console.log("Error: " + error);
      });
    }).catch((error :any) => {
      console.log("Error: " + error);
    });
  }

  public render(): void {
    const element: React.ReactElement<ISpFxGraphPlannerProps > = React.createElement(
      SpFxGraphPlanner,
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
