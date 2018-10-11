import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxGraphPlannerWebPartStrings';
import SpfxGraphPlanner from './components/SpfxGraphPlanner';
import { ISpfxGraphPlannerProps } from './components/ISpfxGraphPlannerProps';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface ISpfxGraphPlannerWebPartProps {
  description: string;
}

export default class SpfxGraphPlannerWebPart extends BaseClientSideWebPart<ISpfxGraphPlannerWebPartProps> {

  private myPlans: MicrosoftGraph.PlannerPlan[];

  @override
  public async onInit(): Promise<void> {
    await this.context.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      await client.api('me/planner/plans').get().then((data:any) => {
        this.myPlans = data.value;
      }).catch((error: any) => {
        console.log(error);
      });
    }).catch((error :any) => {
      console.log(error);
    });
  }

  public render(): void {
    const element: React.ReactElement<ISpfxGraphPlannerProps > = React.createElement(
      SpfxGraphPlanner,
      {
        msGraphClientFactory: this.context.msGraphClientFactory,
        myPlans: this.myPlans
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
