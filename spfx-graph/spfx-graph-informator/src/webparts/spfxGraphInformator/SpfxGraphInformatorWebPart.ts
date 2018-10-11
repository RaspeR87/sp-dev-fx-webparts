import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxGraphInformatorWebPartStrings';
import SpfxGraphInformator from './components/SpfxGraphInformator';
import { ISpfxGraphInformatorProps } from './components/ISpfxGraphInformatorProps';

import { MSGraphClient } from '@microsoft/sp-http';

import * as moment from 'moment';

export interface ISpfxGraphInformatorWebPartProps {
  description: string;
}

export default class SpfxGraphInformatorWebPart extends BaseClientSideWebPart<ISpfxGraphInformatorWebPartProps> {

  private nrNewMails: Number;
  private nrTotalMails: Number;
  private nrUpcomingEvents: Number;

  @override
  public async onInit(): Promise<void> {
    await this.context.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      await client.api('/me/mailfolders/Inbox').get().then((value:any) => {
        this.nrNewMails = value.unreadItemCount;
        this.nrTotalMails = value.totalItemCount;
      }).catch((error: any) => {
        console.log(error);
      });
    }).catch((error :any) => {
      console.log(error);
    });

    var fromDate = moment();
    var toDate = moment().add(7, 'days');

    await this.context.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      await client.api('/me/calendarview?startdatetime=' + fromDate.format("YYYY-MM-DDTHH:mm:ss") + 'Z&enddatetime=' + toDate.format("YYYY-MM-DDTHH:mm:ss") + 'Z').get().then((data:any) => {
        this.nrUpcomingEvents = data.value.length;
      }).catch((error: any) => {
        console.log(error);
      });
    }).catch((error :any) => {
      console.log(error);
    });
  }

  public render(): void {
    const element: React.ReactElement<ISpfxGraphInformatorProps > = React.createElement(
      SpfxGraphInformator,
      {
        nrNewMails: this.nrNewMails,
        nrTotalMails: this.nrTotalMails,
        nrUpcomingEvents: this.nrUpcomingEvents
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
