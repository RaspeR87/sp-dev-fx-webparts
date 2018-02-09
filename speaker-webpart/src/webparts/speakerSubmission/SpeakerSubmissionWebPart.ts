import * as React from 'react';
import * as ReactDom from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'SpeakerSubmissionWebPartStrings';
import SpeakerSubmission from './components/SpeakerSubmission';
import { ISpeakerSubmissionProps, ISpeakerSubmissionState } from './components/ISpeakerSubmissionProps';

import { DisplayMode } from '@microsoft/sp-core-library';
import * as SPTermStore from './components/SPTermStoreService'; 

const LOG_SOURCE: string = 'SpeakerSubmissionWebPart';

export interface ISpeakerSubmissionWebPartProps {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;

  countryTermsetName: string;
}

export default class SpeakerSubmissionWebPart extends BaseClientSideWebPart<ISpeakerSubmissionWebPartProps> {

  private _countryTermsetItems: SPTermStore.ISPTermObject[];

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized SpeakerSubmissionWebPart`);
    
    let termStoreService: SPTermStore.SPTermStoreService = new SPTermStore.SPTermStoreService({
      spHttpClient: this.context.spHttpClient,
      siteAbsoluteUrl: this.context.pageContext.web.absoluteUrl,
    });

    this._countryTermsetItems = await termStoreService.getTermsFromTermSetAsync(this.properties.countryTermsetName != null ? this.properties.countryTermsetName : "Country Termset");

    return Promise.resolve<void>();
  }

  public render(): void {
    const element: React.ReactElement<ISpeakerSubmissionProps> = React.createElement(
      SpeakerSubmission,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        countryItems: this._countryTermsetItems.map((i) => {
          return ({ 
            key: i.guid, 
            name: i.name 
          });
        })
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
                PropertyPaneTextField('countryTermsetName', {
                  label: strings.CountryTermsetNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
