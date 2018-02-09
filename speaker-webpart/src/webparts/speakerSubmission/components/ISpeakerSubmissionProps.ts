import { DisplayMode } from '@microsoft/sp-core-library';
import * as SPTermStore from './SPTermStoreService'; 

export interface ISpeakerSubmissionProps {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;

  countryItems: {key:string; name:string;}[];
}

export interface ISpeakerSubmissionState {
  isPickerDisabled: boolean;

  validateTextSName: string;
  validateTextSLastName: string;
  validateTextSEmail: string;
  validateTextSCountry: string;
  validateTextSShortBio: string;
  validateTextSFile: string;
}