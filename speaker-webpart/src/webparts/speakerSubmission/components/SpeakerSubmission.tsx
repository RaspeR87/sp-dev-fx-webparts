import * as React from 'react';
import styles from './SpeakerSubmission.module.scss';
import { ISpeakerSubmissionProps, ISpeakerSubmissionState } from './ISpeakerSubmissionProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { TagPicker } from 'office-ui-fabric-react/lib/components/pickers/TagPicker/TagPicker';
import {
  Checkbox,
  ICheckboxProps
} from 'office-ui-fabric-react/lib/Checkbox';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import {
  Label,
  TextField,
  PrimaryButton,
  Button
} from 'office-ui-fabric-react';

export default class SpeakerSubmission extends React.Component<ISpeakerSubmissionProps, ISpeakerSubmissionState> {
  
  private sNameInput = null;
  private sLastName = null;
  private sEmail = null;
  private sCountry = null;
  private sShortBio = null;
  private sFile = null;

  constructor(props: ISpeakerSubmissionProps) {
    super(props);
    
    this.sNameInput = "";
    this.sLastName = "";
    this.sEmail = "";
    this.sFile = null;

    this.state = {
      isPickerDisabled: false,
      validateTextSName: "",
      validateTextSLastName: "",
      validateTextSEmail: "",
      validateTextSCountry: "",
      validateTextSShortBio: "",
      validateTextSFile: ""
    };
  }

  public isValidationOK(): boolean {
    if (this.state.validateTextSName == "" && this.state.validateTextSLastName == "" && this.state.validateTextSEmail == "" && 
      this.state.validateTextSCountry == "" && this.state.validateTextSShortBio == "") {
      return true;
    } else {
      return false;
    }
  }
  
  public render(): React.ReactElement<ISpeakerSubmissionProps> {
    return (
      <div className={ styles.speakerSubmission }>
        <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} />
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <TextField ref={(input) => { this.sNameInput = input; }} label='Name' name='SName' required={ true } autoComplete='on' errorMessage= { this.state.validateTextSName } />
              <TextField ref={(input) => { this.sLastName = input; }} label='Last Name' name='SLastName' required={ true } autoComplete='on' errorMessage= { this.state.validateTextSLastName } />
              <TextField ref={(input) => { this.sEmail = input; }} label='Email' name='SEmail' required={ true } autoComplete='on' onGetErrorMessage={ this.Email_CheckInput } errorMessage= { this.state.validateTextSEmail } />
              <Label required={ true }>Country</Label>
              <TagPicker ref={(input) => { this.sCountry = input; }}
                onResolveSuggestions={ this.Country_onFilterChanged }
                onChange={ this.Country_onChange }
                getTextFromItem={ this.Country_getTextFromItem }
                pickerSuggestionsProps={
                  {
                    suggestionsHeaderText: 'Suggested Items',
                    noResultsFoundText: 'No Items Found'
                  }
                }
                disabled={ this.state.isPickerDisabled }
                inputProps={ {
                  onBlur: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onBlur called'),
                  onFocus: (ev: React.FocusEvent<HTMLInputElement>) => console.log('onFocus called')
                } }
              />
              <span style={ this.state.validateTextSCountry == "" ? { display: 'none' } : { } }>
                <div aria-live="assertive" style={ { paddingBottom: 10 } }>
                  <p className="ms-TextField-errorMessage ms-u-slideDownIn20 errorMessage_2410e4c7" data-automation-id="error-message">
                    <i className="ms-Icon ms-Icon--Error errorIcon_2410e4c7" aria-hidden="true"></i>{ this.state.validateTextSCountry }
                  </p>
                </div>
              </span>
              <TextField label='Company' name='SCompany' required={ false } autoComplete='on' />
              <TextField label='Job Title' name='SJobTitle' required={ false } autoComplete='on' />
              <Label required={ false }>Credentials</Label>
              <Checkbox
                label='MCT'
                defaultChecked={ false }
              />
              <Checkbox
                label='MVP'
                defaultChecked={ false }
              />
              <Checkbox
                label='MCM/MCSM'
                defaultChecked={ false }
              />
              <ChoiceGroup
                options={ [
                  {
                    key: 'S',
                    text: 'S',
                  },
                  {
                    key: 'M',
                    text: 'M',
                  },
                  {
                    key: 'L',
                    text: 'L'
                  },
                  {
                    key: 'XL',
                    text: 'XL',
                  },
                  {
                    key: 'XXL',
                    text: 'XXL',
                  }
                ] }
                label='Shirt size'
                required={ false }
              />
              <TextField ref={(input) => { this.sShortBio = input; }} label='Short Bio' name='SShortBio' multiline autoAdjustHeight required={ true } errorMessage= { this.state.validateTextSShortBio } />
              <div style={ { paddingBottom: 10 } }>
                <Label required={ false }>Picture</Label>
                <input ref={ (input) => { this.sFile = input; } } type="file" />
                <span style={ this.state.validateTextSFile == "" ? { display: 'none' } : { } }>
                  <div aria-live="assertive">
                    <p className="ms-TextField-errorMessage ms-u-slideDownIn20 errorMessage_2410e4c7" data-automation-id="error-message">
                      <i className="ms-Icon ms-Icon--Error errorIcon_2410e4c7" aria-hidden="true"></i>{ this.state.validateTextSFile }
                    </p>
                  </div>
                </span>
              </div>
              <PrimaryButton type='submit' text='Submit' title='Submit' onClick={ () => { this.onSubmit(); } }></PrimaryButton>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private onSubmit():void {

    this.setState({
      isPickerDisabled: this.state.isPickerDisabled,
      validateTextSName: this.sNameInput.value.trim() == "" ? 'Input cannot be empty.' : '',
      validateTextSLastName: this.sLastName.value.trim() == "" ?  'Input cannot be empty.' : '',
      validateTextSEmail: this.sEmail.value.trim() == "" ?  'Input cannot be empty.' : '',
      validateTextSCountry: this.sCountry.items.length == 0 ? 'Input cannot be empty.' : '',
      validateTextSShortBio: this.sShortBio.value.trim() == "" ? 'Input cannot be empty.' : '',
      validateTextSFile: this.sFile.files.length == 0 ? 'Input cannot be empty.' : ''
    });

    if (this.isValidationOK()) {
      
    }
  }

  @autobind
  private Country_onChange(tagList: { key: string, name: string }[]) {
    if (tagList && tagList.length >= 1) {
      this.setState({
        isPickerDisabled: true,
        validateTextSName: this.state.validateTextSName,
        validateTextSLastName: this.state.validateTextSLastName,
        validateTextSEmail: this.state.validateTextSEmail,
        validateTextSCountry: this.state.validateTextSCountry,
        validateTextSShortBio: this.state.validateTextSCountry,
        validateTextSFile: this.state.validateTextSFile
      });
    } else {
      this.setState({
        isPickerDisabled: false,
        validateTextSName: this.state.validateTextSName,
        validateTextSLastName: this.state.validateTextSLastName,
        validateTextSEmail: this.state.validateTextSEmail,
        validateTextSCountry: this.state.validateTextSCountry,
        validateTextSShortBio: this.state.validateTextSCountry,
        validateTextSFile: this.state.validateTextSFile
      });
    }
  }

  @autobind
  private Country_onFilterChanged(filterText: string, tagList: { key: string, name: string }[]) {
    return filterText ? this.props.countryItems.filter(tag => tag.name.toLowerCase().indexOf(filterText.toLowerCase()) === 0).filter(item => !this._listContainsDocument(item, tagList)) : [];
  }

  private Country_getTextFromItem(item: any): any {
    return item.name;
  }

  private _listContainsDocument(tag: { key: string, name: string }, tagList: { key: string, name: string }[]) {
    if (!tagList || !tagList.length || tagList.length === 0) {
      return false;
    }
    return tagList.filter(compareTag => compareTag.key === tag.key).length > 0;
  }

  private TextField_EmptyValidator(value: string): string {
    return value.trim() == "" ? 'Input cannot be empty.' : '';
  }

  private Email_CheckInput(value: string): string {
    var re = /^(?:[a-z0-9!#$%&amp;'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&amp;'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])$/;
    return value.length == 0 || re.test(value) ? '' : 'Email is not in correct format.';
  }
}
