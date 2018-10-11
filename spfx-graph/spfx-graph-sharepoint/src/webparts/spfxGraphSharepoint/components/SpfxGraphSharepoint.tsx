import * as React from 'react';
import styles from './SpfxGraphSharepoint.module.scss';
import { ISpfxGraphSharepointProps, ISpfxGraphSharepointState } from './ISpfxGraphSharepointProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import { MSGraphClient, HttpClient } from '@microsoft/sp-http';
import Dropzone from 'react-dropzone';

export default class SpfxGraphSharepoint extends React.Component<ISpfxGraphSharepointProps, ISpfxGraphSharepointState> {
  
  constructor(props: ISpfxGraphSharepointProps) {
    super(props);

    this.state = {
      searchFor: "",
      results: [],
      cTs: []
    };
  }
  
  public render(): React.ReactElement<ISpfxGraphSharepointProps> {
    var resultEl = [];
    this.state.results.forEach((item) => {
      resultEl.push(<li>{item.displayName} ({ item.webUrl })</li>);
    });

    var cTsEl = [];
    this.state.cTs.forEach((item) => {
      cTsEl.push(<li>{item.name} ({ item.group })</li>);
    });

    return (
      <div className={ styles.spfxGraphSharepoint }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            <span className={ styles.title }>Welcome to MS Graph!</span>
              <div>
                <p className={ styles.description }>Search for SP Site: <input type="text" onChange={ this.searchForChanged.bind(this) }></input> <button onClick={ this.SearchClick.bind(this) }>Search</button></p>
                <ul>{ resultEl }</ul>
              </div>
              <div>
                <p className={ styles.description }>Get current Site Content Types: <button onClick={ this.GetCTsClick.bind(this) }>Search</button></p>
                <ul>{ cTsEl }</ul>
              </div>
              <div>
                <p className={ styles.description }>Upload File to Documents Library:</p>
                <div>
                  <Dropzone onDrop={ (files) => { this.Dropzone_OnDrop(files); } }>
                    <div className={ styles.dropfileContainer }>
                      <div className={ styles.dropfileText }>Drop image here or click to select image to upload.</div>
                    </div>
                  </Dropzone>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
  
  @autobind
  private Dropzone_OnDrop(files) {
    if (files.length > 0) {
      var file = files[0];

      var reader = new FileReader();
      reader.onload = function () {
        var readerResult = reader.result;

        this.props.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
          client.api('/sites/rr87.sharepoint.com:/' + escape(this.props.currWebUrl)).get().then((response) => {
            var siteId = response.id;
            var fileContentApiUrl = '/sites/' + siteId + "/drive/root:/" + file.name + ":/content";
            client.api(fileContentApiUrl).put({
              readerResult
            }).then(() => {
              alert("File successfully uploaded.");
            }).catch((error: any) => {
              console.log(error);
            });
          }).catch((error: any) => {
            console.log(error);
          });
        }).catch((error :any) => {
          console.log(error);
        });
      }.bind(this);

      reader.readAsArrayBuffer(file);
    }
  }

  public searchForChanged(element) {
    this.setState({
      searchFor: element.target.value 
    });
  }

  public SearchClick() {
    this.props.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      client.api('/sites?search=' + this.state.searchFor).get().then((response) => {
        this.setState({
          results: response.value
        });
      }).catch((error: any) => {
        console.log(error);
      });
    }).catch((error :any) => {
      console.log(error);
    });
  }

  public GetCTsClick() {
   this.props.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
     client.api('/sites/rr87.sharepoint.com:/' + escape(this.props.currWebUrl)).get().then((response) => {
        var siteId = response.id;
        client.api('/sites/' + siteId + "/contentTypes").get().then((response2) => {
          this.setState({
            cTs: response2.value
          });
        }).catch((error: any) => {
          console.log(error);
        });
      }).catch((error: any) => {
        console.log(error);
      });
    }).catch((error :any) => {
      console.log(error);
    });
  }
}
