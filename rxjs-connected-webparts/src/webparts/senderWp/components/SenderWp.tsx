import * as React from 'react';
import styles from './SenderWp.module.scss';
import { ISenderWpProps } from './ISenderWpProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Button } from 'office-ui-fabric-react/lib/Button';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { ISenderWpState } from './ISenderWpState';

import { RxJsEventEmitter } from '../../../libraries/rxJsEventEmitter/RxJsEventEmitter';
import { EventData } from '../../../libraries/rxJsEventEmitter/EventData';

export default class SenderWp extends React.Component<ISenderWpProps, ISenderWpState> {

  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();

  constructor(props: ISenderWpProps) {
    super(props);

    this.state = { text: "Custom text" };
  }

  public render(): React.ReactElement<ISenderWpProps> {
    return (
      <div className={styles.senderWp}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <h2>Sender Web Part</h2>
              <TextField label="Message:" className={styles.textfield} value={this.state.text} onChanged={(e) => this.state.text = e} />
              <Button onClick={this.senderData.bind(this)} id="btnSend">
                Send data
              </Button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  /**
   * Data to all receivers.
   */
  protected senderData(): void {
        this._eventEmitter.emit("myCustomEvent:start", { text: this.state.text } as EventData);
      }
}
