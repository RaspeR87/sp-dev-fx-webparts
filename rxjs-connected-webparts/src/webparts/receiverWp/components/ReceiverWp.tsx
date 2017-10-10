import * as React from 'react';
import styles from './ReceiverWp.module.scss';
import { IReceiverWpProps } from './IReceiverWpProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { IReceiverWpState } from './IReceiverWpState';

import { RxJsEventEmitter } from "../../../libraries/rxJsEventEmitter/RxJsEventEmitter";
import { EventData } from "../../../libraries/rxJsEventEmitter/EventData";

export default class ReceiverWp extends React.Component<IReceiverWpProps, IReceiverWpState> {

  private readonly _eventEmitter: RxJsEventEmitter = RxJsEventEmitter.getInstance();

  constructor(props: IReceiverWpProps) {
    super(props);

    this.state = { eventsList: [] };

    // subscribe for event by event name.
    this._eventEmitter.on("myCustomEvent:start", this.receivedEvent.bind(this));
  }

  public render(): React.ReactElement<IReceiverWpProps> {
    return (
      <div className={styles.receiverWp}>
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <h2>Reactiver Web Part</h2>
              <h2>Received data:</h2>
              {
                this.state.eventsList.map((item: { index: number, data: string }) => {
                  return <div key={item.index}>{item.data}</div>;
                })
              }
            </div>
          </div>
        </div>
      </div>
    );
  }

  protected receivedEvent(data: EventData): void {
    
        // update the events list with the newly received data from the event subscriber.
        this.state.eventsList.push(
          {
            index: this.state.eventsList.length,
            data: data.text
          }
        );
    
        // set new state.
        this.setState((previousState: IReceiverWpState, props: IReceiverWpProps): IReceiverWpState => {
          previousState.eventsList = this.state.eventsList;
          return previousState;
        });
    
      }
}
