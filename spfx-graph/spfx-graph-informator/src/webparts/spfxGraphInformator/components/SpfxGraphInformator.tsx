import * as React from 'react';
import styles from './SpfxGraphInformator.module.scss';
import { ISpfxGraphInformatorProps } from './ISpfxGraphInformatorProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpfxGraphInformator extends React.Component<ISpfxGraphInformatorProps, {}> {
  public render(): React.ReactElement<ISpfxGraphInformatorProps> {
    return (
      <div className={ styles.spfxGraphInformator }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            <span className={ styles.title }>Welcome to MS Graph!</span>
              <p className={ styles.subTitle }>:: My informations ::</p>
              <p className={ styles.description }>Number of New Emails: {this.props.nrNewMails} / {this.props.nrTotalMails}</p>
              <p className={ styles.description }>Number of Upcoming Events: {this.props.nrUpcomingEvents}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
