import * as React from 'react';
import styles from './SpfxGraphSendmail.module.scss';
import { ISpfxGraphSendmailProps, ISpfxGraphSendmailState } from './ISpfxGraphSendmailProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';

export default class SpfxGraphSendmail extends React.Component<ISpfxGraphSendmailProps, ISpfxGraphSendmailState> {

  constructor(props: ISpfxGraphSendmailProps) {
    super(props);

    this.state = {
      emailAddress: ""
    };
  }

  public render(): React.ReactElement<ISpfxGraphSendmailProps> {
    return (
      <div className={ styles.spfxGraphSendmail }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to MS Graph!</span>
              <p className={ styles.description }>Send test email to: <input type="text" onChange={ this.emailAddressChanged.bind(this) }></input> <button onClick={ this.SendEmailClick.bind(this) }>Send Email</button></p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  public emailAddressChanged(element) {
    this.setState({
      emailAddress: element.target.value 
    });
  }

  public async SendEmailClick() {
    await this.props.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
      await client.api('/me/sendMail').post({
        "message": {
          "subject": "Test email",
          "body": {
            "contentType": "Text",
            "content": "Test content"
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": this.state.emailAddress
              }
            }
          ]
        }
      }).then(() => {
        alert("The mail has been sent successfully.");
      }).catch((error: any) => {
        console.log(error);
      });
    }).catch((error :any) => {
      console.log(error);
    });
  }
}
