import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  public render(): React.ReactElement<IHelloWorldProps> {
    return (
      <div>
        <WebPartTitle displayMode={this.props.displayMode}
                title={this.props.title}
                updateProperty={this.props.updateProperty} />
        <div className={ styles.helloWorld }>
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                <span className={ styles.title }>Welcome to SharePoint!</span>
                <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
                <a href="https://aka.ms/spfx" className={ styles.button }>
                  <span className={ styles.label }>Learn more</span>
                </a>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
