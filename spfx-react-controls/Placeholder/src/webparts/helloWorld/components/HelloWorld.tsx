import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export default class HelloWorld extends React.Component<IHelloWorldProps, {}> {
  constructor(props: IHelloWorldProps) {
    super(props);

    this._onConfigure = this._onConfigure.bind(this);
  }
  
  public render(): React.ReactElement<IHelloWorldProps> {
    if (this.props.configured) {
      return (
        <div>Configured!</div>
      );
    }
    else {
      return (
        <Placeholder
          iconName='Edit'
          iconText='Configure your web part'
          description='Please configure the web part.'
          buttonLabel='Configure'
          onConfigure={this._onConfigure} />
      );
    }
  }

  private _onConfigure() {
    this.props.context.propertyPane.open();
  }
}
