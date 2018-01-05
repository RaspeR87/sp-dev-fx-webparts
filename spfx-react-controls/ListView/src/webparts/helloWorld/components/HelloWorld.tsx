import * as React from 'react';
import styles from './HelloWorld.module.scss';
import { IHelloWorldProps, IHelloWorldState } from './IHelloWorldProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { SPHttpClient } from '@microsoft/sp-http';

export default class HelloWorld extends React.Component<IHelloWorldProps, IHelloWorldState> {
  constructor(props: IHelloWorldProps) {
    super(props);

    this.state = {
      items: []
    };
  }

  public componentDidMount() {
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Test sample list')/items`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        this.setState({
          items: items.value ? items.value : []
        });
      });
  }

  public render(): React.ReactElement<IHelloWorldProps> {
    const viewFields: IViewField[] = [
      {
        name: 'Title',
        displayName: 'Last Name',
        sorting: true,
        maxWidth: 80
      },
      {
        name: 'FirstName',
        displayName: 'First Name',
        sorting: true,
        maxWidth: 80
      },
      {
        name: 'Company',
        displayName: "Company",
        sorting: true,
        maxWidth: 80
      },
      {
        name: 'WorkPhone',
        displayName: "Business Phone",
        sorting: true,
        maxWidth: 80
      },
      {
        name: 'HomePhone',
        displayName: "Home Phone",
        sorting: true,
        maxWidth: 80
      },
      {
        name: 'Email',
        displayName: "Email Address",
        sorting: true,
        maxWidth: 100,
        render: (item: any) => {
          return <a href={"mailto:" + item['Email']}>{item['Email']}</a>;
        }
      }
    ];
    
    const groupByFields: IGrouping[] = [
      {
        name: "Company", 
        order: GroupOrder.ascending 
      }
    ];

    return (
      <ListView
      items={this.state.items}
      viewFields={viewFields}
      iconFieldName="ServerRelativeUrl"
      compact={true}
      selectionMode={SelectionMode.multiple}
      selection={this._getSelection}
      groupByFields={groupByFields} />
    );
  }

  private _getSelection(items: any[]) {
    console.log('Selected items:', items);
  }
}
