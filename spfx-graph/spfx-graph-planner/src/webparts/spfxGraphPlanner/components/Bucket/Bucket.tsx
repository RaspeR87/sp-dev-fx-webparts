import * as React from 'react';

import styles from '../SpfxGraphPlanner.module.scss';
import { IBucketProps, IBucketState } from './IBucketProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

import * as moment from 'moment';

export default class Bucket extends React.Component<IBucketProps, IBucketState> {

    constructor(props: IBucketProps) {
        super(props);

        this.state = {
            tasks: []
        };
      }

      async componentDidMount() {
        await this.props.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
            await client.api('planner/buckets/' + this.props.bucket.id + '/tasks').get().then((data:any) => {
                this.setState({
                    tasks: data.value
                });
            }).catch((error: any) => {
                console.log(error);
            });
            }).catch((error :any) => {
            console.log(error);
        });
      }

    public render(): React.ReactElement<IBucketProps> {
        var tasksEl = [];
        this.state.tasks.forEach((item) => {
            tasksEl.push(<p>
                <div>Task: <b>{ item.title }</b></div>
                <div>Percentage Compete: <b>{ item.percentComplete }</b></div>
                <div>Start Date: <b>{ item.startDateTime ? moment(item.startDateTime).format("LLL") : "" }</b></div>
                <div>Created: <b>{ item.createdDateTime ? moment(item.createdDateTime).format("LLL") : "" }</b></div>
                <div>Due Date: <b>{ item.dueDateTime ? moment(item.dueDateTime).format("LLL") : "" }</b></div>
            </p>);
        });

        return (
            <div>
                <div>Bucket: <b>{ this.props.bucket.name }</b></div>
                { tasksEl }
            </div>
        );
    }

}