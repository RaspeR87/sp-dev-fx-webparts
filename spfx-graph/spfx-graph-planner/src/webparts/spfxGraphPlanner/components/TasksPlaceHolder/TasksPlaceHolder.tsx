import * as React from 'react';
import styles from '../SpfxGraphPlanner.module.scss';
import { ITasksPlaceHolderProps, ITasksPlaceHolderState } from './ITasksPlaceHolderProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';

import Bucket from '../Bucket/Bucket';

export default class TasksPlaceHolder extends React.Component<ITasksPlaceHolderProps, ITasksPlaceHolderState> {

    constructor(props: ITasksPlaceHolderProps) {
        super(props);

        this.state = {
            buckets: [],
            showTasks: false
        };
      }

    public ShowTasksClick() {
        this.props.msGraphClientFactory.getClient().then(async (client: MSGraphClient): Promise<void> => {
          client.api('planner/plans/' + this.props.planId + '/buckets').get().then((data:any) => {
              this.setState({
                buckets: data.value,
                showTasks: true
              })
          }).catch((error: any) => {
            console.log(error);
          });
        }).catch((error :any) => {
          console.log(error);
        });
      }

    public render(): React.ReactElement<ITasksPlaceHolderProps> {
        var bucketsEl = [];
        this.state.buckets.forEach((item) => {
            bucketsEl.push(<Bucket msGraphClientFactory={ this.props.msGraphClientFactory } bucket={ item }></Bucket>);
        });

        return (
            <div>
                <p><button onClick={ this.ShowTasksClick.bind(this) }>Show Tasks</button></p>
                {
                    (this.state.showTasks ? 
                        <div>
                            { bucketsEl }
                        </div> 
                        : <span></span>)
                }
            </div>
        );
    }

}