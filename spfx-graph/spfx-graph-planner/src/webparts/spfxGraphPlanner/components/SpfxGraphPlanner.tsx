import * as React from 'react';
import styles from './SpfxGraphPlanner.module.scss';
import { ISpfxGraphPlannerProps } from './ISpfxGraphPlannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

import TasksPlaceHolder from './TasksPlaceHolder/TasksPlaceHolder';

import * as moment from 'moment';

export default class SpfxGraphPlanner extends React.Component<ISpfxGraphPlannerProps, {}> {

  public render(): React.ReactElement<ISpfxGraphPlannerProps> {
    return (
      <div className={ styles.spfxGraphPlanner }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to MS Graph!</span>
              <p className={ styles.subTitle }>:: My Plans from Planner ::</p>
              {
                this.props.myPlans.map(plan => {
                  return (<div>
                    <p className={ styles.description }>Title: <b>{plan.title}</b></p>
                    <p className={ styles.description }>Created: <b>{moment(plan.createdDateTime).format("LLL")}</b></p>
                    <TasksPlaceHolder msGraphClientFactory={ this.props.msGraphClientFactory } planId={ plan.id }></TasksPlaceHolder>
                  </div>);
                })
              }
            </div>
          </div>
        </div>
      </div>
    );
  }
}
