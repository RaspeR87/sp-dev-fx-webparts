import * as React from 'react';
import styles from './SpFxGraphPlanner.module.scss';
import { ISpFxGraphPlannerProps } from './ISpFxGraphPlannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class SpFxGraphPlanner extends React.Component<ISpFxGraphPlannerProps, {}> {
  public render(): React.ReactElement<ISpFxGraphPlannerProps> {
    return (
      <div className={ styles.spFxGraphPlanner }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to MS Graph!</span>
              <p className={ styles.subTitle }>:: My informations ::</p>
              <p className={ styles.description }>Display Name: {this.props.myInformations.displayName}</p>
              <p className={ styles.description }>Given Name: {this.props.myInformations.givenName}</p>
              <p className={ styles.description }>Surname: {this.props.myInformations.surname}</p>
              <p className={ styles.description }>Job Title: {this.props.myInformations.jobTitle}</p>
              <p className={ styles.description }>Mail: {this.props.myInformations.mail}</p>
              <p className={ styles.description }>Mobile Phone: {this.props.myInformations.mobilePhone}</p>
              <p className={ styles.description }>Office Location: {this.props.myInformations.officeLocation}</p>
              <p className={ styles.description }>Preferred Language: {this.props.myInformations.preferredLanguage}</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
