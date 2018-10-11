import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface ITasksPlaceHolderProps {
    msGraphClientFactory: any;

    planId: any;
}
  
export interface ITasksPlaceHolderState {
    showTasks: boolean;

    buckets: MicrosoftGraph.PlannerBucket[];
}