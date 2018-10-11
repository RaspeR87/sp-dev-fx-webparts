import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IBucketProps {
    msGraphClientFactory: any;

    bucket: MicrosoftGraph.PlannerBucket;
}

export interface IBucketState {
    tasks: MicrosoftGraph.PlannerTask[];
}