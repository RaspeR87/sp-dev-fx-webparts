import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface ISpfxGraphPlannerProps {
  msGraphClientFactory: any;
  
  myPlans: MicrosoftGraph.PlannerPlan[];
}

export interface ISpfxGraphPlannerState {
  myPlansDetails: any;
}