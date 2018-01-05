import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IHelloWorldProps {
  context: WebPartContext;
}

export interface IHelloWorldState {
  items: any[];
}