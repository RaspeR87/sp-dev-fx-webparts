import { DisplayMode } from '@microsoft/sp-core-library';

export interface IHelloWorldProps {
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
