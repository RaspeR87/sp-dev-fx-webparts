export interface ISpfxGraphSharepointProps {
  msGraphClientFactory: any;
  currWebUrl: string;
}

export interface ISpfxGraphSharepointState {
  searchFor: string;
  results: any[];
  cTs: any[];
}

