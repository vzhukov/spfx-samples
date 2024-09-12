import IWhoAmI from "../model/IWhoAmI";

export interface IDataverseComponentState {
  whoAmI?: IWhoAmI;
  entities: any[],
  loading: boolean;
  error?: any;
  instanceUrl?: string;
}
