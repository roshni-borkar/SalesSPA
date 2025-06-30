 import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { ISettings } from "./Sales";
export interface ISalesProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  sp: SPFI;
  context: WebPartContext;
  View:string;
  //isDeployedToMainSite: boolean;
  //ProjectsSite: string;
  //ProjectName: string;
  //ComponentDropdown: string;
 // TimeSite: string;
  //ComponentHeight: string;
  //Planner: string;
 // ViewAs: any[];
 // siteOption: "mainSite" | "teamSite";
 // renderingMode: TRenderingMode;

  //groupId: string;
  //groupName: string;
  //groupPlans: IGroupPlan[];
 // groupOwners: any[];

  settings: ISettings;
  onConfigChange(config: ISettings): void;
}
