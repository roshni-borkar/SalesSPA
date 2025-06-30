/* eslint-disable @typescript-eslint/no-explicit-any */
export interface ISettingsState {
  selectedKey: string;

  hideO365BrandNavbar: boolean;
  hideCommentsWrapper: boolean;
  hideSiteHeader: boolean;
  hideCommandBar: boolean;
  hideSideAppBar: boolean;
  hidePageTitle: boolean;
  hideSharepointHubNavbar: boolean;
isPOPrefixPanelOpen?: boolean;
poPrefix?: string;
  qtnPrefix?: string;
  isQTNPrefixPanelOpen?: boolean;
oppPrefix?: string;
isOPPPrefixPanelOpen?: boolean;
dateFormat?: string;
isDateFormatDialogOpen?: boolean;
  [key: string]: any;
  currencySeparator?: "International" | "India" | "None";
isCurrencyDialogOpen?: boolean;

}

export const pageSettings = [
  {
    label: "Top Command Bar",
    stateVariable: "hideCommandBar",
    sharepointElement: "#spCommandBar",
    tooltip: "Hides the Command Bar (containing New, Share, Edit, etc...)"
  },
  {
    label: "Side App Bar",
    stateVariable: "hideSideAppBar",
    sharepointElement: "#sp-appBar",
    tooltip: "Hides the SharePoint Side Navigation Bar"
  },
  {
    label: "Page Title",
    stateVariable: "hidePageTitle",
    sharepointElement: "[id*='PageTitle']",
    tooltip: "Hides the Page Title"
  },
  {
    label: "Site Navigation Bar",
    stateVariable: "hideSiteHeader",
    sharepointElement: "#spSiteHeader",
    tooltip: "Hides the SharePoint Site Navigation Bar"
  },
  {
    label: "Comments Section",
    stateVariable: "hideCommentsWrapper",
    sharepointElement: "#CommentsWrapper",
    tooltip: "Hides the Like/Comment Section"
  },
  {
    label: "O365 Brand Navigation Bar",
    stateVariable: "hideO365BrandNavbar",
    sharepointElement: "#SuiteNavWrapper",
    tooltip: "Hides the O365 Navigation Bar"
  },
  {
    label: "SharePoint Hub Navigation Bar",
    stateVariable: "hideSharepointHubNavbar",
    sharepointElement: ".ms-HubNav",
    tooltip: "Hides the SharePoint Hub Navigation"
  },
];


