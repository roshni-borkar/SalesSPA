/* eslint-disable prefer-const */
/* eslint-disable max-lines */
/* eslint-disable no-constant-condition */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable no-unused-expressions */
/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
// import styles from './Sales.module.scss';
import { ISalesProps } from "./ISalesProps";
import {
  // DatePicker,
  // Dropdown,
  // IButtonStyles,
  IDropdownOption,
  INavLink,
  INavLinkGroup,
  // IconButton,
  // Label,
  Nav,
  // PrimaryButton,
  Stack,
  // TextField,
  // getTheme,
  Persona,
  IconButton,
} from "@fluentui/react";
import { LivePersona } from "@pnp/spfx-controls-react"; // Add this import
import Graph from "./Graph/graph";
// import { Text as FluentText } from '@fluentui/react';
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import OpportunityViewer from "./ViewOpportunities/ViewOpportunies";
import QuotationViewer from "./ViewQuotation/ViewQuotation";
import POViewer from "./VIewPO/ViewPO";
import Settings from "./Settings/Settings";
import Logo from "./Logo";
import { pageSettings } from "./Settings/ISettingsState";
import SalesDashboardComponent from "./SalesDashboard/SalesDashboard";

// import Settings from './Settings/Settings';

// import { pageSettings } from './Settings/ISettingsState';
// import OpportunityViewer from './ViewOpportunities/ViewOpportunies';
// const ganttViews: INavLink = {
//   key: "GanttView",
//   name: "Gantt View",
//   isLabel: true,
//   url: "",
//   isExpanded: true,
//   links: []
// };
// const planners: INavLink = {
//   key: "Planner",
//   name: "Planner",
//   isLabel: true,
//   url: "",
//   isExpanded: true,
//   links: []
// };

// const AddNewOpportunities = {
//   key: "Opportunity",
//   title: "Add New Opportunities",
//   name: "Add New Opportunities",
//   url: "",
//   // iconProps: {
//   //   iconName: "AnalyticsReport"
//   // }
// };

// const AddNewQuotation = {
//   key: "Quotation",
//   title: "AddNewQuotation",
//   name: "Add New Quotation",
//   url: "",
//   // iconProps: {
//   //   iconName: "PeopleAdd"
//   // }
// };

// const AddNewPO = {
//   key: "Purchase Order",
//   title: "AddNewPO",
//   name: "Add New PO",
//   url: "",
//   // iconProps: {
//   //   iconName: "Settings"
//   // }
// };
const context: any = {};
export interface ISettings {
  [key: string]: any;
}
export type TContextObject = {
  //
  site: {
    //
    webAbsoluteProject: string;
    webAbsoluteTime: string;
    projectsSiteName: string;
    timeSiteName: string;
    ListTAT: string;
  };
  //viewAs: IPropertyFieldGroupOrPerson;
  config: {
    memberRemovalDayLimit: number;
    totalProjectHours: number;
    installedCountry: string;
    projectStatusOptions: string;
    adminCreateUnAssigned: string;
    teamsForVendorHours: string;
    isAllocationEnabled: boolean;
    isHolidaysInDayPilotEnabled: boolean;
    isTaskAssignmentModuleEnabled: boolean;
    isPOFunctionalityEnabled: boolean;
    isPOManuallyEntered: boolean;

    daypilotAssignmentColor: string;
    daypilotTeamColor: string;
    daypilotProjectColor: string;
    daypilotLeaveColor: string;
    daypilotOptionalHolidayColor: string;
    daypilotGeneralHolidayColor: string;
    daypilotTotalColor: string;
    daypilotTeamUnAssignedProjectColor: string;
    showProjectColorInCount: boolean;

    isOrgansationViewEnabledForDH: boolean;
    projectTypeOptions: string;
    daypilotTotalMembersLabel: string;
    unassignedPMMail: string;
    isProjectCodeEditable: boolean;
    showDropdownForAssignmentPercent: boolean;
    isDHFieldHiddenInProjectCreation: boolean;
  };
};
export const TaskContext: any = React.createContext(context);
const ViewOpportunities = {
  key: "View Opportunities",
  title: "ViewOpportunities",
  name: "Opportunities",
  url: "",
  iconProps: {
    iconName: "OpenInNewWindow",
  },
};
const ViewQuotation = {
  key: "View Quotation",
  title: "ViewQuotation",
  name: "Quotations",
  url: "",
  iconProps: {
    iconName: "Document",
  },
};
const ViewPurchaseOrder = {
  key: "View PO",
  title: "ViewPurchaseOrder",
  name: "Purchase Orders",
  url: "",
  iconProps: {
    iconName: "ReceiptCheck",
  },
};
const SalesDashboard = {
  key: "ViewSalesDashboard",
  title: "ViewSalesDashboard",
  name: "ViewSalesDashboard",
  url: "",
  iconProps: {
    iconName: "BIDashboard",
  },
};
const Setting = {
  key: "Setting",
  title: "Setting",
  name: "Settings",
  url: "",
  iconProps: {
    iconName: "Settings",
  },
};
const navLinks: INavLinkGroup[] = [
  {
    links: [
      // AddNewOpportunities,
      // AddNewQuotation,
      // AddNewPO,
      ViewOpportunities,
      ViewQuotation,
      ViewPurchaseOrder,
      SalesDashboard,
      Setting,
    ],
  },
];
// const currencies: IDropdownOption[] = [
//   { key: "EUR", text: "EUR" },
//   { key: "USD", text: "USD" },
//   { key: "GBP", text: "GBP" },
// ];
// const sectionHeaderStyle = {
//   backgroundColor: "#176D7E",
//   color: "#fff",
//   fontWeight: 600,
//   fontSize: 20,
//   padding: "8px 16px",
//   borderRadius: "4px",
//   marginTop: 24,
// };

// const lineItemStyle = {
//   display: "grid",
//   gridTemplateColumns: "1fr 1fr 1fr 50px",
//   gap: 12,
//   alignItems: "center",
//   marginTop: 12,
// };
// const statusOptions: IDropdownOption[] = [
//   { key: "Open", text: "Open" },
//   { key: "Won", text: "Won" },
//   { key: "Lost", text: "Lost" },
//   { key: "On Hold", text: "On Hold" },
// ];

// const riskLevels: IDropdownOption[] = [
//   { key: "Low", text: "Low" },
//   { key: "Medium", text: "Medium" },
//   { key: "High", text: "High" },
// ];

// const strategicOptions: IDropdownOption[] = [
//   { key: "Yes", text: "Yes" },
//   { key: "No", text: "No" },
// ];

// const businessSizes: IDropdownOption[] = [
//   { key: "Small", text: "Small" },
//   { key: "Medium", text: "Medium" },
//   { key: "Large", text: "Large" },
// ];

// const dropdownOptions: IDropdownOption[] = [{ key: 'find', text: 'Find items' }];
// const currencyOptions: IDropdownOption[] = [{ key: 'eur', text: 'EUR' }];

// const columnStyles = {
//   root: {
//     display: 'grid',
//     gridTemplateColumns: 'repeat(auto-fit, minmax(250px, 1fr))',
//     gap: '16px',
//   },
// };
// const columnGridStyle = {
//   root: {
//     display: "grid",
//     gridTemplateColumns: "repeat(3, 1fr)",
//     gap: "16px",
//   },
// };
// const gridStyle = {
//   display: "grid",
//   gridTemplateColumns: "repeat(3, 1fr)",
//   gap: "16px",
// };
// const theme = getTheme();

// const primaryButtonStyles: IButtonStyles = {
//   root: {
//     backgroundColor: theme.palette.themePrimary,
//     borderColor: theme.palette.themePrimary,
//   },
//   rootHovered: {
//     backgroundColor: theme.palette.themeDark,
//     borderColor: theme.palette.themeDark,
//   },
//   rootPressed: {
//     backgroundColor: theme.palette.themeDarker,
//     borderColor: theme.palette.themeDarker,
//   },
// };
interface ISalesState {
  tab: string;
  opportunityForm: any;
  quotationForm: any;
  purchaseOrderForm: any;
  opportunityOptions: IDropdownOption[];
  quoteOptions: IDropdownOption[];
  purchaseOrders: any[];
  quotationFile: any;
  lineItems: { Title: string; Comments: string; Value: number }[];
  isNavOpen: boolean;
}

export default class Sales extends React.Component<ISalesProps, ISalesState> {
  removeLineItem = (index: number) => {
    const newItems = [...this.state.lineItems];
    newItems.splice(index, 1);
    this.setState({ lineItems: newItems });
  };
  private sp: SPFI;
  private graph: Graph;
  private imageUrl: string | undefined;
  private contextObject: any = {};
  navLinks: INavLinkGroup[] | null;

  //   public async componentDidMount(): Promise<void> {
  //     console.log(this.sp);
  //     void this.fetchOpportunities();
  //     this.createList();
  //       this.graph = new Graph(this.props.context.msGraphClientFactory);
  //  this.loadImageUrl();
  //     this.getConfiguration();
  //     this.setPageSettings();
  //     //this.sp = spfi().using(SPFx(this.props.context));
  //     // SPFI instance is already configured with SPFx context during initialization
  //     console.log("SPFI instance initialized with SPFx context");
  //   }
  public async componentDidMount(): Promise<void> {
    try {
      await Promise.all([
        // this.fetchOpportunities(),
        this.createList(),
        this.OEMcreateList(),
      
      this.CustomercreateList(),
      this.KeyContactcreateList(),
    
this.ConfigcreateList(),
this.AuditLogcreateList(),
        this.loadImageUrl(),
        this.getConfiguration(),
      ]);
      this.setPageSettings();
    } catch (error) {
      console.error("Initialization failed", error);
    }
  }

  constructor(props: ISalesProps) {
    super(props);
    this.state = {
      tab: "View Opportunities",
      opportunityForm: {},
      quotationForm: {},
      purchaseOrderForm: {},
      opportunityOptions: [],
      quoteOptions: [],
      purchaseOrders: [],
      quotationFile: null,
      lineItems: [{ Title: "", Comments: "", Value: 0 }],
      isNavOpen: true,
    };

    this.contextObject = {
      context: this.props.context,
      sp:spfi().using(SPFx(this.props.context)),

      site: {
      
        // ComponentHeight: this.props.ComponentHeight,
      },
      group: {
        //  id: this.props.groupId,
        //  name: this.props.groupName,
      
        members: [],
      
      },
      user: {
        name: this.props.context.pageContext.user.displayName,
        email: this.props.context.pageContext.user.email.toLowerCase(),
        loginName: this.props.context.pageContext.user.loginName.toLowerCase(),
        role: null,
      },
      //viewAs: this.props.ViewAs ? this.props.ViewAs.length ? this.props.ViewAs[0] : null : null,
      //siteOption: this.props.siteOption,
      config: {
        item: context.config?.item || null,
        settings: this.props.settings,
        onChange: this.onSettingsChange,
      },
     
    };
     this.onSettingsChange = this.onSettingsChange.bind(this);
    this.onLinkCLick = this.onLinkCLick.bind(this);
    console.log(this.props.context);
    console.log(this.props.sp);
    this.sp = spfi().using(SPFx(this.props.context));

    
    console.log(this.sp);
  }

  // private setTab = (t: string): void => {
  //   this.setState({ tab: t });
  // };
    private onSettingsChange(settings: ISettings) {
    this.contextObject.config.settings = settings;
    this.props.onConfigChange(settings);
    // this.getNavLinks();
  }
  private async loadImageUrl() {
   try{this.imageUrl = (await this.graph.getUserPhoto(
      this.props.context.pageContext.user.email
    )) as string;
    console.log("Image URL:", this.imageUrl);}
    catch (error) {
      console.error("❌ Failed to load user photo:", error);
      this.imageUrl = ""; // Set to undefined if loading fails
    }
    
  }
private async getConfiguration() {
  try {
    const items = await this.sp.web.lists.getByTitle("CWSalesConfiguration").items();

    const config = items.find((item) => item.Title === "Settings");

    if (!config) {
      console.warn("⚠️ No 'Settings' item found in CWSalesConfiguration list.");
      return;
    }

    const settings: ISettings = config.MultiValue
      ? JSON.parse(config.MultiValue)
      : {};

    console.log("✅ Settings loaded from config item:", settings);

    // Ensure contextObject and config are defined
    if (!this.contextObject) {
      this.contextObject = {};
    }
    if (!this.contextObject.config) {
      this.contextObject.config = {};
    }

    this.contextObject.config.item = config;
    this.contextObject.config.settings = settings;
    this.contextObject.config.onChange = this.onSettingsChange;

    // Optional: call props callback if needed
    if (typeof this.props.onConfigChange === "function") {
      this.props.onConfigChange(settings);
    }

    this.setPageSettings();
  } catch (error) {
    console.error("❌ Failed to load configuration:", error);
  }
}

  // private fetchOpportunities = async () => {
  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.select("Id", "Title", "OpportunityID")();

  //     const options = items.map((item) => ({
  //       key: item.OpportunityID || item.Id,
  //       text: item.OpportunityID || `Opportunity ${item.Id}`,
  //     }));

  //     this.setState({ opportunityOptions: options });
  //   } catch (err) {
  //     console.error("Failed to fetch opportunities", err);
  //   }
  // };
  // private fetchQuotesForOpportunity = async (
  //   opportunityID: string
  // ): Promise<void> => {
  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.filter(`OpportunityID eq '${opportunityID}'`)
  //       .top(1)();

  //     if (items.length === 0) {
  //       this.setState({ quoteOptions: [] });
  //       return;
  //     }

  //     const item = items[0];
  //     const quotes: IDropdownOption[] = [];

  //     for (let i = 1; i <= 5; i++) {
  //       const suffix = i === 1 ? "" : i.toString();
  //       const quoteId = item[`QuoteID${suffix}`];
  //       if (quoteId) {
  //         quotes.push({ key: quoteId, text: quoteId });
  //       }
  //     }

  //     this.setState({ quoteOptions: quotes });
  //   } catch (err) {
  //     console.error("Failed to fetch quotes for opportunity", err);
  //     this.setState({ quoteOptions: [] });
  //   }
  // };
  // private fetchPurchaseOrders = async (): Promise<void> => {
  //   const opportunityID = this.state.purchaseOrderForm.OpportunityID;

  //   if (!opportunityID) {
  //     alert("Select an Opportunity ID first.");
  //     return;
  //   }

  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.filter(`OpportunityID eq '${opportunityID}'`)
  //       .top(1)();

  //     if (items.length === 0) {
  //       alert("No Opportunity found.");
  //       return;
  //     }

  //     const item = items[0];
  //     const purchaseOrders: any[] = [];

  //     for (let i = 1; i <= 5; i++) {
  //       const suffix = i === 1 ? "" : i.toString();
  //       const poId = item[`POID${suffix}`];
  //       if (poId) {
  //         purchaseOrders.push({
  //           POID: poId,
  //           POReceivedDate: item[`POReceivedDate${suffix}`],
  //           CustomerPONumber: item[`CustomerPONumber${suffix}`],
  //           POStatus: item[`POStatus${suffix}`],
  //           QuoteID: item[`QuoteID${suffix}`],
  //           AmountEUR: item[`AmountEUR${suffix}`],
  //           Currency: item[`Currency${suffix}`],
  //           LineItemsJSON: item[`LineItemsJSON${suffix}`],
  //         });
  //       }
  //     }

  //     this.setState({ purchaseOrders });
  //   } catch (err) {
  //     console.error("Error fetching purchase orders", err);
  //     alert("Failed to fetch purchase orders.");
  //   }
  // };
  // private renderPurchaseOrderList = () => {
  //   const { purchaseOrders } = this.state;
  //   if (!purchaseOrders.length) return null;

  //   return (
  //     <div style={{ marginTop: "1rem" }}>
  //       <Label styles={{ root: { fontSize: "20px", fontWeight: "bold" } }}>
  //         Existing Purchase Orders
  //       </Label>

  //       <Stack tokens={{ childrenGap: 12 }}>
  //         {purchaseOrders.map((po, index) => (
  //           <Stack
  //             key={index}
  //             tokens={{ childrenGap: 6 }}
  //             styles={{
  //               root: {
  //                 border: "1px solid #E1DFDD",
  //                 borderRadius: 8,
  //                 padding: 12,
  //                 backgroundColor: "#ffffff",
  //                 boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
  //               },
  //             }}
  //           >
  //             <span>
  //               <strong>PO ID:</strong> {po.POID}
  //             </span>
  //             <span>
  //               <strong>Quote ID:</strong> {po.QuoteID}
  //             </span>
  //             <span>
  //               <strong>Status:</strong> {po.POStatus}
  //             </span>
  //             <span>
  //               <strong>Customer PO Number:</strong> {po.CustomerPONumber}
  //             </span>
  //             <span>
  //               <strong>Amount (EUR):</strong> {po.AmountEUR}
  //             </span>
  //             <span>
  //               <strong>Currency:</strong> {po.Currency}
  //             </span>
  //             <span>
  //               <strong>Received Date:</strong> {po.POReceivedDate}
  //             </span>
  //           </Stack>
  //         ))}
  //       </Stack>
  //     </div>
  //   );
  // };
  private onLinkCLick(_: any, item: INavLink) {
    this.setState({ tab: item.key || "View Opportunities" });
    this.renderComponent();
  }
  private createList = async (): Promise<void> => {
    const listTitle = "CWSalesRecords";

    // Try to get the list
    let listExists = true;
    try {
      await this.sp.web.lists.getByTitle(listTitle)();
      // const viewTitle = "CustomView";
      // const itemLimit = 10;
      // const list = this.sp.web.lists.getByTitle("SalesRecords");
      // await list.views.add(viewTitle, false, {
      //    RowLimit: itemLimit,
      //    PersonalView: false,
      //    SetAsDefaultView: true,
      //    ViewQuery: "", // optional CAML query if you want filters
      //    ViewFields: ["Title"], // choose your fields
      //  });
      await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .views.add("Sales_View5", false, {
          RowLimit: 0,
          ViewQuery: `<Query>
       <Where>
         <Or>
           <Eq>
             <FieldRef Name="None" />
             <Value Type="Text"></Value>
           </Eq>
           <Eq>
             <FieldRef Name="None" />
             <Value Type="Text"></Value>
           </Eq>
         </Or>
       </Where>
     </Query>`,
        });
      await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .views.getByTitle("Sales_View5")
        .renderAsHtml();

      // RowLimit: 0,
      //   SetAsDefaultView: true,
      //   ViewFields: ["Title"], // make sure these internal fields exist
      //   ViewQuery: "", // Optional CAML query
      // });
    } catch {
      listExists = false;
    }

    // Create list if not exists
    if (!listExists) {
      await this.sp.web.lists.add(listTitle, "Unified Sales Data", 100, false);

      console.log(`Created list: ${listTitle}`);
    } else {
      console.log(`List '${listTitle}' already exists`);
    }
    if (!listExists) {
      const fields = [
        // {
        //   name: "RecordType",
        //   type: "Choice",
        //   choices: ["Opportunity", "Quotation", "PurchaseOrder"],
        // },
        { name: "OpportunityID", type: "Text" },
        { name: "QuoteID", type: "Text" },
        { name: "QuoteID2", type: "Text" },
        { name: "QuoteID3", type: "Text" },
        { name: "QuoteID4", type: "Text" },
        { name: "QuoteID5", type: "Text" },
        { name: "POID", type: "Text" },
        { name: "POID2", type: "Text" },
        { name: "POID3", type: "Text" },
        { name: "POID4", type: "Text" },
        { name: "POID5", type: "Text" },
        { name: "Business", type: "Text" },
        { name: "BusinessUnit", type: "Text" },
        { name: "OEM", type: "Text" },
        { name: "KeyContact", type: "Text" },
        { name: "DecisionMaker", type: "Text" },
        { name: "Customer", type: "Text" },
        { name: "Status", type: "Text" },
        { name: "EndCustomer", type: "Text" },
        {
          name: "QuoteBusinessSize",
          type: "Choice",
          choices: ["Small", "Medium", "Large"],
        },
        {
          name: "QuoteBusinessSize2",
          type: "Choice",
          choices: ["Small", "Medium", "Large"],
        },
        {
          name: "QuoteBusinessSize3",
          type: "Choice",
          choices: ["Small", "Medium", "Large"],
        },
        {
          name: "QuoteBusinessSize4",
          type: "Choice",
          choices: ["Small", "Medium", "Large"],
        },
        {
          name: "QuoteBusinessSize5",
          type: "Choice",
          choices: ["Small", "Medium", "Large"],
        },

        { name: "ReportDate", type: "DateTime" },
        { name: "TentativeStartDate", type: "DateTime" },
        { name: "TentativeDecisionDate", type: "DateTime" },
        { name: "QuoteTentativeDecisionDate", type: "DateTime" },
        { name: "QuoteTentativeDecisionDate2", type: "DateTime" },
        { name: "QuoteTentativeDecisionDate3", type: "DateTime" },
        { name: "QuoteTentativeDecisionDate4", type: "DateTime" },
        { name: "QuoteTentativeDecisionDate5", type: "DateTime" },

        { name: "QuoteDate", type: "DateTime" },
        { name: "QuoteDate2", type: "DateTime" },
        { name: "QuoteDate3", type: "DateTime" },
        { name: "QuoteDate4", type: "DateTime" },
        { name: "QuoteDate5", type: "DateTime" },
        { name: "QuoteRevisionNumber", type: "Number" },
        { name: "QuoteRevisionNumber2", type: "Number" },
        { name: "QuoteRevisionNumber3", type: "Number" },
        { name: "QuoteRevisionNumber4", type: "Number" },
        { name: "QuoteRevisionNumber5", type: "Number" },
        { name: "OppAmount", type: "Number" },
        { name: "OppAmount2", type: "Number" },
        { name: "OppAmount3", type: "Number" },
        { name: "OppAmount4", type: "Number" },
        { name: "OppAmount5", type: "Number" },
        { name: "QuoteAmount", type: "Number" },
        { name: "QuoteAmount2", type: "Number" },
        { name: "QuoteAmount3", type: "Number" },
        { name: "QuoteAmount4", type: "Number" },
        { name: "QuoteAmount5", type: "Number" },
        { name: "POAmount", type: "Number" },
        { name: "POAmount2", type: "Number" },
        { name: "POAmount3", type: "Number" },
        { name: "POAmount4", type: "Number" },
        { name: "POAmount5", type: "Number" },
        { name: "POValue", type: "Number" },
        { name: "POValue2", type: "Number" },
        { name: "POValue3", type: "Number" },
        { name: "POValue4", type: "Number" },
        { name: "POValue5", type: "Number" },
        { name: "POQuoteID", type: "String" },
        { name: "POQuoteID2", type: "String" },
        { name: "POQuoteID3", type: "String" },
        { name: "POQuoteID4", type: "String" },
        { name: "POQuoteID5", type: "String" },

        { name: "QuoteReceivedDate1", type: "DateTime" },
        { name: "QuoteReceivedDate", type: "DateTime" },
        { name: "QuoteReceivedDate2", type: "DateTime" },
        { name: "QuoteReceivedDate3", type: "DateTime" },
        { name: "QuoteReceivedDate4", type: "DateTime" },
        { name: "QuoteReceivedDate5", type: "DateTime" },
        { name: "POReceivedDate", type: "DateTime" },
        { name: "POReceivedDate2", type: "DateTime" },
        { name: "POReceivedDate3", type: "DateTime" },
        { name: "POReceivedDate4", type: "DateTime" },
        { name: "POReceivedDate5", type: "DateTime" },

        {
          name: "POStatus",
          type: "Choice",
          choices: ["Draft", "Issued", "Approved", "Cancelled"],
        },
        {
          name: "POStatus2",
          type: "Choice",
          choices: ["Draft", "Issued", "Approved", "Cancelled"],
        },
        {
          name: "POStatus3",
          type: "Choice",
          choices: ["Draft", "Issued", "Approved", "Cancelled"],
        },
        {
          name: "POStatus4",
          type: "Choice",
          choices: ["Draft", "Issued", "Approved", "Cancelled"],
        },
        {
          name: "POStatus5",
          type: "Choice",
          choices: ["Draft", "Issued", "Approved", "Cancelled"],
        },
        { name: "CustomerPONumber", type: "Text" },
        { name: "CustomerPONumber2", type: "Text" },
        { name: "CustomerPONumber3", type: "Text" },
        { name: "CustomerPONumber4", type: "Text" },
        { name: "CustomerPONumber5", type: "Text" },
        { name: "IsChildPO", type: "Boolean" },
        { name: "IsChildPO2", type: "Boolean" },
        { name: "IsChildPO3", type: "Boolean" },
        { name: "IsChildPO4", type: "Boolean" },
        { name: "IsChildPO5", type: "Boolean" },
        { name: "ParentPOID", type: "Text" },
         { name: "ParentPOID2", type: "Text" },
          { name: "ParentPOID3", type: "Text" },
           { name: "ParentPOID4", type: "Text" },
            { name: "ParentPOID5", type: "Text" },
        { name: "AmountEUR", type: "Currency" },
        { name: "Currency", type: "Choice", choices: ["EUR", "USD", "GBP"] },
        // { name: "POCurrency", type: "Choice", choices: ["EUR", "USD", "GBP"] },
        // { name: "POCurrency2", type: "Choice", choices: ["EUR", "USD", "GBP"] },
        // { name: "POCurrency3", type: "Choice", choices: ["EUR", "USD", "GBP"] },
        // { name: "POCurrency4", type: "Choice", choices: ["EUR", "USD", "GBP"] },
        // { name: "POCurrency5", type: "Choice", choices: ["EUR", "USD", "GBP"] },
        // {
        //   name: "QuoteCurrency",
        //   type: "Choice",
        //   choices: ["EUR", "USD", "GBP"],
        // },
        // {
        //   name: "QuoteCurrency2",
        //   type: "Choice",
        //   choices: ["EUR", "USD", "GBP"],
        // },
        // {
        //   name: "QuoteCurrency3",
        //   type: "Choice",
        //   choices: ["EUR", "USD", "GBP"],
        // },
        // {
        //   name: "QuoteCurrency4",
        //   type: "Choice",
        //   choices: ["EUR", "USD", "GBP"],
        // },
        // {
        //   name: "QuoteCurrency5",
        //   type: "Choice",
        //   choices: ["EUR", "USD", "GBP"],
        // },
        { name: "ConvertedAmount", type: "Number" },
        { name: "QuoteRevenueQuoted", type: "Number" },
        { name: "QuoteRevenueQuoted2", type: "Number" },
        { name: "QuoteRevenueQuoted3", type: "Number" },
        { name: "QuoteRevenueQuoted4", type: "Number" },
        { name: "QuoteRevenueQuoted5", type: "Number" },
        {
          name: "RiskLevel",
          type: "Choice",
          choices: ["Low", "Medium", "High"],
        },
        { name: "Strategic", type: "Choice", choices: ["Yes", "No"] },
        {
          name: "OpportunityStatus",
          type: "Choice",
          choices: ["Open", "Won", "Lost", "On Hold"],
        },
        { name: "OppComments", type: "Note" },
        { name: "OppComments2", type: "Note" },
        { name: "OppComments3", type: "Note" },
        { name: "OppComments4", type: "Note" },
        { name: "OppComments5", type: "Note" },
        { name: "QuoteComments", type: "Note" },
        { name: "QuoteComments2", type: "Note" },
        { name: "QuoteComments3", type: "Note" },
        { name: "QuoteComments4", type: "Note" },
        { name: "QuoteComments5", type: "Note" },
        { name: "POComments", type: "Note" },
        { name: "POComments2", type: "Note" },
        { name: "POComments3", type: "Note" },
        { name: "POComments4", type: "Note" },
        { name: "POComments5", type: "Note" },
        { name: "LineItemsJSON", type: "Note" },
        { name: "LineItemsJSON2", type: "Note" },
        { name: "LineItemsJSON3", type: "Note" },
        { name: "LineItemsJSON4", type: "Note" },
        { name: "LineItemsJSON5", type: "Note" },
      ];

      const list = this.sp.web.lists.getByTitle(listTitle);

      for (const field of fields) {
        try {
          if (field.type === "Choice") {
            await list.fields.addChoice(field.name, {
              Choices: field.choices!,
              Required: false,
            });
          } else if (field.type === "Boolean") {
            await list.fields.addBoolean(field.name);
          } else if (field.type === "Currency") {
            await list.fields.addCurrency(field.name);
          } else if (field.type === "Number") {
            await list.fields.addNumber(field.name);
          } else if (field.type === "DateTime") {
            await list.fields.addDateTime(field.name);
          } else if (field.type === "Note") {
            await list.fields.addMultilineText(field.name);
          } else {
            await list.fields.addText(field.name);
          }
          console.log(`✔ Field added: ${field.name}`);
        } catch (err) {
          console.warn(
            `⚠ Field '${field.name}' might already exist or failed:`,
            err
          );
        }
      }
    }
  };
    private OEMcreateList = async (): Promise<void> => {
    const listTitle = "CWSalesOEM";

    // Try to get the list
    let listExists = true;
    try {
      await this.sp.web.lists.getByTitle(listTitle)();
      // const viewTitle = "CustomView";
      // const itemLimit = 10;
      // const list = this.sp.web.lists.getByTitle("SalesRecords");
      // await list.views.add(viewTitle, false, {
      //    RowLimit: itemLimit,
      //    PersonalView: false,
      //    SetAsDefaultView: true,
      //    ViewQuery: "", // optional CAML query if you want filters
      //    ViewFields: ["Title"], // choose your fields
      //  });
      await this.sp.web.lists
        .getByTitle("CWSalesOEM")
        .views.add("Sales_View5", false, {
          RowLimit: 0,
          ViewQuery: `<Query>
       <Where>
         <Or>
           <Eq>
             <FieldRef Name="None" />
             <Value Type="Text"></Value>
           </Eq>
           <Eq>
             <FieldRef Name="None" />
             <Value Type="Text"></Value>
           </Eq>
         </Or>
       </Where>
     </Query>`,
        });
      await this.sp.web.lists
        .getByTitle("CWSalesOEM")
        .views.getByTitle("Sales_View5")
        .renderAsHtml();

      // RowLimit: 0,
      //   SetAsDefaultView: true,
      //   ViewFields: ["Title"], // make sure these internal fields exist
      //   ViewQuery: "", // Optional CAML query
      // });
    } catch {
      listExists = false;
    }

    // Create list if not exists
    if (!listExists) {
      await this.sp.web.lists.add(listTitle, "Unified Sales Data", 100, false);

      console.log(`Created list: ${listTitle}`);
    } else {
      console.log(`List '${listTitle}' already exists`);
    }
    if (!listExists) {
      const fields = [
        // {
        //   name: "RecordType",
        //   type: "Choice",
        //   choices: ["Opportunity", "Quotation", "PurchaseOrder"],
        // },
        { name: "OEM", type: "Text" },
        
      ];

      const list = this.sp.web.lists.getByTitle(listTitle);

      for (const field of fields) {
        try {
          
            await list.fields.addText(field.name);
          
          console.log(`✔ Field added: ${field.name}`);
        } catch (err) {
          console.warn(
            `⚠ Field '${field.name}' might already exist or failed:`,
            err
          );
        }
      }
    }
  };
 private CustomercreateList = async (): Promise<void> => {
  const listTitle = "CWSalesCustomer";

  // Check if the list exists
  let listExists = true;
  try {
    await this.sp.web.lists.getByTitle(listTitle)();

    // Optional: Add a custom view
    await this.sp.web.lists
      .getByTitle(listTitle)
      .views.add("Sales_View5", false, {
        RowLimit: 0,
        ViewQuery: `<Query>
         <Where>
           <Or>
             <Eq>
               <FieldRef Name="None" />
               <Value Type="Text"></Value>
             </Eq>
             <Eq>
               <FieldRef Name="None" />
               <Value Type="Text"></Value>
             </Eq>
           </Or>
         </Where>
       </Query>`,
      });

    await this.sp.web.lists
      .getByTitle(listTitle)
      .views.getByTitle("Sales_View5")
      .renderAsHtml();
  } catch {
    listExists = false;
  }

  // Create the list if it does not exist
  if (!listExists) {
    await this.sp.web.lists.add(listTitle, "Unified Sales Data", 100, false);
    console.log(`✔ Created list: ${listTitle}`);
  } else {
    console.log(`ℹ List '${listTitle}' already exists`);
  }

  // Add fields only if the list was newly created
  if (!listExists) {
    const list = this.sp.web.lists.getByTitle(listTitle);

    // Define field creation logic
    try {
      // Customer - Text field
      await list.fields.addText("Customer");
      await list.fields.addText("City");
      console.log("✔ Field added: Customer");

      // PersonResponsible - People picker field
      await list.fields.addUser("PersonResponsible", {
        Required: false,
        SelectionMode: 1, // 1 = PeopleOnly, 0 = PeopleAndGroups
      });
      console.log("✔ Field added: PersonResponsible (People Picker)");

      // City - Location field
      // await list.fields.createFieldAsXml(`
      //   <Field 
      //     DisplayName='City' 
      //     Name='City' 
      //     Type='Location' 
      //     Group='Custom Columns' 
      //     Required='FALSE' />
      // `);
      console.log("✔ Field added: City (Location)");
    } catch (err) {
      console.error("⚠ Error while adding fields:", err);
    }
  }
};

    private KeyContactcreateList = async (): Promise<void> => {
    const listTitle = "CWSalesKeyContact";

    // Try to get the list
    let listExists = true;
    try {
      await this.sp.web.lists.getByTitle(listTitle)();
      // const viewTitle = "CustomView";
      // const itemLimit = 10;
      // const list = this.sp.web.lists.getByTitle("SalesRecords");
      // await list.views.add(viewTitle, false, {
      //    RowLimit: itemLimit,
      //    PersonalView: false,
      //    SetAsDefaultView: true,
      //    ViewQuery: "", // optional CAML query if you want filters
      //    ViewFields: ["Title"], // choose your fields
      //  });
      await this.sp.web.lists
        .getByTitle("CWSalesKeyContact")
        .views.add("Sales_View5", false, {
          RowLimit: 0,
          ViewQuery: `<Query>
       <Where>
         <Or>
           <Eq>
             <FieldRef Name="None" />
             <Value Type="Text"></Value>
           </Eq>
           <Eq>
             <FieldRef Name="None" />
             <Value Type="Text"></Value>
           </Eq>
         </Or>
       </Where>
     </Query>`,
        });
      await this.sp.web.lists
        .getByTitle("CWSalesKeyContact")
        .views.getByTitle("Sales_View5")
        .renderAsHtml();

      // RowLimit: 0,
      //   SetAsDefaultView: true,
      //   ViewFields: ["Title"], // make sure these internal fields exist
      //   ViewQuery: "", // Optional CAML query
      // });
    } catch {
      listExists = false;
    }

    // Create list if not exists
    if (!listExists) {
      await this.sp.web.lists.add(listTitle, "Unified Sales Data", 100, false);

      console.log(`Created list: ${listTitle}`);
    } else {
      console.log(`List '${listTitle}' already exists`);
    }
    if (!listExists) {
      const fields = [
        // {
        //   name: "RecordType",
        //   type: "Choice",
        //   choices: ["Opportunity", "Quotation", "PurchaseOrder"],
        // },
        { name: "Customer", type: "Text" },
        { name: "Contact", type: "Text" },
        { name: "Email", type: "Text" },
        { name: "Address", type: "Text" },
        { name: "City", type: "Text" },
        { name: "BusinessPhone", type: "Text" },
        { name: "MobileNumber", type: "Text" },
        { name: "Designation", type: "Text" },
        { name: "Department", type: "Text" },
       
      ];

      const list = this.sp.web.lists.getByTitle(listTitle);

      for (const field of fields) {
        try {
         
            await list.fields.addText(field.name);
          
          console.log(`✔ Field added: ${field.name}`);
        } catch (err) {
          console.warn(
            `⚠ Field '${field.name}' might already exist or failed:`,
            err
          );
        }
      }
    }
  };
  private ConfigcreateList = async (): Promise<void> => {
  const listTitle = "CWSalesConfiguration";

  let listExists = true;
  try {
    await this.sp.web.lists.getByTitle(listTitle)();
    
    await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .views.add("Sales_View5", false, {
        RowLimit: 0,
        ViewQuery: `<Query>
          <Where>
            <Or>
              <Eq>
                <FieldRef Name="None" />
                <Value Type="Text"></Value>
              </Eq>
              <Eq>
                <FieldRef Name="None" />
                <Value Type="Text"></Value>
              </Eq>
            </Or>
          </Where>
        </Query>`,
      });

    await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .views.getByTitle("Sales_View5")
      .renderAsHtml();
  } catch {
    listExists = false;
  }

  // Create the list if it doesn't exist
  if (!listExists) {
    await this.sp.web.lists.add(listTitle, "Unified Sales Data", 100, false);
    console.log(`Created list: ${listTitle}`);
  } else {
    console.log(`List '${listTitle}' already exists`);
  }

  if (!listExists) {
    const list = this.sp.web.lists.getByTitle(listTitle);

    // Add multiline text field
    try {
      await list.fields.addText("DefaultCurrency");
      await list.fields.addText("Value");
     await list.fields.createFieldAsXml(
  `<Field Type="Note" Name="MultiValue" DisplayName="MultiValue" RichText="TRUE" RichTextMode="FullHtml" NumLines="6" />`
);
      console.log(`✔ Field added: MultiValue`);
    } catch (err) {
      console.warn(`⚠ Field 'MultiValue' might already exist or failed:`, err);
    }
const PageSetting =
      JSON.parse(`{"roleBasedAccess":false,"hideSiteHeader":true,"hideCommentsWrapper":true,"hideSideAppBar":true,"hideO365BrandNavbar":true,"hidePageTitle":true,"isHoursTrackable":true,"hideSharepointHubNavbar":true,"hideCommandBar":true,"customFields":[]}`)
   
    // Create default item with hardcoded values
    try {
      await list.items.add({
        Title: "Settings",
        MultiValue: PageSetting,
      });
      const response = await fetch("https://api.exchangerate-api.com/v4/latest/EUR");
      const CurrencyData = await response.json();
       await list.items.add({
        Title: "Currency",
        MultiValue:CurrencyData ,
        DefaultCurrency:"EUR"
      });
      console.log("✔ Default config item created");
    } catch (err) {
      console.warn("⚠ Failed to create default item:", err);
    }
  }
};
  private AuditLogcreateList = async (): Promise<void> => {
  const listTitle = "CWSalesAuditLog";

  let listExists = true;
  try {
    await this.sp.web.lists.getByTitle(listTitle)();
    
    await this.sp.web.lists
      .getByTitle("CWSalesAuditLog")
      .views.add("Sales_View5", false, {
        RowLimit: 0,
        ViewQuery: `<Query>
          <Where>
            <Or>
              <Eq>
                <FieldRef Name="None" />
                <Value Type="Text"></Value>
              </Eq>
              <Eq>
                <FieldRef Name="None" />
                <Value Type="Text"></Value>
              </Eq>
            </Or>
          </Where>
        </Query>`,
      });

    await this.sp.web.lists
      .getByTitle("CWSalesAuditLog")
      .views.getByTitle("Sales_View5")
      .renderAsHtml();
  } catch {
    listExists = false;
  }

  // Create the list if it doesn't exist
  if (!listExists) {
    await this.sp.web.lists.add(listTitle, "Unified Sales Data", 100, false);
    console.log(`Created list: ${listTitle}`);
  } else {
    console.log(`List '${listTitle}' already exists`);
  }

  if (!listExists) {
    const list = this.sp.web.lists.getByTitle(listTitle);

    // Add multiline text field
    try {
      await list.fields.addText("Timestamp");
      await list.fields.addText("Action");
      await list.fields.addText("OpportunityID");
      await list.fields.addText("ModifiedBy");
      await list.fields.addMultilineText("DataSnapshot");
      await list.fields.addMultilineText("Comments");
      console.log(`✔ Field added: MultiValue`);
    } catch (err) {
      console.warn(`⚠ Field 'MultiValue' might already exist or failed:`, err);
    }

    // Create default item with hardcoded values

  }
};
  // private handleOpportunityChange = (field: string, value: any): void => {
  //   this.setState((prevState) => ({
  //     opportunityForm: {
  //       ...prevState.opportunityForm,
  //       [field]: value,
  //     },
  //   }));
  // };
  // private handleQuotationChange = (field: string, value: any): void => {
  //   this.setState((prevState) => ({
  //     quotationForm: {
  //       ...prevState.quotationForm,
  //       [field]: value,
  //     },
  //   }));
  // };
  // private handleQuotationFileChange = (
  //   e: React.ChangeEvent<HTMLInputElement>
  // ) => {
  //   const file = e.target.files?.[0];
  //   this.setState({ quotationFile: file });
  // };
  // private handlePOChanges = (field: string, value: any): void => {
  //   this.setState((prevState) => ({
  //     purchaseOrderForm: {
  //       ...prevState.purchaseOrderForm,
  //       [field]: value,
  //     },
  //   }));
  // };
  // private handleLineItemChange = (index: number, field: string, value: any) => {
  //   const items = [...this.state.lineItems];
  //   if (field in items[index]) {
  //     (items[index] as any)[field] = value;
  //   }
  //   this.setState({ lineItems: items });
  // };

  // private addLineItem = () => {
  //   this.setState((prev) => ({
  //     lineItems: [...prev.lineItems, { Title: "", Comments: "", Value: 0 }],
  //   }));
  // };
  // private submitOpportunity = async (): Promise<void> => {
  //   try {
  //     await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.add(this.state.opportunityForm);
  //     alert("Opportunity saved successfully!");
  //   } catch (error) {
  //     console.error("Error saving Opportunity:", error);
  //     alert("Failed to save Opportunity.");
  //   }
  // };
  // private submitQuotation = async (): Promise<void> => {
  //   const { quotationFile } = this.state;
  //   const {
  //     OpportunityID,
  //     QuoteID,
  //     QuoteDate,
  //     RevisionNumber,
  //     QuoteRevenueQuoted,
  //   } = this.state.quotationForm;

  //   if (!OpportunityID) {
  //     alert("Please select an Opportunity ID.");
  //     return;
  //   }

  //   try {
  //     // Get the existing Opportunity record
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.filter(`OpportunityID eq '${OpportunityID}'`)
  //       .top(1)();

  //     if (items.length === 0) {
  //       alert("No Opportunity found with this ID.");
  //       return;
  //     }

  //     const item = items[0];
  //     const fieldsToUpdate: any = {};

  //     // Check which QuoteID slot is available
  //     for (let i = 1; i <= 5; i++) {
  //       const suffix = i === 1 ? "" : i.toString();
  //       if (!item[`QuoteID${suffix}`]) {
  //         fieldsToUpdate[`QuoteID${suffix}`] = QuoteID || `Q-${Date.now()}`;
  //         fieldsToUpdate[`QuoteDate${suffix}`] = QuoteDate;
  //         fieldsToUpdate[`QuoteRevisionNumber${suffix}`] = RevisionNumber;
  //         fieldsToUpdate[`QuoteRevenueQuoted${suffix}`] = QuoteRevenueQuoted;
  //         break;
  //       }
  //     }

  //     if (Object.keys(fieldsToUpdate).length === 0) {
  //       alert("All quotation slots are full (max 5).");
  //       return;
  //     }

  //     await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.getById(item.Id)
  //       .update(fieldsToUpdate);
  //     if (quotationFile) {
  //       const folderName = `${`Q${Date.now()}`}`;
  //       await this.sp.web.folders.addUsingPath(
  //         `Shared Documents/${folderName}`
  //       );
  //       const fileBuffer = await quotationFile.arrayBuffer();
  //       await this.sp.web
  //         .getFolderByServerRelativePath(`Shared Documents/${folderName}`)
  //         .files.addUsingPath(quotationFile.name, fileBuffer, {
  //           Overwrite: true,
  //         });

  //       alert("File uploaded to document library.");
  //     }

  //     alert("Quotation saved successfully.");
  //   } catch (err) {
  //     console.error("Error saving quotation", err);
  //     alert("Failed to save quotation.");
  //   }
  // };
  // private submitPurchaseOrder = async (): Promise<void> => {
  //   const form = this.state.purchaseOrderForm;
  //   const opportunityID = form.OpportunityID;

  //   if (!opportunityID) {
  //     alert("Please select an Opportunity ID.");
  //     return;
  //   }

  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.filter(`OpportunityID eq '${opportunityID}'`)
  //       .top(1)();

  //     if (items.length === 0) {
  //       alert("No Opportunity found with this ID.");
  //       return;
  //     }

  //     const item = items[0];
  //     const fieldsToUpdate: any = {};

  //     // Check which POID slot is available
  //     for (let i = 1; i <= 5; i++) {
  //       const suffix = i === 1 ? "" : i.toString();
  //       if (!item[`POID${suffix}`]) {
  //         fieldsToUpdate[`POID${suffix}`] = form.POID || `PO-${Date.now()}`;
  //         fieldsToUpdate[`POReceivedDate${suffix}`] = form.POReceivedDate;
  //         fieldsToUpdate[`CustomerPONumber${suffix}`] = form.CustomerPONumber;
  //         fieldsToUpdate[`QuoteID${suffix}`] = form.QuoteID;
  //         fieldsToUpdate[`POStatus${suffix}`] = form.POStatus;
  //         fieldsToUpdate[`AmountEUR${suffix}`] = form.AmountEUR;
  //         fieldsToUpdate[`Currency${suffix}`] = form.Currency;
  //         fieldsToUpdate[`LineItemsJSON${suffix}`] = JSON.stringify(
  //           this.state.lineItems
  //         );
  //         break;
  //       }
  //     }

  //     if (Object.keys(fieldsToUpdate).length === 0) {
  //       alert("All PO slots are full.");
  //       return;
  //     }

  //     await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.getById(item.Id)
  //       .update(fieldsToUpdate);

  //     alert("Purchase Order saved successfully.");
  //   } catch (err) {
  //     console.error("Error saving Purchase Order:", err);
  //     alert("Failed to save Purchase Order.");
  //   }
  // };
  private renderComponent = () => {
    return (
      <div>
        {/* {this.state.tab === "Opportunity" && this.renderOpportunityForm()}
        {this.state.tab === "Quotation" && this.renderQuotationForm()}
        {this.state.tab === "Purchase Order" && this.renderPurchaseOrderForm()} */}

        {this.state.tab === "View Opportunities" && (
          <div style={{ marginTop: 24 }}>
            <OpportunityViewer
              context={this.props.context}
              salesProps={this.props}
            />
          </div>
        )}
        {this.state.tab === "View Quotation" && (
          <div style={{ marginTop: 24 }}>
            <QuotationViewer
              context={this.props.context}
              salesProps={this.props}
            />
          </div>
        )}
        {this.state.tab === "View PO" && (
          <div style={{ marginTop: 24 }}>
            <POViewer context={this.props.context} />
          </div>
        )}
         {this.state.tab === "ViewSalesDashboard" && (
          <div style={{ marginTop: 24 }}>
            <SalesDashboardComponent context={this.props.context} description={""} isDarkTheme={false} environmentMessage={""} hasTeamsContext={false} userDisplayName={""} sp={new SPFI} View={""} settings={{}} onConfigChange={function (config: ISettings): void {
              throw new Error("Function not implemented.");
            } } />
          </div>
        )}
        {this.state.tab === "Setting" && (
          <div style={{ marginTop: 24 }}>
            <TaskContext.Provider value={this.contextObject}>
              <Settings context={this.contextObject} />
            </TaskContext.Provider>
          </div>
        )}
      </div>
    );
  };
  //   private renderTabs = () => (

  // {/* <Pivot
  //   selectedKey={this.state.tab}
  //   onLinkClick={(item) => this.setTab(item?.props.itemKey || 'Opportunity')}
  // >
  //   <PivotItem headerText="Opportunity" itemKey="Opportunity" />
  //   <PivotItem headerText="Quotation" itemKey="Quotation" />
  //   <PivotItem headerText="Purchase Order" itemKey="Purchase Order" />
  // </Pivot> */}
  //   );
  // private renderOpportunityList = () => (
  //   <div style={{ marginTop: '1rem' }}>
  //     <h3>Existing Opportunities</h3>
  //     <ul>
  //       {this.state.opportunityOptions.map((o, i) => (
  //         <li key={i}>{o.text}</li>
  //       ))}
  //     </ul>
  //   </div>
  // );
  // private renderOpportunityForm = () => (
  //   // <div>
  //   //   <TextField label="Business" onChange={(_, v) => this.handleOpportunityChange("Business", v)} />
  //   //   <TextField label="Business Unit" onChange={(_, v) => this.handleOpportunityChange("BusinessUnit", v)} />
  //   //   <TextField label="OEM" onChange={(_, v) => this.handleOpportunityChange("OEM", v)} />
  //   //   <TextField label="End Customer" onChange={(_, v) => this.handleOpportunityChange("EndCustomer", v)} />
  //   //   <TextField label="Customer" onChange={(_, v) => this.handleOpportunityChange("Customer", v)} />
  //   //   <TextField label="Key Contact" onChange={(_, v) => this.handleOpportunityChange("KeyContact", v)} />
  //   //   <TextField label="Decision Maker" onChange={(_, v) => this.handleOpportunityChange("DecisionMaker", v)} />
  //   //   <DatePicker label="Report Date" onSelectDate={(d) => this.handleOpportunityChange("ReportDate", d)} />
  //   //   <DatePicker label="Tentative Start Date" onSelectDate={(d) => this.handleOpportunityChange("TentativeStartDate", d)} />
  //   //   <DatePicker label="Tentative Decision Date" onSelectDate={(d) => this.handleOpportunityChange("TentativeDecisionDate", d)} />
  //   //   <TextField label="Amount (EUR)" type="number" onChange={(_, v) => this.handleOpportunityChange("Amount", parseFloat(v || "0"))} />
  //   //   <Dropdown label="Currency" options={currencies} onChange={(_, o) => this.handleOpportunityChange("Currency", o?.key)} />
  //   //   <TextField label="Converted Amount" disabled />
  //   //   <Dropdown label="Risk Level" options={riskLevels} onChange={(_, o) => this.handleOpportunityChange("RiskLevel", o?.key)} />
  //   //   <Dropdown label="Strategic" options={strategicOptions} onChange={(_, o) => this.handleOpportunityChange("Strategic", o?.key)} />
  //   //   <Dropdown label="Opportunity Status" options={statusOptions} onChange={(_, o) => this.handleOpportunityChange("Status", o?.key)} />
  //   //   <TextField label="Comments" multiline rows={3} onChange={(_, v) => this.handleOpportunityChange("Comments", v)} />
  //   //   <PrimaryButton text="Submit Opportunity" onClick={this.submitOpportunity} />
  //   //   <PrimaryButton text="View Opportunities" onClick={() => <OpportunityViewer context={this.props.context} />} className="mt-2" />
  //   //   {<div><OpportunityViewer context={this.props.context} /></div>}
  //   // </div>
  //   <div style={{ padding: 16 }}>
  //     <div style={gridStyle}>
  //       <TextField
  //         label="Opportunity ID"
  //         onChange={(_, v) => this.handleOpportunityChange("OpportunityID", v)}
  //       />
  //       <TextField
  //         label="Business"
  //         onChange={(_, v) => this.handleOpportunityChange("Business", v)}
  //       />
  //       <TextField
  //         label="Business Unit"
  //         onChange={(_, v) => this.handleOpportunityChange("BusinessUnit", v)}
  //       />
  //       <TextField
  //         label="OEM"
  //         onChange={(_, v) => this.handleOpportunityChange("OEM", v)}
  //       />
  //       <TextField
  //         label="End Customer"
  //         onChange={(_, v) => this.handleOpportunityChange("EndCustomer", v)}
  //       />
  //       <TextField
  //         label="Customer"
  //         onChange={(_, v) => this.handleOpportunityChange("Customer", v)}
  //       />
  //       <TextField
  //         label="Key Contact"
  //         onChange={(_, v) => this.handleOpportunityChange("KeyContact", v)}
  //       />
  //       <TextField
  //         label="Decision Maker"
  //         onChange={(_, v) => this.handleOpportunityChange("DecisionMaker", v)}
  //       />
  //       <DatePicker
  //         label="Report Date"
  //         onSelectDate={(d) => this.handleOpportunityChange("ReportDate", d)}
  //       />
  //       <DatePicker
  //         label="Tentative Start Date"
  //         onSelectDate={(d) =>
  //           this.handleOpportunityChange("TentativeStartDate", d)
  //         }
  //       />
  //       <DatePicker
  //         label="Tentative Decision Date"
  //         onSelectDate={(d) =>
  //           this.handleOpportunityChange("TentativeDecisionDate", d)
  //         }
  //       />
  //       <TextField
  //         label="Amount (EUR)"
  //         type="number"
  //         onChange={(_, v) =>
  //           this.handleOpportunityChange("Amount", parseFloat(v || "0"))
  //         }
  //       />
  //       <Dropdown
  //         label="Currency"
  //         options={currencies}
  //         onChange={(_, o) => this.handleOpportunityChange("Currency", o?.key)}
  //       />
  //       <TextField label="Converted Amount" disabled />
  //       <Dropdown
  //         label="Risk Level"
  //         options={riskLevels}
  //         onChange={(_, o) => this.handleOpportunityChange("RiskLevel", o?.key)}
  //       />
  //       <Dropdown
  //         label="Strategic"
  //         options={strategicOptions}
  //         onChange={(_, o) => this.handleOpportunityChange("Strategic", o?.key)}
  //       />
  //       <Dropdown
  //         label="Opportunity Status"
  //         options={statusOptions}
  //         onChange={(_, o) => this.handleOpportunityChange("Status", o?.key)}
  //       />
  //       <TextField
  //         label="Comments"
  //         multiline
  //         rows={3}
  //         onChange={(_, v) => this.handleOpportunityChange("Comments", v)}
  //       />
  //     </div>

  //     {/* Action Buttons */}
  //     <Stack horizontal tokens={{ childrenGap: 12 }} style={{ marginTop: 24 }}>
  //       <PrimaryButton
  //         text="Submit Opportunity"
  //         onClick={this.submitOpportunity}
  //         styles={primaryButtonStyles}
  //       />
  //       {/* <PrimaryButton
  //           text="View Opportunities"
  //           onClick={() => this.setState({ tab: 'View Opportunities' })}
  //           styles={primaryButtonStyles}
  //         /> */}
  //     </Stack>

  //     {/* Optional Opportunity Viewer */}
  //     {/* <div style={{ marginTop: 24 }}>
  //         <OpportunityViewer context={this.props.context} />
  //       </div> */}
  //   </div>
  // );

  // private renderQuotationForm = () => (
  //   //     <div>
  //   //       <TextField label="Quote ID" disabled onChange={(_, v) => this.handleQuotationChange("QuoteID", v)}/>
  //   //       <Dropdown
  //   //   label="Opportunity ID"
  //   //   options={this.state.opportunityOptions}
  //   //   onChange={(_, option) => this.handleQuotationChange("OpportunityID", option?.key)}
  //   // />
  //   //       <Dropdown label="Business Size" options={businessSizes} onChange={(_, v) => this.handleQuotationChange("BusinessSize", v)}/>
  //   //       <DatePicker label="Quote Date" onSelectDate={(d) => this.handleQuotationChange("QuoteDate", d)}/>
  //   //       <DatePicker label="Tentative Decision Date" onSelectDate={(d) => this.handleQuotationChange("TentativeDecisionDate", d)}/>
  //   //       <TextField label="Revision Number" onChange={(_, v) => this.handleQuotationChange("RevisionNumber", v)}/>
  //   //       <TextField label="Amount (EUR)" type="number" onChange={(_, v) => this.handleQuotationChange("AmmountEUR", v)}/>
  //   //       <Dropdown label="Currency" options={currencies} onChange={(_, v) => this.handleQuotationChange("Currency", v)}/>
  //   //       <TextField label="Revenue Quoted" disabled onChange={(_, v) => this.handleQuotationChange("QuoteRevenueQuoted", v)}/>
  //   //       <input type="file" onChange={this.handleQuotationFileChange} />

  //   //       <PrimaryButton text="Submit Quotation" onClick={this.submitQuotation}/>
  //   //     </div>
  //   <div style={{ padding: 16 }}>
  //     <div style={columnGridStyle.root}>
  //       <TextField
  //         label="Quote ID"
  //         disabled
  //         onChange={(_, v) => this.handleQuotationChange("QuoteID", v)}
  //       />
  //       <Dropdown
  //         label="Opportunity ID"
  //         options={this.state.opportunityOptions}
  //         onChange={(_, option) =>
  //           this.handleQuotationChange("OpportunityID", option?.key)
  //         }
  //       />
  //       <Dropdown
  //         label="Business Size"
  //         options={businessSizes}
  //         onChange={(_, option) =>
  //           this.handleQuotationChange("BusinessSize", option?.key)
  //         }
  //       />
  //       <DatePicker
  //         label="Quote Date"
  //         onSelectDate={(date) => this.handleQuotationChange("QuoteDate", date)}
  //       />
  //       <DatePicker
  //         label="Tentative Decision Date"
  //         onSelectDate={(date) =>
  //           this.handleQuotationChange("TentativeDecisionDate", date)
  //         }
  //       />
  //       <TextField
  //         label="Revision Number"
  //         onChange={(_, v) => this.handleQuotationChange("RevisionNumber", v)}
  //       />
  //       <TextField
  //         label="Amount (EUR)"
  //         type="number"
  //         onChange={(_, v) => this.handleQuotationChange("AmmountEUR", v)}
  //       />
  //       <Dropdown
  //         label="Currency"
  //         options={currencies}
  //         onChange={(_, option) =>
  //           this.handleQuotationChange("Currency", option?.key)
  //         }
  //       />
  //       <TextField
  //         label="Revenue Quoted"
  //         disabled
  //         onChange={(_, v) =>
  //           this.handleQuotationChange("QuoteRevenueQuoted", v)
  //         }
  //       />

  //       {/* File Upload Field */}
  //       <Stack tokens={{ childrenGap: 8 }}>
  //         <Label>Attachment</Label>
  //         <input type="file" onChange={this.handleQuotationFileChange} />
  //       </Stack>
  //     </div>

  //     {/* Submit Button */}
  //     <Stack horizontalAlign="start" styles={{ root: { marginTop: 20 } }}>
  //       <PrimaryButton
  //         text="Submit Quotation"
  //         onClick={this.submitQuotation}
  //         styles={primaryButtonStyles}
  //       />
  //       {/* <PrimaryButton
  //           text="View Opportunities"
  //           onClick={() => this.setState({ tab: 'View Quotation' })}
  //           styles={primaryButtonStyles}
  //         /> */}
  //     </Stack>
  //   </div>
  // );

  // private renderPurchaseOrderForm = () => (
  //   //     <div>
  //   // <Dropdown
  //   //   label="Opportunity ID"
  //   //   options={this.state.opportunityOptions}
  //   //   onChange={(_, option) => {
  //   //     this.handlePOChanges("OpportunityID", option?.key);
  //   //     // eslint-disable-next-line no-void
  //   //     if (option?.key) void this.fetchQuotesForOpportunity(option.key.toString());
  //   //   }}
  //   // />
  //   //       <Dropdown
  //   //   label="Quote ID"
  //   //   options={this.state.quoteOptions}
  //   //   onChange={(_, option) => this.handlePOChanges("QuoteID", option?.key)}
  //   // />

  //   //       <TextField label="PO ID" disabled />
  //   //       <Dropdown label="Is Child PO" options={[{ key: 'true', text: 'Yes' }, { key: 'false', text: 'No' }]} onChange={(_,v)=>this.handlePOChanges("IsChildPO",v)}/>
  //   //       <TextField label="Parent PO ID" />
  //   //       <DatePicker label="PO Received Date" onSelectDate={(d)=>this.handlePOChanges("POReceivedDate",d)}/>
  //   //       <Dropdown label="PO Status" options={[{ key: 'Draft', text: 'Draft' }, { key: 'Issued', text: 'Issued' }, { key: 'Approved', text: 'Approved' }, { key: 'Cancelled', text: 'Cancelled' }] } onChange={(_,v)=>this.handlePOChanges("POStatus",v?.text)}/>
  //   //       <TextField label="Customer PO Number" onChange={(_,v)=>this.handlePOChanges("CustomerPONumber",v)}/>
  //   //       <TextField label="PO Value (EUR)" type="number" onChange={(_,v)=>this.handlePOChanges("POValue",v)} />
  //   //       {this.state.lineItems.map((item, index) => (
  //   //   <div key={index} style={{ marginBottom: 12, border: '1px solid #ccc', padding: 10, borderRadius: 4 }}>
  //   //     <TextField
  //   //       label={`Title`}
  //   //       value={item.Title}
  //   //       onChange={(_, v) => this.handleLineItemChange(index, 'Title', v)}
  //   //     />
  //   //     <TextField
  //   //       label={`Comments`}
  //   //       value={item.Comments}
  //   //       onChange={(_, v) => this.handleLineItemChange(index, 'Comments', v)}
  //   //     />
  //   //     <TextField
  //   //       label={`Value`}
  //   //       type="number"
  //   //       value={item.Value.toString()}
  //   //       onChange={(_, v) => this.handleLineItemChange(index, 'Value', parseFloat(v || '0'))}
  //   //     />
  //   //   </div>
  //   // ))}

  //   //       <PrimaryButton text="View POs" onClick={this.fetchPurchaseOrders} className="mr-2" />

  //   // <PrimaryButton text="Submit Purchase Order" onClick={this.submitPurchaseOrder} />
  //   // {this.renderPurchaseOrderList()}
  //   // <PrimaryButton text="+ Add Line Item" onClick={this.addLineItem} />

  //   //       {/* <PrimaryButton text="Submit Purchase Order" onClick={this.submitPurchaseOrder}/> */}
  //   //     </div>
  //   <div style={{ padding: 16 }}>
  //     <div style={gridStyle}>
  //       <Dropdown
  //         label="Opportunity ID"
  //         options={this.state.opportunityOptions}
  //         onChange={(_, option) => {
  //           this.handlePOChanges("OpportunityID", option?.key);
  //           if (option?.key)
  //             void this.fetchQuotesForOpportunity(option.key.toString());
  //         }}
  //       />
  //       <Dropdown
  //         label="Quote ID"
  //         options={this.state.quoteOptions}
  //         onChange={(_, option) => this.handlePOChanges("QuoteID", option?.key)}
  //       />
  //       <TextField label="PO ID" disabled />
  //       <Dropdown
  //         label="Is Child PO"
  //         options={[
  //           { key: "true", text: "Yes" },
  //           { key: "false", text: "No" },
  //         ]}
  //         onChange={(_, v) => this.handlePOChanges("IsChildPO", v?.key)}
  //       />
  //       <TextField
  //         label="Parent PO ID"
  //         onChange={(_, v) => this.handlePOChanges("ParentPOID", v)}
  //       />
  //       <DatePicker
  //         label="PO Received Date"
  //         onSelectDate={(d) => this.handlePOChanges("POReceivedDate", d)}
  //       />
  //       <Dropdown
  //         label="PO Status"
  //         options={[
  //           { key: "Draft", text: "Draft" },
  //           { key: "Issued", text: "Issued" },
  //           { key: "Approved", text: "Approved" },
  //           { key: "Cancelled", text: "Cancelled" },
  //         ]}
  //         onChange={(_, v) => this.handlePOChanges("POStatus", v?.text)}
  //       />
  //       <TextField
  //         label="Customer PO Number"
  //         onChange={(_, v) => this.handlePOChanges("CustomerPONumber", v)}
  //       />
  //       <TextField
  //         label="PO Value (EUR)"
  //         type="number"
  //         onChange={(_, v) => this.handlePOChanges("POValue", v)}
  //       />
  //     </div>
  //     {/* Line Items */}
  //     <div style={sectionHeaderStyle}>Line Items *</div>
  //     {this.state.lineItems.map((item: any, index: number) => (
  //       <div key={index} style={lineItemStyle}>
  //         <TextField
  //           placeholder="Title"
  //           value={item.Title}
  //           onChange={(_, v) => this.handleLineItemChange(index, "Title", v)}
  //         />
  //         <TextField
  //           placeholder="Comment"
  //           value={item.Comments}
  //           onChange={(_, v) => this.handleLineItemChange(index, "Comments", v)}
  //         />
  //         <TextField
  //           placeholder="Value"
  //           type="number"
  //           value={item.Value.toString()}
  //           onChange={(_, v) =>
  //             this.handleLineItemChange(index, "Value", parseFloat(v || "0"))
  //           }
  //         />
  //         <IconButton
  //           iconProps={{ iconName: "Cancel" }}
  //           title="Remove"
  //           ariaLabel="Remove"
  //           onClick={() => this.removeLineItem(index)}
  //         />
  //       </div>
  //     ))}
  //     <div style={{ marginTop: 16 }}>
  //           <PrimaryButton
  //       text="+ Add Line Item"
  //       onClick={this.addLineItem}
  //       styles={primaryButtonStyles}
  //     />
  //     </div>

  //     {/* Action Buttons */}
  //     <Stack horizontal tokens={{ childrenGap: 12 }} style={{ marginTop: 24 }}>
  //       <PrimaryButton
  //         text="View POs"
  //         onClick={this.fetchPurchaseOrders}
  //         styles={primaryButtonStyles}
  //       />
  //       <PrimaryButton
  //         text="Submit Purchase Order"
  //         onClick={this.submitPurchaseOrder}
  //         styles={primaryButtonStyles}
  //       />
  //       <PrimaryButton
  //         text="View PO"
  //         onClick={() => this.setState({ tab: "View PO" })}
  //         styles={primaryButtonStyles}
  //       />
  //     </Stack>

  //     {/* Optional List Render */}
  //     <div style={{ marginTop: 24 }}>{this.renderPurchaseOrderList()}</div>
  //   </div>
  // );
  private setPageSettings() {
    const settings = this.contextObject.config.settings;
    console.log("Settings", settings);
    if (settings) {
      console.log("Settings", settings);
      pageSettings.forEach((ps) => {
        const isHidden = settings[ps.stateVariable];
        const HTMLElement: any = document.querySelector(
          `${ps.sharepointElement}`
        );
        if (HTMLElement) {
          HTMLElement.style.setProperty(
            "display",
            isHidden ? "none" : "block",
            "important"
          );
        }
      });
    }
  }
  public render(): React.ReactElement<ISalesProps> {
    return (
      // <div className="p-4">
     <div style={{ width: "100%", height: "100vh" }} id="sales-webpart-root">
        {/* {
              this.props.View === "Opportunity" && (
                  <div>
                    {this.renderOpportunityForm()}
                  </div>
              )
        } */}

        <Stack
          horizontal
          tokens={{ childrenGap: 10 }}
          styles={{ root: { height: `100%` } }}
        >
          <Stack.Item
            grow={false}
            style={{
               width: this.state.isNavOpen ? '225px' : "55px",
              borderRight: "1px solid #ccc",
              transition: "ease-in-out .25s",
              display: "flex",
              flexDirection: "column",
            }}
          >
            {/* Add content here if needed */}

            {/* Logo & Title */}
            <div
              style={{
                display: "grid",
                gap: "0.5em",
                alignItems: "center",
                gridTemplateColumns: "auto 1fr",
              }}
            >
              <Logo height={40} width={40} isNavOpen={true} />
              <p
                style={{
                  margin: 0,
                  display: "flex",
                  flexDirection: "column",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: "1rem",
                  fontWeight: 500,
                  userSelect: "none",
                }}
              >
                 {this.state.isNavOpen ? <p
            style={{
              margin: 0,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "1rem",
              fontWeight: 500,
              userSelect: "none",
            }}>
            <label>Sales</label>
           
          </p> : null}
                {/* <label>Tracker</label> */}
              </p>
            </div>
            <div style={{ padding: "1em 0 0" }}>
          <IconButton
            style={{
              display: "block",
              marginLeft: this.state.isNavOpen ? "auto" : "0",
              marginRight: "0.5em",
              transition: "ease-in-out .25s"
            }}
            iconProps={{
              iconName: this.state.isNavOpen ? "Clear" : "DensityComfy"
            }}
            onClick={() => this.setState((prevState) => ({
              isNavOpen: !prevState.isNavOpen,
            }))}
          />
        </div>
            <div style={{ overflow: "auto" }} className="customNavigation">
              <Nav
                groups={navLinks}
                // selectedKey={this.state.selectedKey}
                onLinkClick={this.onLinkCLick}
              />
            </div>
            <div
              style={{
                display: "block",
                marginTop: "auto",
                padding: "0.5em 0.75em",
                marginRight: "0.5em",
                background: "white",
              }}
            >
              <LivePersona
                upn={this.props.context.pageContext.user?.email?.toLowerCase()}
                template={
                  <Persona
                    coinSize={24}
                    text={this.props.context.pageContext.user?.displayName}
                    imageUrl={this.imageUrl}
                    hidePersonaDetails={false}
                  />
                }
                serviceScope={this.props.context.serviceScope as any}
              />
            </div>
          </Stack.Item>
          <Stack.Item style={{ display: "grid", flex: 1 }}>
            <h2 className="text-xl font-bold mb-4">Sales Tracker</h2>
            {/* {this.renderTabs()} */}
            {this.renderComponent()}
          </Stack.Item>
        </Stack>
        {/* {this.renderOpportunityList()} */}
      </div>
    );
  }
}
