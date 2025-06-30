/* eslint-disable react/self-closing-comp */
/* eslint-disable prefer-const */
/* eslint-disable eqeqeq */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-void */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from "react";
import {
  DetailsList,
  SelectionMode,
  IColumn,
} from "@fluentui/react/lib/DetailsList";
import { CommandBar } from "@fluentui/react/lib/CommandBar";
import { Panel, PanelType } from "@fluentui/react/lib/Panel";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { TextField } from "@fluentui/react/lib/TextField";
import * as XLSX from "xlsx";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  getTheme,
  IButtonStyles,
  Icon,
  IconButton,
  IDropdownOption,
  Label,
  PrimaryButton,
  Stack,
} from "@fluentui/react";
import DropZoneUploader from "../DropZoneUploader";
import DocumentViewer from "../DocumentViewer";
import ColumnsViewPanel from "../ColumnsViewSettings/ColumnsViewPanel";
import { IPOViewerProps } from "./IViewPOProps";
import { IPOViewerState, PO } from "./IViewPOState";

// Define PO  type
// type PO = {
//   lineItems: any;
//   // lineItems(lineItems: any): unknown;
//   POQuoteID:string;
//   POID: string;
//   OpportunityID: string;
//   POReceivedDate: Date | string;
//   POStatus: string;
//   POAmount: string;
//   POCurrency: string;
//   CustomerPONumber: string;
//   lineItemsJSON: string;
//   POQuoteId: string;
//   POComments: string;
//   IsChildPO: string;
// };

//const mockPO: PO[] = [];
const gridStyle = {
  display: "grid",
  gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
  gap: "16px",
};
const lineItemStyle = {
  display: "grid",
  gridTemplateColumns: "1fr 1fr 1fr 50px",
  gap: 12,
  alignItems: "center",
  marginTop: 12,
};
// const columns: IColumn[] = [
//   {
//     key: "poId",
//     name: "PO Number",
//     fieldName: "poId",
//     minWidth: 100,
//     isResizable: true,
//   },
//   { key: "date", name: "PO Received Date", fieldName: "date", minWidth: 100 },
//   { key: "status", name: "PO Status", fieldName: "status", minWidth: 90 },
//   { key: "amount", name: "Amount EUR", fieldName: "amount", minWidth: 100 },
// ];
const theme = getTheme();

const primaryButtonStyles: IButtonStyles = {
  root: {
    backgroundColor: theme.palette.themePrimary,
    borderColor: theme.palette.themePrimary,
  },
  rootHovered: {
    backgroundColor: theme.palette.themeDark,
    borderColor: theme.palette.themeDark,
  },
  rootPressed: {
    backgroundColor: theme.palette.themeDarker,
    borderColor: theme.palette.themeDarker,
  },
};
const sectionHeaderStyle = {
  backgroundColor: "#176D7E",
  color: "#fff",
  fontWeight: 600,
  fontSize: 20,
  padding: "8px 16px",
  borderRadius: "4px",
  marginTop: 24,
};
const formatDate = (dateInput: string | Date, format: string): string => {
  const date = typeof dateInput === "string" ? new Date(dateInput) : dateInput;

  const map = {
    DD: date.getDate().toString().padStart(2, "0"),
    MMM: date.toLocaleString("en-US", { month: "short" }),
    MM: (date.getMonth() + 1).toString().padStart(2, "0"),
    YYYY: date.getFullYear().toString(),
  };

  return format
    .replace(/DD/, map.DD)
    .replace(/MMM/, map.MMM)
    .replace(/MM/, map.MM)
    .replace(/YYYY/, map.YYYY);
};

const formatCurrency = (value: number | undefined | null, format: string): string => {
  if (typeof value !== "number" || isNaN(value)) return "0.00";

  if (format === "None") return value.toFixed(2);

  const parts = value.toFixed(2).split(".");
  let intPart = parts[0];

  if (format === "India") {
    const lastThree = intPart.slice(-3);
    const rest = intPart.slice(0, -3).replace(/\B(?=(\d{2})+(?!\d))/g, ",");
    intPart = rest ? `${rest},${lastThree}` : lastThree;
  } else if (format === "International") {
    intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  }

  return `${intPart}.${parts[1]}`;
};
// interface POViewerProps {
//   context: any; // Replace 'any' with the specific type of your context if available
// }

class POViewer extends React.Component<
 IPOViewerProps, IPOViewerState
  // {
  //   selectedOpportunity: PO | null;
  //   panelOpen: boolean;
  //   searchQuery: string;
  //   filteredOpportunities: PO[];
  //   purchaseOrderForm: any;
  //   quoteOptions: IDropdownOption[];
  //   opportunityOptions: IDropdownOption[];
  //   lineItems: any[];
  //   purchaseOrders: any[];
  //   showSales: boolean;
  //   currentPage: number;
  //   pageSize: number;
  //   selectedFile: File[];
  //   selectedFileUrl: string;
  //   isViewerOpen: boolean;
  //   filePreviews?: { file: File; previewUrl: string }[];
  //   selectedPreviewFile?: File;
  //   selectedFileName?: string;
  //   allOpportunities: PO[];
  //       isColumnsPanelOpen: boolean;
  //       visibleColumns: IColumn[];
  //       existingFiles: { name: string; url: string }[];


  // }
> {

  private sp: SPFI;
  private _columns: IColumn[];
  constructor(props: IPOViewerProps) {
    super(props);
    this._columns = this.getColumns();
    this.state = {
      selectedOpportunity: null,
      panelOpen: false,
      searchQuery: "",
      purchaseOrderForm: {},
      filteredOpportunities: [],
      quoteOptions: [],
      opportunityOptions: [],
      lineItems: [{ Title: "", Comments: "", Value: 0 }],
      purchaseOrders: [],
      showSales: false,
      currentPage: 1,
      pageSize: 10,
      selectedFile: [],
      selectedFileUrl: "",
      isViewerOpen: false,
      filePreviews: [],
      selectedPreviewFile: undefined,
      selectedFileName: "",
      allOpportunities: [],
          isColumnsPanelOpen: false,
      visibleColumns: this.getColumns(), 
      existingFiles: [],
parentPOOptions: [],
      exchangeRates: {},
      defaultCurrency: "",
      OPPCurrency: "",
    };
    this.sp = spfi().using(SPFx(this.props.context));
  }
  private fetchOpportunitiesOptions = async (): Promise<void> => {
    try {
      const items = await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.select("Id", "Title", "OpportunityID")();

      const options = items.map((item) => ({
        key: item.Title || item.Id,
        text: item.Title || `Opportunity ${item.Id}`,
      }));

      this.setState({ opportunityOptions: options });
    } catch (err) {
      console.error("Failed to fetch opportunities", err);
    }
  };
  private fetchAllPOIDs = async (): Promise<void> => {
  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items
      .select("*")
      .top(5000)();

    const poIdsSet = new Set<string>();

    items.forEach(item => {
      Object.keys(item).forEach(key => {
        if (key.startsWith("POID") && item[key]) {
          poIdsSet.add(item[key]);
        }
      });
    });

    const options: IDropdownOption[] = Array.from(poIdsSet).map(poId => ({
      key: poId,
      text: poId,
    }));

    this.setState({ parentPOOptions: options });
  } catch (error) {
    console.error("Failed to fetch PO IDs", error);
    this.setState({ parentPOOptions: [] });
  }
};
private loadCurrencyConfig = async (): Promise<void> => {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.select("Title", "DefaultCurrency", "MultiValue")();

    const config = configItems.find((item) => item.Title === "Currency");

    if (config?.MultiValue) {
      const parsedRates = JSON.parse(config.MultiValue);

      this.setState({
        defaultCurrency: parsedRates.base,
        exchangeRates: parsedRates.rates,
        purchaseOrderForm: {
          ...this.state.purchaseOrderForm,
          POCurrency: parsedRates.base, // default currency
        },
      });
    } else {
      console.warn("Currency configuration not found or invalid.");
    }
  } catch (err) {
    console.error("Error loading currency configuration:", err);
  }
};
private dateFormat: string = "DD-MMM-YYYY"; // fallback

private async loadDateFormatConfig() {
  try {
    const config = await this.sp.web.lists.getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'DateFormat'")
      .top(1)();

    if (config.length > 0) {
      this.dateFormat = JSON.parse(config[0].MultiValue || "{}").format || this.dateFormat;
    }
  } catch (err) {
    console.warn("Failed to load date format config, using default.");
  }
}

  //   private fetchOpportunities = async () => {
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
  private fetchQuotesForOpportunity = async (
    opportunityID: string
  ): Promise<void> => {
    try {
      const items = await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.filter(`Title eq '${opportunityID}'`)
        .top(1)();

      if (items.length === 0) {
        this.setState({ quoteOptions: [] });
        return;
      }

      const item = items[0];
      const quotes: IDropdownOption[] = [];

      for (let i = 1; i <= 5; i++) {
        const suffix = i === 1 ? "" : i.toString();
        const quoteId = item[`QuoteID${suffix}`];
        if (quoteId) {
          quotes.push({ key: quoteId, text: quoteId });
        }
      }

      this.setState({ quoteOptions: quotes });
    } catch (err) {
      console.error("Failed to fetch quotes for opportunity", err);
      this.setState({ quoteOptions: [] });
    }
  };
  //  private fetchOpportunities = async (): Promise<void> => {
  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.select(
  //         "Id",
  //         "Title",
  //         "OpportunityID",
  //         "Customer",
  //         "OpportunityStatus",
  //         "ReportDate",
  //         "AmountEUR"
  //       )(); // optional if you use 'RecordType'
  //     console.log(items);
  //     const opportunities: PO[] = items.map(
  //       (item: {
  //         OpportunityID: any;
  //         Id: { toString: () => any };
  //         Title: any;
  //         Customer: any;
  //         OpportunityStatus: any;
  //         ReportDate: any;
  //         AmountEUR: any;

  //       }) => ({
  //         lineItems: [],
  //         POID: "", // No PO ID in this context
  //         OpportunityID: item.OpportunityID || item.Id.toString(),
  //         date: item.ReportDate || "",
  //         POtatus: item.OpportunityStatus || "",
  //         Amount: item.AmountEUR ? item.AmountEUR.toString() : "0",
  //         Currency: "EUR",
  //         CustomerPONumber: "",
  //         LineItemsJSON: "[]",
  //         QuoteId: "",
  //       })
  //     );
  //     console.log("Fetched opportunities:", opportunities);
  //     //mockQuotation.push(...opportunities);
  //     this.setState({ filteredOpportunities: opportunities,
  // allOpportunities: opportunities, });
  //   } catch (err) {
  //     console.error("Failed to fetch opportunities", err);
  //   }
  // };
  // private fetchAllQuotations = async (): Promise<void> => {
  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("SalesRecords")
  //       .items.select(
  //         "Id",
  //         "OpportunityID",
  //         "QuoteID",
  //         "QuoteID2",
  //         "QuoteID3",
  //         "QuoteID4",
  //         "QuoteID5",
  //         "QuoteDate",
  //         "QuoteDate2",
  //         "QuoteDate3",
  //         "QuoteDate4",
  //         "QuoteDate5",
  //         "QuoteRevisionNumber",
  //         "QuoteRevisionNumber2",
  //         "QuoteRevisionNumber3",
  //         "QuoteRevisionNumber4",
  //         "QuoteRevisionNumber5",
  //         "QuoteRevenueQuoted",
  //         "QuoteRevenueQuoted2",
  //         "QuoteRevenueQuoted3",
  //         "QuoteRevenueQuoted4",
  //         "QuoteRevenueQuoted5"
  //       )();

  //     const quotations: any[] = [];

  //     items.forEach((item) => {
  //       for (let i = 1; i <= 5; i++) {
  //         const suffix = i === 1 ? "" : i.toString();
  //         const quoteId = item[`QuoteID${suffix}`];
  //         if (quoteId) {
  //           quotations.push({
  //             quoteId,
  //             opportunityId: item.OpportunityID,
  //             revisionNumber: item[`QuoteRevisionNumber${suffix}`],
  //             date: item[`QuoteDate${suffix}`],
  //             revenue: item[`QuoteRevenueQuoted${suffix}`],
  //           });
  //         }
  //       }
  //     });

  //     console.log("Quotations:", quotations);
  //     // Optionally set state
  //     // this.setState({ quotations });
  //     //this.setState({ filteredOpportunities:quotations });
  //   } catch (error) {
  //     console.error("Failed to fetch quotations", error);
  //   }
  // };
  private fetchAllPOs = async (): Promise<void> => {
    try {
      const items = await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.select(
          "*"
        )();

      const pos: any[] = [];

      items.forEach((item) => {
        for (let i = 1; i <= 5; i++) {
          const suffix = i === 1 ? "" : i.toString();
          const POID = item[`POID${suffix}`];
          if (POID) {
            pos.push({
              POID,
              OpportunityID: item.Title,
              POReceivedDate: item[`POReceivedDate${suffix}`],
              POStatus: item[`POStatus${suffix}`],
              POAmount: item[`POAmount${suffix}`],
              Currency: item[`Currency${suffix}`],
              IsChildPO: item[`IsChildPO${suffix}`],
              CustomerPONumber: item[`CustomerPONumber${suffix}`],
              lineItems: item[`LineItemsJSON${suffix}`],
              POQuoteID: item[`POQuoteID${suffix}`],
              POComments: item[`POComments${suffix}`],
              ParentPOID: item[`ParentPOID${suffix}`] || "", 
             POValue: item[`POValue${suffix}`],// Assuming ParentPOID is stored similarly
             // POQuoteID: item[`QuoteID${suffix}`] || "", // Assuming QuoteID is stored similarly
              // Add other fields as needed

            });
          }
        }
      });

      console.log("Purchase Orders:", pos);
      // Optionally set state
      // this.setState({ purchaseOrders: pos });
      this.setState({ filteredOpportunities: pos,allOpportunities: pos, });
    } catch (error) {
      console.error("Failed to fetch purchase orders", error);
    }
  };
  private fetchPurchaseOrders = async (): Promise<void> => {
    const opportunityID = this.state.purchaseOrderForm.OpportunityID;

    if (!opportunityID) {
      alert("Select an Opportunity ID first.");
      return;
    }

    try {
      const items = await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.filter(`OpportunityID eq '${opportunityID}'`)
        .top(1)();

      if (items.length === 0) {
        alert("No Opportunity found.");
        return;
      }

      const item = items[0];
      const purchaseOrders: any[] = [];

      for (let i = 1; i <= 5; i++) {
        const suffix = i === 1 ? "" : i.toString();
        const poId = item[`POID${suffix}`];
        if (poId) {
          purchaseOrders.push({
            POID: poId,
            POReceivedDate: item[`POReceivedDate${suffix}`],
            CustomerPONumber: item[`CustomerPONumber${suffix}`],
            POStatus: item[`POStatus${suffix}`],
            POQuoteID: item[`POQuoteID${suffix}`],
            POAmount: item[`POAmount${suffix}`],
            Currency: item[`POCurrency${suffix}`],
            IsChildPO: item[`IsChildPO${suffix}`],
            LineItemsJSON: item[`LineItemsJSON${suffix}`],
            POComments: item[`POComments${suffix}`],
            QuoteID: item[`QuoteID${suffix}`] || "", // Assuming QuoteID is stored similarly
          });
        }
      }

      this.setState({ purchaseOrders });
    } catch (err) {
      console.error("Error fetching purchase orders", err);
      alert("Failed to fetch purchase orders.");
    }
  };
//   private generateNextPOID = async (): Promise<string> => {
//   const currentYear = new Date().getFullYear();
//   const items = await this.sp.web.lists
//     .getByTitle("CWSalesRecords")
//     .items
//     .select("*")
//     .top(5000)(); // Adjust if list is huge

//   let maxNumber = 0;

//   items.forEach(item => {
//     Object.keys(item).forEach(key => {
//       if (key.startsWith("POID")) {
//         const rawId = item[key];
//         if (typeof rawId === "string" && rawId.startsWith(`PO-${currentYear}`)) {
//           const parts = rawId.split("-");
//           const num = parseInt(parts[2]);
//           if (!isNaN(num) && num > maxNumber) {
//             maxNumber = num;
//           }
//         }
//       }
//     });
//   });

//   const next = maxNumber + 1;
//   return `PO-${currentYear}-${next.toString().padStart(4, "0")}`;
// };

private async generateNextPOID(): Promise<string> {
  const currentYear = new Date().getFullYear();
  let prefix = "PO"; // fallback default

  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'POConfig'")
      .top(1)();

    if (configItems.length > 0) {
      const config = configItems[0];
      const values = JSON.parse(config.MultiValue || "{}");
      prefix = values.prefix || prefix;
    }
  } catch (error) {
    console.warn("PO prefix config not found or invalid, using default.");
  }

  const items = await this.sp.web.lists
    .getByTitle("CWSalesRecords")
    .items.select("*")
    .top(5000)();

  let maxNumber = 0;

  items.forEach(item => {
    Object.keys(item).forEach(key => {
      if (key.startsWith("POID")) {
        const val = item[key];
        if (typeof val === "string" && val.startsWith(`${prefix}-${currentYear}`)) {
          const parts = val.split("-");
          const num = parseInt(parts[2]);
          if (!isNaN(num) && num > maxNumber) {
            maxNumber = num;
          }
        }
      }
    });
  });

  const nextNumber = maxNumber + 1;
  return `${prefix}-${currentYear}-${nextNumber.toString().padStart(4, "0")}`;
}

  private addLineItem = () => {
    this.setState((prev) => ({
      lineItems: [...prev.lineItems, { Title: "", Comments: "", Value: 0 }],
    }));
  };
  private decodeHtmlEntities = (str: string): string => {
  const txt = document.createElement("textarea");
  txt.innerHTML = str;
  return txt.value;
};
  // private handlePOChanges = (field: string, value: any): void => {
  //   this.setState((prevState) => ({
  //     purchaseOrderForm: {
  //       ...prevState.purchaseOrderForm,
  //       [field]: value,
  //     },
  //   }));
  // };

private handlePOChanges = (field: string, value: any): void => {
  const updatedForm = {
    ...this.state.purchaseOrderForm,
    [field]: value,
  };

  if (field === "POValue") {
    const oppID = updatedForm.OpportunityID;
    if (oppID) {
      this.fetchCurrencyForOpportunity(oppID).then((currency) => {
        const baseAmount = parseFloat(value || "0");
        let rate = 1;
        if (currency && this.state.exchangeRates[currency]) {
          rate = this.state.exchangeRates[currency];
          updatedForm.CustomerCurrency = currency;
          updatedForm.POValue = baseAmount;
          updatedForm.POAmount = (baseAmount * rate).toFixed(2);
        } else {
          updatedForm.POAmount = baseAmount.toFixed(2);
        }

        this.setState({ purchaseOrderForm: updatedForm });
      });
      return;
    }
  }

  this.setState({ purchaseOrderForm: updatedForm });
};
private fetchCurrencyForOpportunity = async (opportunityID: string): Promise<string | null> => {
  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items
      .filter(`Title eq '${opportunityID}'`)
      .select("Currency")
      .top(1)();
this.setState({OPPCurrency: items[0]?.Currency || ""});
    return items[0]?.Currency || null;
    
  } catch (error) {
    console.error("Error fetching opportunity currency:", error);
    return null;
  }
};

  private handleLineItemChange = (index: number, field: string, value: any) => {
    const items = [...this.state.lineItems];
    if (field in items[index]) {
      (items[index] as any)[field] = value;
    }
    this.setState({ lineItems: items });
  };

  private renderPurchaseOrderList = () => {
    const { purchaseOrders } = this.state;
    if (!purchaseOrders.length) return null;

    return (
      <div style={{ marginTop: "1rem" }}>
        <Label styles={{ root: { fontSize: "20px", fontWeight: "bold" } }}>
          Existing Purchase Orders
        </Label>

        <Stack tokens={{ childrenGap: 12 }}>
          {purchaseOrders.map((po, index) => (
            <Stack
              key={index}
              tokens={{ childrenGap: 6 }}
              styles={{
                root: {
                  border: "1px solid #E1DFDD",
                  borderRadius: 8,
                  padding: 12,
                  backgroundColor: "#ffffff",
                  boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
                },
              }}
            >
              <span>
                <strong>PO ID:</strong> {po.POID}
              </span>
              <span>
                <strong>Quote ID:</strong> {po.QuoteID}
              </span>
              <span>
                <strong>Status:</strong> {po.POStatus}
              </span>
              <span>
                <strong>Customer PO Number:</strong> {po.CustomerPONumber}
              </span>
              <span>
                <strong>Amount (EUR):</strong> {po.AmountEUR}
              </span>
              <span>
                <strong>Currency:</strong> {po.Currency}
              </span>
              <span>
                <strong>Received Date:</strong> {po.POReceivedDate}
              </span>
            </Stack>
          ))}
        </Stack>
      </div>
    );
  };

   private logAuditEntry = async (
  opportunityID: string,
  action: "Created" | "Updated",
  snapshot: any,
  comments = ""
): Promise<void> => {
  try {
    await this.sp.web.lists.getByTitle("CWSalesAuditLog").items.add({
      OpportunityID: opportunityID,
      Action: action,
      Timestamp: new Date().toISOString(),
      ModifiedBy: this.props.context.pageContext.user.displayName,
      DataSnapshot: JSON.stringify(snapshot),
      Comments: comments,
    });
  } catch (error) {
    console.error("Failed to log audit entry:", error);
  }
}; 

// private submitPurchaseOrder = async (): Promise<void> => {
//   const form = this.state.purchaseOrderForm;
//   const opportunityID = form.OpportunityID;

//   if (!opportunityID) {
//     alert("Please select an Opportunity ID.");
//     return;
//   }

//   try {
//     const items = await this.sp.web.lists
//       .getByTitle("CWSalesRecords")
//       .items.filter(`OpportunityID eq '${opportunityID}'`)
//       .top(1)();

//     if (items.length === 0) {
//       alert("No Opportunity found with this ID.");
//       return;
//     }

//     const item = items[0];
//     const fieldsToUpdate: any = {};
//     let foundSlot = "";
//     let logAction: "Created" | "Updated" = "Created";
//     let finalPoId = form.POID;

//     // Check for update slot first
//     for (let i = 1; i <= 5; i++) {
//       const suffix = i === 1 ? "" : i.toString();
//       const poIdField = `POID${suffix}`;
//       if (item[poIdField] == form.POID) {
//         console.log(item[poIdField], form.POID);
//         foundSlot = suffix;
//         logAction = "Updated";
//         break;
//       }
//     }

//     // If editing existing PO
//     if (logAction == "Updated") {
//       fieldsToUpdate[`POID${foundSlot}`] = form.POID;
//       fieldsToUpdate[`POReceivedDate${foundSlot}`] = form.POReceivedDate;
//       fieldsToUpdate[`CustomerPONumber${foundSlot}`] = form.CustomerPONumber;
//       fieldsToUpdate[`QuoteID${foundSlot}`] = form.QuoteID;
//       fieldsToUpdate[`POStatus${foundSlot}`] = form.POStatus;
//       fieldsToUpdate[`AmountEUR${foundSlot}`] = form.POValue;
//       fieldsToUpdate[`Currency${foundSlot}`] = form.Currency;
//       fieldsToUpdate[`LineItemsJSON${foundSlot}`] = JSON.stringify(this.state.lineItems);
//       fieldsToUpdate[`POComments${foundSlot}`] = form.POComments;
//     } else {
//       // New PO, find empty slot
//       for (let i = 1; i <= 5; i++) {
//         const suffix = i === 1 ? "" : i.toString();
//         if (!item[`POID${suffix}`]) {
//           finalPoId = form.POID || `PO-${Date.now()}`;
//           fieldsToUpdate[`POID${suffix}`] = finalPoId;
//           fieldsToUpdate[`POReceivedDate${suffix}`] = form.POReceivedDate;
//           fieldsToUpdate[`CustomerPONumber${suffix}`] = form.CustomerPONumber;
//           fieldsToUpdate[`QuoteID${suffix}`] = form.QuoteID;
//           fieldsToUpdate[`POStatus${suffix}`] = form.POStatus;
//           fieldsToUpdate[`AmountEUR${suffix}`] = form.POValue;
//           fieldsToUpdate[`Currency${suffix}`] = form.Currency;
//           fieldsToUpdate[`LineItemsJSON${suffix}`] = JSON.stringify(this.state.lineItems);
//           fieldsToUpdate[`POComments${suffix}`] = form.POComments;
//           break;
//         }
//       }
//     }

//     if (Object.keys(fieldsToUpdate).length === 0) {
//       alert("All PO slots are full.");
//       return;
//     }

//     // Perform update
//     await this.sp.web.lists.getByTitle("CWSalesRecords").items.getById(item.Id).update(fieldsToUpdate);

//     // 🔒 Log audit entry
//     await this.logAuditEntry(
//       opportunityID,
//       logAction,
//       {
//         POID: finalPoId,
//         POReceivedDate: form.POReceivedDate,
//         CustomerPONumber: form.CustomerPONumber,
//         QuoteID: form.QuoteID,
//         POStatus: form.POStatus,
//         POValue: form.POValue,
//         Currency: form.Currency,
//         LineItems: this.state.lineItems
//       },
//       "Purchase Order submission"
//     );

//     alert("Purchase Order saved successfully.");

//     // Upload file(s)
//     const files = this.state.selectedFile;
//     if (files.length > 0) {
//       const folderName = `OpportunityID_${opportunityID}`;
//       await this.sp.web.folders.addUsingPath(`Shared Documents/${folderName}`);

//       for (const file of files) {
//         const fileBuffer = await file.arrayBuffer();
//         await this.sp.web
//           .getFolderByServerRelativePath(`Shared Documents/${folderName}`)
//           .files.addUsingPath(file.name, fileBuffer, { Overwrite: true });
//       }

//       alert("Purchase Order and file uploaded successfully.");
//     } else {
//       alert("Purchase Order saved (no file uploaded).");
//     }

//     // Optional: Refresh UI
//     this.fetchAllPOs();
//     this.setState({ showSales: false, selectedOpportunity: null, purchaseOrderForm: {}, selectedFile: [] });

//   } catch (err) {
//     console.error("Error saving Purchase Order:", err);
//     alert("Failed to save Purchase Order.");
//   }
// };
private submitPurchaseOrder = async (): Promise<void> => {
  const form = this.state.purchaseOrderForm;
  const opportunityID = form.OpportunityID;
  const POQuoteID = form.POQuoteID;

  if (!opportunityID) {
    alert("Please select an Opportunity ID.");
    return;
  }

  try {
    const [item] = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items.filter(`Title eq '${opportunityID}'`)
      .top(1)();

    if (!item) {
      alert("No Opportunity found with this ID.");
      return;
    }

    let fieldsToUpdate: any = {};
    // let foundSlot = "";
    let logAction: "Created" | "Updated" = "Created";
    let finalPoId = form.POID || await this.generateNextPOID();


    // Helper to map fields
    const mapFieldsWithSuffix = (suffix: string) => {
      fieldsToUpdate[`POID${suffix}`] = finalPoId;
      fieldsToUpdate[`POReceivedDate${suffix}`] = form.POReceivedDate;
      fieldsToUpdate[`CustomerPONumber${suffix}`] = form.CustomerPONumber;
      fieldsToUpdate[`POQuoteID${suffix}`] = form.POQuoteID;
      fieldsToUpdate[`POStatus${suffix}`] = form.POStatus;
      fieldsToUpdate[`POAmount${suffix}`] = form.POAmount;
      fieldsToUpdate[`POValue${suffix}`] = form.POValue;
      // fieldsToUpdate[`POCurrency${suffix}`] = form.POCurrency;
      fieldsToUpdate[`IsChildPO${suffix}`] = form.IsChildPO;
      fieldsToUpdate[`ParentPOID${suffix}`] = form.ParentPOID || ""; // Assuming ParentPOID is stored similarly
      fieldsToUpdate[`LineItemsJSON${suffix}`] = JSON.stringify(this.state.lineItems);
      fieldsToUpdate[`POComments${suffix}`] = form.POComments;
     // fieldsToUpdate[`QuoteID${suffix}`] = form.QuoteID || ""; // Assuming QuoteID is stored similarly
    };

    // Check if it's an update
    for (let i = 1; i <= 5; i++) {
      const suffix = i === 1 ? "" : i.toString();
      if (item[`POID${suffix}`] === form.POID) {
        // foundSlot = suffix;
        logAction = "Updated";
        mapFieldsWithSuffix(suffix);
        break;
      }
    }

    // If new PO, find empty slot
    if (logAction === "Created") {
      for (let i = 1; i <= 5; i++) {
        const suffix = i === 1 ? "" : i.toString();
        if (!item[`POID${suffix}`]) {
          // foundSlot = suffix;
          mapFieldsWithSuffix(suffix);
          break;
        }
      }
    }

// if (typeof foundSlot !== "string") {
//   alert("All PO slots are full.");
//   return;
// }

    // Perform SharePoint update
    await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items.getById(item.Id)
      .update(fieldsToUpdate);

    // Log audit
    await this.logAuditEntry(
      opportunityID,
      logAction,
      {
        POID: finalPoId,
        POReceivedDate: form.POReceivedDate,
        CustomerPONumber: form.CustomerPONumber,
        POQuoteID: form.POQuoteID,
        POStatus: form.POStatus,
        POValue: form.POValue,
        Currency: form.POCurrency,
        LineItems: this.state.lineItems,
        POComments: form.POComments,
        ParentPOID: form.ParentPOID || "", // Assuming ParentPOID is stored similarly


      },
      "Purchase Order submission"
    );

    // Upload attachments if present
    const files = this.state.selectedFile;
    let fileUploadMessage = "";

    if (files.length > 0) {
      const folderName = `${form.POID || finalPoId}`;
      const folderPath = `Shared Documents/${opportunityID}/${POQuoteID}/${folderName}`;

           try{
      await this.sp.web.folders.addUsingPath(`Shared Documents/${opportunityID}/${POQuoteID}/${folderName}`);
    } catch (error) {
      if (!error.message.includes("already exists")) {
        console.error("Error creating folder:", error);
        throw error;
      }
      console.log("Folder already exists, skipping creation.");
    }

      for (const file of files) {
        const fileBuffer = await file.arrayBuffer();
        await this.sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.addUsingPath(file.name, fileBuffer, { Overwrite: true });
      }

      fileUploadMessage = " and file(s) uploaded";
    }

    alert(`Purchase Order ${logAction.toLowerCase()} successfully${fileUploadMessage}.`);

    // Refresh state/UI
    this.fetchAllPOs();
    this.setState({
      showSales: false,
      selectedOpportunity: null,
      purchaseOrderForm: {},
      selectedFile: [],
    });

  } catch (err) {
    console.error("Error saving Purchase Order:", err);
    alert("Failed to save Purchase Order. Check console for details.");
  }
};

private loadOpportunityFiles = async (oppId:string,quoteId: string,POId:string) => {
  const folderName = `${POId}`;

  try {
    const files = await this.sp.web
      .getFolderByServerRelativePath(`Shared Documents/${oppId}/${quoteId}/${folderName}`)
      .files.select("Name", "ServerRelativeUrl")();
console.log("Files in folder:", files);
    const existingFiles = files.map((f: any) => ({
      name: f.Name,
      url: f.ServerRelativeUrl,
    }));

    this.setState({ existingFiles });
  } catch (err) {
    console.warn("No existing files found for this opportunity.", err);
    this.setState({ existingFiles: [] });
  }
};
private deleteFileFromLibrary = async (fileName: string) => {
  if (!this.state.selectedOpportunity) return;
const ParentFolder = `${this.state.selectedOpportunity.OpportunityID}`;
const ParentQuoteFolder = `${this.state.selectedOpportunity.POQuoteID}`;
  const folderName = `${this.state.selectedOpportunity.POID}`;
  const confirm = window.confirm(`Delete file "${fileName}" from SharePoint?`);

  if (!confirm) return;

  try {
    const files = await this.sp.web
      .getFolderByServerRelativePath(`Shared Documents/${ParentFolder}/${ParentQuoteFolder}/${folderName}`)
      .files.filter(`Name eq '${fileName}'`)();

    if (files.length > 0) {
      await this.sp.web.getFileByServerRelativePath(files[0].ServerRelativeUrl).recycle(); // OR .delete() for permanent
      alert("File deleted.");
      this.loadOpportunityFiles(this.state.selectedOpportunity.OpportunityID,this.state.selectedOpportunity.POQuoteID,this.state.selectedOpportunity.POID);
    } else {
      alert("File not found.");
    }
  } catch (err) {
    console.error("Error deleting file:", err);
    alert("Failed to delete the file.");
  }
};

private deletePO = async (poIdToDelete: string): Promise<void> => {
  const confirmDelete = window.confirm(`Are you sure you want to delete PO "${poIdToDelete}"?`);
  if (!confirmDelete) return;

  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items.filter(`POID eq '${poIdToDelete}' or POID2 eq '${poIdToDelete}' or POID3 eq '${poIdToDelete}' or POID4 eq '${poIdToDelete}' or POID5 eq '${poIdToDelete}'`)
      .top(1)();

    if (items.length === 0) {
      alert("Purchase Order not found.");
      return;
    }

    const item = items[0];
    const itemId = item.Id;
    const fieldsToUpdate: any = {};

    // Clear matching PO slot
    for (let i = 1; i <= 5; i++) {
      const suffix = i === 1 ? "" : i.toString();
      if (item[`POID${suffix}`] === poIdToDelete) {
        fieldsToUpdate[`POID${suffix}`] = null;
        fieldsToUpdate[`POReceivedDate${suffix}`] = null;
        fieldsToUpdate[`POStatus${suffix}`] = null;
        fieldsToUpdate[`CustomerPONumber${suffix}`] = null;
        fieldsToUpdate[`POAmount${suffix}`] = null;
        fieldsToUpdate[`POCurrency${suffix}`] = null;
        fieldsToUpdate[`POQuoteID${suffix}`] = null;
        fieldsToUpdate[`LineItemsJSON${suffix}`] = null;
        fieldsToUpdate[`POComments${suffix}`] = null;
        //fieldsToUpdate[`QuoteID${suffix}`] = null; // Assuming QuoteID is stored similarly
        break;
      }
    }

    await this.sp.web.lists.getByTitle("CWSalesRecords").items.getById(itemId).update(fieldsToUpdate);

    alert("Purchase Order deleted successfully.");
    await this.fetchAllPOs();
  } catch (error) {
    console.error("Failed to delete PO:", error);
    alert("Failed to delete PO.");
  }
};
 private openColumnsPanel = () => {
    this.setState({ isColumnsPanelOpen: true });
  };

  private closeColumnsPanel = () => {
    this.setState({ isColumnsPanelOpen: false });
  };

  private updateColumns = (selectedColumns: string[]) => {
    const visibleColumns = this.getConfiguredColumns(selectedColumns);
    this.setState({ visibleColumns: visibleColumns.filter((col): col is IColumn => col !== undefined) });
  };

  private getConfiguredColumns = (selectedColumns: any[]) => {
    const columns = selectedColumns.map((keyObj) => {
      // Check if any key in selectedColumns matches the key of the current column and return
      const column = this._columns.find((col) => col.key === keyObj.key);
      return column;
    });
    console.log(columns);
    return columns;
  };
private currencySeparator: string = "";

private async loadCurrencySeparatorFormat() {
  const configItems = await this.sp.web.lists
    .getByTitle("CWSalesConfiguration")
    .items.filter("Title eq 'CurrencyFormat'")
    .top(1)();

  if (configItems.length > 0) {
    this.currencySeparator = JSON.parse(configItems[0].MultiValue || "{}").format || "International";
  }
}

componentDidMount() {
    // this.fetchOpportunities();
    // this.fetchAllQuotations();
    this.fetchAllPOs();
    this.fetchOpportunitiesOptions();
    this.fetchAllPOIDs();
     this.loadCurrencyConfig();
this.loadDateFormatConfig();
this.loadCurrencySeparatorFormat();

     const savedColumnConfig = localStorage.getItem('editPOItem');
    
         // If there is a saved configuration, update the state
         if (savedColumnConfig) {
          const parsedConfig = JSON.parse(savedColumnConfig);
          const visibleColumns = this.getConfiguredColumns(parsedConfig);
          this.setState({ visibleColumns: visibleColumns.filter((col): col is IColumn => col !== undefined) });
        } else {
          // Provide a default configuration if no saved configuration is found
          this.setState({ visibleColumns: this._columns });
        }
        console.log("Visible Columns")
        console.log(this.state.visibleColumns)
  }
  componentWillUnmount() {
  if (this.state.filePreviews) {
    this.state.filePreviews.forEach(p => URL.revokeObjectURL(p.previewUrl));
  }
}
  onRowClick = (item?: PO, _index?: number, _ev?: Event) => {
    if (item) {
      this.setState({ selectedOpportunity: item, panelOpen: true });
    }
  };

onSearchChange = (_ev: any, newValue?: string) => {
  const searchQuery = newValue || "";
  const lowerValue = searchQuery.toLowerCase();
  const filtered = this.state.allOpportunities.filter(
    (opp) =>
      opp.POID.includes(searchQuery) ||
      opp.OpportunityID.toLowerCase().includes(lowerValue) ||
      opp.POStatus.toLowerCase().includes(lowerValue) ||
      opp.CustomerPONumber.includes(searchQuery)
  );
  this.setState({ searchQuery, filteredOpportunities: filtered });
};
  getPagedItems = (): PO[] => {
    const { filteredOpportunities, currentPage, pageSize } = this.state;
    const startIndex = (currentPage - 1) * pageSize;
    return filteredOpportunities.slice(startIndex, startIndex + pageSize);
  };
  nextPage = () => {
    const maxPage = Math.ceil(
      this.state.filteredOpportunities.length / this.state.pageSize
    );
    if (this.state.currentPage < maxPage) {
      this.setState((prevState) => ({
        currentPage: prevState.currentPage + 1,
      }));
    }
  };

  prevPage = () => {
    if (this.state.currentPage > 1) {
      this.setState((prevState) => ({
        currentPage: prevState.currentPage - 1,
      }));
    }
  };
  exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(
      this.state.filteredOpportunities
    );
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "PurchaseOrder");
    XLSX.writeFile(workbook, "PurchaseOrder.xlsx");
  };
removeLineItem = (index: number) => {
  this.setState((prevState) => ({
    lineItems: prevState.lineItems.filter((_, i) => i !== index)
  }));
};
getColumns=():IColumn[] => [
  {
    key: "POID",
    name: "PO Number",
    fieldName: "POID",
    minWidth: 100,
    isResizable: true,
  },
  // { key: "date", name: "PO Received Date", fieldName: "date", minWidth: 100 },
  { key: "POStatus", name: "PO Status", fieldName: "POStatus", minWidth: 90 , isResizable: true},
  { key: "POAmount", name: "Amount", fieldName: "POAmount", minWidth: 100 , isResizable: true,onRender: (item) => formatCurrency(item.POAmount || 0, this.currencySeparator)},
  { key: "OpportunityID", name: "Opportunity", fieldName: "OpportunityID", minWidth: 100 , isResizable: true},
  { key: "POReceivedDate", name: "POReceivedDate", fieldName: "POReceivedDate", minWidth: 100 , isResizable: true, onRender: (item: PO) => {
  if (!item.POReceivedDate) return "";
  return formatDate(item.POReceivedDate, this.dateFormat);
}
,},
    { key: "POQuoteID", name: "POQuote ID", fieldName: "POQuoteID", minWidth: 100 , isResizable: true},
  {
          key: "actions",
          name: "",
          fieldName: "actions",
          minWidth: 40,
          maxWidth: 40,
          isResizable: false,
          onRender: (item: PO) => {
            return (
              <DefaultButton
                menuIconProps={{ iconName: "MoreVertical" }}
                  styles={{
    root: {
      padding: 0,
      minWidth: 28,
      height: 28,
      border: "none",
      backgroundColor: "transparent",
    },
    rootHovered: {
      backgroundColor: "#f3f2f1", // subtle hover
    },
    rootFocused: {
      border: "none",
      outline: "none",
    },
    rootPressed: {
      backgroundColor: "#edebe9", // subtle press
    },
    icon: {
      color: "#605e5c", // icon color
    },
  }}
                menuProps={{
                  items: [
                    {
                      key: "edit",
                      text: "Edit",
                      iconProps: { iconName: "Edit" },
     onClick: () => {
      this.loadOpportunityFiles(item.OpportunityID,item.POQuoteID,item.POID);
console.log("Edit clicked for item:", item);
let parsedLineItems = [{ Title: "", Comments: "", Value: 0 }];
  try {
    const raw = item.lineItems || "[]"; // Ensure it's a string
    const decoded = this.decodeHtmlEntities(raw); // Decode HTML entities
    parsedLineItems = JSON.parse(decoded); // Parse into object
    console.log("Parsed line items:", parsedLineItems);
    console.log("Parsed line items:", decoded);
  } catch (e) {
    console.warn("Failed to decode lineItems JSON:", e);
  }
  this.setState({
    selectedOpportunity: item,
    purchaseOrderForm: { ...item },
    lineItems: parsedLineItems,
    showSales: true,
  });
},
                    },

                    {
              key: "delete",
              text: "Delete",
              iconProps: { iconName: "Delete" },
              onClick: () => this.deletePO(item.POID),
            },
                  ],
                }}
              />
            );
          },
        },

];
  render() {
    const {
      selectedOpportunity,
      panelOpen,
      searchQuery,
      // filteredOpportunities,
    } = this.state;


    const commandBarItems = [
      {
        key: "new",
        text: "New PO",
        iconProps: { iconName: "Add" },
  onClick: async () => {
  const newPOID = await this.generateNextPOID();
  await this.fetchAllPOIDs(); // fetch PO IDs before opening panel
  this.setState({
    showSales: true,
    purchaseOrderForm: { POID: newPOID },
  });
},

        styles: { primaryButtonStyles },
      },
      // {
      //   key: "edit",
      //   text: "Edit PO",
      //   iconProps: { iconName: "Edit" },
      //   disabled: !selectedOpportunity,
      //   onClick: () => {
      //     this.setState({
      //       showSales: true,
      //       purchaseOrderForm: selectedOpportunity,
      //     });
      //   },
      //   styles: { primaryButtonStyles },
      // },
      {
        key: "export",
        text: "Export to Excel",
        iconProps: { iconName: "ExcelDocument" },
        onClick: this.exportToExcel,
        styles: { primaryButtonStyles },
      },
    {
  key:"columns",
  text: "Edit Columns",
  iconProps: { iconName: "ColumnOptions" },
  onClick: this.openColumnsPanel,
  styles: { primaryButtonStyles },
}  
    ];
    return (
      <div style={{ width: "100%", height: "100vh" }} id="sales-webpart-root">
        <MessageBar messageBarType={MessageBarType.info}>
          Welcome to the PO Viewer
        </MessageBar>
        <TextField
          placeholder="Search Purchase Order..."
          value={searchQuery}
          onChange={this.onSearchChange}
          styles={{ root: { marginBottom: 10 } }}
        />
        <CommandBar items={commandBarItems} />
             <div style={{ overflowX: 'scroll', overflowY: 'scroll', height: "calc(100vh - 300px)", width: '100%'}}>
        <div style={{ minWidth: `${this.state.visibleColumns.length * 70}px`,}}>
        <DetailsList
          items={this.getPagedItems()}
          columns={this.state.visibleColumns}
          setKey="set"
          // onItemInvoked={(item, index, ev) =>
          //   this.onRowClick(item as PO, index, ev)
          // }
          selectionMode={SelectionMode.none}
          //onItemInvoked={this.onRowClick}
        />
        </div>
        </div>
        <Stack
          horizontal
          tokens={{ childrenGap: 10 }}
          style={{ marginTop: 10 }}
        >
          <PrimaryButton
            text="Previous"
            onClick={this.prevPage}
            disabled={this.state.currentPage === 1}
            styles={primaryButtonStyles}
          />
          <PrimaryButton
            text="Next"
            onClick={this.nextPage}
            disabled={
              this.state.currentPage >=
              Math.ceil(
                this.state.filteredOpportunities.length / this.state.pageSize
              )
            }
            styles={primaryButtonStyles}
          />
        </Stack>
        <div>
          {this.state.isColumnsPanelOpen && (
            <ColumnsViewPanel
              visibleColumns={this.state.visibleColumns}
              allColumns={this._columns}
              onClose={this.closeColumnsPanel}
              onUpdateColumns={this.updateColumns}
              localStorageKey="editPOItem"
            />
          )}
        </div>
        <Panel
          isOpen={panelOpen}
          onDismiss={() => this.setState({ panelOpen: false })}
          headerText={selectedOpportunity?.POID}
          type={PanelType.medium}
        >
          {selectedOpportunity && (
            <div>
              <p>
                <strong>PO Number:</strong> {selectedOpportunity.POID}
              </p>
              <p>
                <strong>PO Received Date:</strong>{" "}
                {selectedOpportunity.POReceivedDate}
              </p>
              <p>
                <strong>PO Status:</strong> {selectedOpportunity.POStatus}
              </p>
              <p>
                <strong>Revenue:</strong> {selectedOpportunity.POAmount}
              </p>
            </div>
          )}
        </Panel>
        <Panel
          isOpen={this.state.showSales}
          onDismiss={() =>
            this.setState({
              showSales: false,
              selectedOpportunity: null,
              purchaseOrderForm: {},
              lineItems: [],
              selectedFile: [],
              existingFiles: [],
              filePreviews: [],
            })
          }
          headerText={
            !selectedOpportunity
              ? "Create Purchase Order"
              : "Edit Purchase Order"
          }
          type={PanelType.large}
          styles={{
            main: {
              backgroundColor: "#f5f5f5",
            },
          }}
        >
          <Stack
            horizontal
            tokens={{ childrenGap: 24 }}
            styles={{ root: { width: "100%" } }}
          >
            <Stack
              tokens={{ childrenGap: 12 }}
              styles={{ root: { width: "50%" } }}
            >
              <div style={{ padding: 16 }}>
                <div style={gridStyle}>
                  <Stack
                    horizontal
                    tokens={{ childrenGap: 12 }}
                    styles={{ root: { width: "100%" } }}
                  >
                    <Dropdown
                      label="Opportunity ID"
                      placeholder="Select an Opportunity ID"
                      // disabled={this.state.selectedOpportunity == null}
                      selectedKey={
                        this.state.purchaseOrderForm.OpportunityID || ""
                      }
                      options={this.state.opportunityOptions}
                      onChange={(_, option) => {
                        this.handlePOChanges("OpportunityID", option?.key);
                        if (option?.key)
                          void this.fetchQuotesForOpportunity(
                            option.key.toString()
                          );
                      }}
                      styles={{ root: { width: "50%" } }}
                    />
                    <Dropdown
                      label="Quote ID"
                      placeholder="Select a Quote ID"
                      // disabled
                      selectedKey={this.state.purchaseOrderForm.POQuoteID || ""}
                      options={this.state.quoteOptions}
                      onChange={(_, option) =>
                        this.handlePOChanges("POQuoteID", option?.text)
                      }
                      styles={{ root: { width: "50%" } }}
                    />
                  </Stack>
                  <TextField
                    label="PO ID"
                    disabled
                    value={this.state.purchaseOrderForm.POID || ""}
                  />

                  <Stack
                    horizontal
                    tokens={{ childrenGap: 12 }}
                    styles={{ root: { width: "100%" } }}
                  >
                    {/* <Dropdown
                label="Is Child PO"
                defaultSelectedKey={
                  this.state.purchaseOrderForm.IsChildPO ||
                this.state.selectedOpportunity?.IsChildPO
                }
                options={[
                  { key: "true", text: "Yes" },
                  { key: "false", text: "No" },
                ]}
                onChange={(_, v) => this.handlePOChanges("IsChildPO", v?.key)}
              styles={{ root: { width: "50%" } }}
              /> */}
                    <Dropdown
                      label="Is Child PO"
                      selectedKey={
                        this.state.purchaseOrderForm.IsChildPO === true
                          ? "true"
                          : this.state.purchaseOrderForm.IsChildPO === false
                          ? "false"
                          : undefined
                      }
                      options={[
                        { key: "true", text: "Yes" },
                        { key: "false", text: "No" },
                      ]}
                      onChange={(_, option) => {
                        const value = option?.key === "true";
                        this.handlePOChanges("IsChildPO", value); // store as boolean
                      }}
                      styles={{ root: { width: "50%" } }}
                    />

                    <Dropdown
  label="Parent PO ID"
  placeholder="Select a Parent PO"
  options={this.state.parentPOOptions}
  disabled={this.state.purchaseOrderForm.IsChildPO === false}
  selectedKey={this.state.purchaseOrderForm.ParentPOID || undefined}
  onChange={(_, option) =>
    this.handlePOChanges("ParentPOID", option?.key)
  }
  styles={{ root: { width: "50%" } }}
/>

                  </Stack>
                  <DatePicker
                    label="PO Received Date"
                    value={
                      this.state.purchaseOrderForm.POReceivedDate
                        ? new Date(this.state.purchaseOrderForm.POReceivedDate)
                        : this.state.selectedOpportunity?.POReceivedDate
                        ? new Date(
                            this.state.selectedOpportunity.POReceivedDate
                          )
                        : undefined
                    }
                    onSelectDate={(d) =>
                      this.handlePOChanges("POReceivedDate", d)
                    }
                  />
                  <Dropdown
                    label="PO Status"
                    defaultSelectedKey={
                      this.state.purchaseOrderForm.POStatus ||
                      this.state.selectedOpportunity?.POStatus
                    }
                    options={[
                      { key: "Draft", text: "Draft" },
                      { key: "Issued", text: "Issued" },
                      { key: "Approved", text: "Approved" },
                      { key: "Cancelled", text: "Cancelled" },
                    ]}
                    onChange={(_, v) =>
                      this.handlePOChanges("POStatus", v?.text)
                    }
                  />
                  <TextField
                    label="Customer PO Number"
                    defaultValue={
                      this.state.purchaseOrderForm.CustomerPONumber ||
                      this.state.selectedOpportunity?.CustomerPONumber
                    }
                    onChange={(_, v) =>
                      this.handlePOChanges("CustomerPONumber", v)
                    }
                  />
                  <TextField
                    label={`PO Value${this.state?.defaultCurrency ? ` (${this.state.defaultCurrency})` : ""}`}
                    type="number"
                    defaultValue={this.state.purchaseOrderForm.POValue || ""}
                    onChange={(_, v) => this.handlePOChanges("POValue", v)}
                  />
                  <TextField
                    label={`PO Amount${this.state?.OPPCurrency ? ` (${this.state.OPPCurrency})` : ""}`}
                    type="number"
                    disabled
                    value={this.state.purchaseOrderForm.POAmount || ""}
                    onChange={(_, v) => this.handlePOChanges("POAmount", v)}
                  />
                  <TextField
                    label="Comments"
                    multiline
                    rows={3}
                    value={this.state.purchaseOrderForm.POComments || ""}
                    // value={selectedOpportunity?.comments || ""}
                    onChange={(_, v) => this.handlePOChanges("POComments", v)}
                  />
                </div>
                {/* Line Items */}
                <div style={sectionHeaderStyle}>Line Items *</div>
                {this.state.lineItems.map((item: any, index: number) => (
                  <div key={index} style={lineItemStyle}>
                    <TextField
                      placeholder="Title"
                      value={item.Title}
                      onChange={(_, v) =>
                        this.handleLineItemChange(index, "Title", v)
                      }
                    />
                    <TextField
                      placeholder="Comment"
                      value={item.Comments}
                      onChange={(_, v) =>
                        this.handleLineItemChange(index, "Comments", v)
                      }
                    />
                    <TextField
                      placeholder="Value"
                      type="number"
                      value={item.Value.toString()}
                      onChange={(_, v) =>
                        this.handleLineItemChange(
                          index,
                          "Value",
                          parseFloat(v || "0")
                        )
                      }
                    />
                    <IconButton
                      iconProps={{ iconName: "Cancel" }}
                      title="Remove"
                      ariaLabel="Remove"
                      onClick={() => this.removeLineItem(index)}
                    />
                  </div>
                ))}
                <div style={{ marginTop: 16 }}>
                  <PrimaryButton
                    text="+ Add Line Item"
                    onClick={this.addLineItem}
                    styles={primaryButtonStyles}
                  />
                </div>
                {/* Action Buttons */}
                {!this.state.purchaseOrderForm.POID && (
                  <div style={{ marginTop: 24 }}>
                    <Label style={{ fontWeight: 600 }}>
                      Upload PO Attachment
                    </Label>
                    {/* <input
    type="file"
    onChange={(e) => {
      const file = e.target.files?.[0];
      if (file) {
        this.setState({ selectedFile: file });
      }
    }}
  /> */}
                    <Stack tokens={{ childrenGap: 8 }}>
                      {/* <Label>Attachment</Label>
                <input type="file" onChange={this.handleQuotationFileChange} /> */}
                      <DropZoneUploader
                        onFilesSelected={(files) => {
                          // const previews = files.map(file => ({
                          //   file,
                          //   previewUrl: URL.createObjectURL(file)
                          // }));
                          this.setState((prevState) => {
                            const newFiles = [
                              ...prevState.selectedFile,
                              ...files,
                            ];
                            const newPreviews = newFiles.map((file) => ({
                              file,
                              previewUrl: URL.createObjectURL(file),
                            }));
                            this.state.filePreviews?.forEach((p) =>
                              URL.revokeObjectURL(p.previewUrl)
                            );
                            this.setState({
                              selectedFile: newFiles,
                              filePreviews: newPreviews,
                              isViewerOpen: false,
                            });
                            // return {
                            //   selectedFile: newFiles,
                            //   filePreviews: newPreviews,
                            //   isViewerOpen: false
                            // };
                          });
                        }}
                      />

                      {this.state.isViewerOpen && (
                        <div>
                          <DocumentViewer
                            url={this.state.selectedFileUrl || ""}
                            isOpen={this.state.isViewerOpen}
                            onDismiss={() =>
                              this.setState({ isViewerOpen: false })
                            }
                            fileName={this.state.selectedFileName || ""}
                          />
                        </div>
                      )}
                    </Stack>
                  </div>
                )}
                {this.state.selectedFile.map((file, index) => (
                  <div
                    key={index}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      border: "1px solid #ccc",
                      borderRadius: 6,
                      padding: "8px 12px",
                      marginTop: 12,
                      backgroundColor: "#fff",
                      maxWidth: 300,
                    }}
                  >
                    <Icon
                      iconName="Page"
                      style={{ fontSize: 16, marginRight: 8 }}
                    />
                    <span
                      style={{
                        flexGrow: 1,
                        cursor: "pointer",
                        color: "#0078d4",
                        textDecoration: "underline",
                      }}
                      onClick={() =>
                        this.setState({
                          isViewerOpen: true,
                          selectedFileUrl: URL.createObjectURL(file),
                          selectedPreviewFile: file,
                          selectedFileName: file.name,
                        })
                      }
                    >
                      {file.name}
                    </span>
                    <Icon
                      iconName="Delete"
                      style={{
                        fontSize: 16,
                        cursor: "pointer",
                        color: "#a80000",
                        marginLeft: 8,
                      }}
                      onClick={() => {
                        const newFiles = [...this.state.selectedFile];
                        newFiles.splice(index, 1);
                        this.setState({ selectedFile: newFiles });
                      }}
                    />
                  </div>
                ))}
                {this.state.purchaseOrderForm.POID && (
                  <Stack tokens={{ childrenGap: 6 }}>
                    {this.state.existingFiles.length > 0 && (
                      <div>
                        <label>
                          <b>Existing Attachments:</b>
                        </label>
                        <ul>
                          {this.state.existingFiles.length > 0 &&
                            this.state.existingFiles.map((file, index) => (
                              <div
                                key={index}
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  border: "1px solid #ccc",
                                  borderRadius: 6,
                                  padding: "8px 12px",
                                  marginTop: 12,
                                  backgroundColor: "#fff",
                                  maxWidth: 300,
                                }}
                              >
                                <Icon
                                  iconName="Page"
                                  style={{ fontSize: 16, marginRight: 8 }}
                                />

                                <a
                                  href={file.url}
                                  target="_blank"
                                  rel="noopener noreferrer"
                                  style={{
                                    flexGrow: 1,
                                    color: "#0078d4",
                                    textDecoration: "underline",
                                    overflow: "hidden",
                                    textOverflow: "ellipsis",
                                    whiteSpace: "nowrap",
                                  }}
                                >
                                  {file.name}
                                </a>
                                <Icon
                                  iconName="Delete"
                                  title="Delete"
                                  ariaLabel="Delete"
                                  style={{
                                    fontSize: 16,
                                    cursor: "pointer",
                                    color: "#a80000",
                                    marginLeft: 8,
                                  }}
                                  onClick={() =>
                                    this.deleteFileFromLibrary(file.name)
                                  }
                                />
                              </div>
                            ))}
                        </ul>
                      </div>
                    )}

                    {/* DropZoneUploader for new files */}
                    <DropZoneUploader
                      //selectedFiles={this.state.selectedFile}
                      onFilesSelected={(files) =>
                        this.setState({ selectedFile: files })
                      }
                    />
                  </Stack>
                )}
                <Stack
                  horizontal
                  tokens={{ childrenGap: 12 }}
                  style={{ marginTop: 24 }}
                >
                  <PrimaryButton
                    text="View POs"
                    onClick={this.fetchPurchaseOrders}
                    styles={primaryButtonStyles}
                  />
                  <PrimaryButton
                    text="Submit"
                    onClick={this.submitPurchaseOrder}
                    styles={primaryButtonStyles}
                  />
                </Stack>

                {/* Optional List Render */}
                <div style={{ marginTop: 24 }}>
                  {this.renderPurchaseOrderList()}
                </div>
              </div>
            </Stack>
            <Stack
              tokens={{ childrenGap: 12 }}
              styles={{ root: { width: "50%" } }}
            >
              {/* Example: Show the first file if available */}
              {this.state.existingFiles.length > 0 ? (
<DocumentViewer
    url={this.state.existingFiles[0].url}
    isOpen={true}
    fileName={this.state.existingFiles[0].name}
    onDismiss={() => this.setState({ isViewerOpen: false })}
  />
              ) : (
                <DocumentViewer
                  url={this.state.selectedFileUrl || ""}
                  isOpen={this.state.isViewerOpen}
                  onDismiss={() => this.setState({ isViewerOpen: false })}
                  fileName={this.state.selectedFileName || ""}
                />
              )}
            </Stack>
          </Stack>
        </Panel>
      </div>
    );
  }
}

export default POViewer;
