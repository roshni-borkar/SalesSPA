/* eslint-disable no-unused-expressions */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable react/self-closing-comp */
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
// import { MSGraphClientV3 } from "@microsoft/sp-http";
// import Sales from '../Sales';
import {
  DatePicker,
  DefaultButton,
  Dropdown,
  getTheme,
  IButtonStyles,
  Icon,
  IDropdownOption,
  // Label,
  // Label,
  PrimaryButton,
  Stack,
} from "@fluentui/react";
import DropZoneUploader from "../DropZoneUploader";
import DocumentViewer from "../DocumentViewer";
import ColumnsViewPanel from "../ColumnsViewSettings/ColumnsViewPanel";
// import { get } from "@microsoft/sp-lodash-subset";
import { IQuotationViewerProps } from "./IViewQuotaionProps";
import { IQuotationViewerState, Quotation } from "./IViewQuotationState";
// Define Opportunity type
// type Quotation = {
//   QuoteID: string;
//   OpportunityID: string;
//   QuoteDate: string;
//   QuoteRevisionNumber: string;
//   QuoteRevenueQuoted: string;
//   QuoteBusinessSize:string;
//   QuoteAmount: number;
//   QuoteCurrency: string;
//   // RevisionNumber:string;
//   QuoteTentativeDecisionDate:Date;
//   QuoteComments: string;
// };

// const mockQuotation: Quotation[] = [];
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
// const currencies: IDropdownOption[] = [
//   { key: "EUR", text: "EUR" },
//   { key: "USD", text: "USD" },
//   { key: "GBP", text: "GBP" },
// ];
// const columns: IColumn[] = [
//   {
//     key: "quoteId",
//     name: "Quotation Number",
//     fieldName: "quoteId",
//     minWidth: 100,
//     isResizable: true,
//   },
//   { key: "date", name: "Quotation Date", fieldName: "date", minWidth: 100 },
//   {
//     key: "revisionNumber",
//     name: "Revision Number",
//     fieldName: "revisionNumber",
//     minWidth: 90,
//   },
//   { key: "revenue", name: "Revenue", fieldName: "revenue", minWidth: 100 },
// ];
// const columnGridStyle = {
//   root: {
//     display: "grid",
//     gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
//     gap: "16px",
//   },
// };
const businessSizes: IDropdownOption[] = [
  { key: "Small", text: "Small" },
  { key: "Medium", text: "Medium" },
  { key: "Large", text: "Large" },
];
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


// interface QuotationViewerProps {
//   context: any; // Replace 'any' with the specific type of your context if available
//   salesProps: any;
// }

class QuotationViewer extends React.Component<
  IQuotationViewerProps,IQuotationViewerState
  // {
  //   selectedOpportunity: Quotation | null;
  //   panelOpen: boolean;
  //   searchQuery: string;
  //   filteredOpportunities: Quotation[];
  //   showSales: boolean;
  //   quotationForm: any;
  //   opportunityOptions: IDropdownOption[];
  //   quotationFile: any;
  //   currentPage: number;
  //   pageSize: number;
  //   filePreviewUrl: string | null;
  //   selectedFile?: File[];
  //   selectedFileUrl?: string | null;
  //   isViewerOpen?: boolean;
  //   filePreviews?: { file: File; previewUrl: string }[];
  //   selectedPreviewFile?: File | null;
  //   selectedFileName?: string | null;
  //       isColumnsPanelOpen: boolean;
  //       visibleColumns: IColumn[];
  //     existingFiles: { name: string; url: string }[];
  //     allQuotations: Quotation[];

  // }
  
> {
  private sp: SPFI;
    private _columns: IColumn[];
  constructor(props: IQuotationViewerProps) {
    super(props);
    this._columns = this.getColumns();
    this.state = {
      selectedOpportunity: null,
      panelOpen: false,
      searchQuery: "",
      filteredOpportunities: [],
      showSales: false,
      quotationForm: {},
      quotationFile: null,
      opportunityOptions: [], // Initialize as an empty array
      currentPage: 1,
      pageSize: 10,
      filePreviewUrl: null,
      selectedFile: [],
      selectedFileUrl: null,
      isViewerOpen: false,
      filePreviews: [],
      selectedPreviewFile: null,
      selectedFileName: null,
      isColumnsPanelOpen: false,
      visibleColumns: this.getColumns(),
      existingFiles: [],
      allQuotations: [], // Initialize as an empty array,
      exchangeRates: {}, // Initialize as an empty object,
      QuoteAmount: "",         // EUR,
      defaultCurrency: "",      // e.g., INR
      RevenueInCustomerCurrency: "", // converted amount
      OPPCurrency: "", // e.g., USD
    };
    this.sp = spfi().using(SPFx(this.props.context));
  }
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

  // private fetchOpportunities = async (): Promise<void> => {
  //   try {
  //     const items = await this.sp.web.lists
  //       .getByTitle("CWSalesRecords")
  //       .items.select(
  //         "Id",
  //         "Title",
  //         "OpportunityID",
  //         "Customer",
  //         "OpportunityStatus",
  //         "ReportDate",
  //         "Amount"
  //       )(); // optional if you use 'RecordType'
  //     console.log(items);
  //     const opportunities = items.map(
  //       (item: {
  //         OpportunityID: any;
  //         Id: { toString: () => any };
  //         Title: any;
  //         Customer: any;
  //         OpportunityStatus: any;
  //         ReportDate: any;
  //         Amount: any;
  //       }) => ({
  //         key: item.OpportunityID || item.Id.toString(),
  //         title: item.Title || `Opportunity ${item.Id}`,
  //         account: item.Customer,
  //         stage: item.OpportunityStatus,
  //         owner: "Unknown", // optionally extend with owner logic
  //         closeDate: item.ReportDate,
  //         probability: 50, // static or custom logic
  //         revenue: item.Amount ? `€${item.Amount}` : "€0",
  //       })
  //     );
  //     console.log("Fetched opportunities:", opportunities);
  //     //mockQuotation.push(...opportunities);
  //     // this.setState({ filteredOpportunities: opportunities });
  //   } catch (err) {
  //     console.error("Failed to fetch opportunities", err);
  //   }
  // };
  private fetchAllQuotations = async (): Promise<void> => {
    try {
      const items = await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.select(
         "*"

        )();

      const quotations: any[] = [];

      items.forEach((item) => {
        for (let i = 1; i <= 5; i++) {
          const suffix = i === 1 ? "" : i.toString();
          const QuoteId = item[`QuoteID${suffix}`];
          if (QuoteId) {
            quotations.push({
              QuoteID: QuoteId,
              Title: item.Title,
              QuoteRevisionNumber: item[`QuoteRevisionNumber${suffix}`],
              QuoteRevenueQuoted: item[`QuoteRevenueQuoted${suffix}`],
              QuoteAmount: item[`QuoteAmount${suffix}`],
              QuoteBusinessSize: item[`QuoteBusinessSize${suffix}`],
              QuoteDate: item[`QuoteDate${suffix}`],
              QuoteTentativeDecisionDate: item[`QuoteTentativeDecisionDate${suffix}`],
              // QuoteAmount: item[`QuoteAmount${suffix}`],
              QuoteCurrency: item[`Currency${suffix}`],
              QuoteComments: item[`QuoteComments${suffix}`],

            });
          }
        }
      });

      console.log("Quotations:", quotations);
      // Optionally set state
      // this.setState({ quotations });
      this.setState({ filteredOpportunities: quotations,allQuotations: quotations });
    } catch (error) {
      console.error("Failed to fetch quotations", error);
    }
  };
// private generateNextQuoteID = async (): Promise<string> => {
//   const currentYear = new Date().getFullYear();

//   const items = await this.sp.web.lists
//     .getByTitle("CWSalesRecords")
//     .items
//     .select("*")
//     .top(5000)(); // Adjust as needed for large lists

//   let maxNumber = 0;

//   items.forEach(item => {
//     Object.keys(item).forEach(key => {
//       if (key.startsWith("QuoteID")) {
//         const val = item[key];
//         if (typeof val === "string" && val.startsWith(`QTN-${currentYear}`)) {
//           const parts = val.split("-");
//           const num = parseInt(parts[2]);
//           if (!isNaN(num) && num > maxNumber) {
//             maxNumber = num;
//           }
//         }
//       }
//     });
//   });

//   const nextNumber = maxNumber + 1;
//   return `QTN-${currentYear}-${nextNumber.toString().padStart(4, "0")}`;
// };

private async generateNextQuoteID(): Promise<string> {
  const currentYear = new Date().getFullYear();
  let prefix = "QTN"; // fallback

  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'QTNConfig'")
      .top(1)();

    if (configItems.length > 0) {
      const config = JSON.parse(configItems[0].MultiValue || "{}");
      prefix = config.prefix || "QTN";
    }
  } catch (err) {
    console.warn("QTN prefix config not found or invalid, using default.");
  }

  const items = await this.sp.web.lists
    .getByTitle("CWSalesRecords")
    .items.select("*")
    .top(5000)();

  let maxNumber = 0;

  items.forEach(item => {
    Object.keys(item).forEach(key => {
      const val = item[key];
      if (key.startsWith("QuoteID") && typeof val === "string" && val.startsWith(`${prefix}-${currentYear}`)) {
        const parts = val.split("-");
        const num = parseInt(parts[2]);
        if (!isNaN(num) && num > maxNumber) {
          maxNumber = num;
        }
      }
    });
  });

  const nextNumber = maxNumber + 1;
  return `${prefix}-${currentYear}-${nextNumber.toString().padStart(4, "0")}`;
}


  private fetchOpportunitiesOption = async (): Promise<void> => {
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
          const poId = item[`POID${suffix}`];
          if (poId) {
            pos.push({
              poId,
              OpportunityId: item.OpportunityID,
              date: item[`POReceivedDate${suffix}`],
              status: item[`POStatus${suffix}`],
              amount: item[`Amount${suffix}`],
              currency: item[`Currency${suffix}`],
            });
          }
        }
      });

      console.log("Purchase Orders:", pos);
      // Optionally set state
      // this.setState({ purchaseOrders: pos });
    } catch (error) {
      console.error("Failed to fetch purchase orders", error);
    }
  };
  // private handleQuotationChange = (field: string, value: any): void => {
  //   this.setState((prevState) => ({
  //     quotationForm: {
  //       ...prevState.quotationForm,
  //       [field]: value,
  //     },
  //   }));
  // };
private handleQuotationChange = async (field: string, value: any): Promise<void> => {
  const updatedForm = {
    ...this.state.quotationForm,
    [field]: value,
  };

  if (field === "QuoteAmount") {
    const oppID = updatedForm.Title;
    if (oppID) {
      const currency = await this.fetchCurrencyForOpportunity(oppID);
      const baseAmount = parseFloat(value || "0");
      let rate = 1;
      if (currency) {
        rate = (this.state.exchangeRates && this.state.exchangeRates[currency]) || 1;
       // updatedForm.CustomerCurrency = currency;
        updatedForm.QuoteRevenueQuoted = (baseAmount * rate).toFixed(2);
        this.setState({
         // defaultCurrency: currency,
        })
      } else {
        //updatedForm.CustomerCurrency = "";
        updatedForm.QuoteRevenueQuoted = baseAmount.toFixed(2);
      }
    }
  }

  this.setState({ quotationForm: updatedForm });
};
private loadCurrencyConfig = async (): Promise<void> => {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.select("Title", "DefaultCurrency", "MultiValue")();

    const config = configItems.find((item) => item.Title === "Currency");

    if (config?.MultiValue) {
      const parsedRates = JSON.parse(config.MultiValue);

      // const currencyOptions: IDropdownOption[] = Object.keys(parsedRates.rates).map((code) => ({
      //   key: code,
      //   text: code,
      // }));

      this.setState({
        defaultCurrency: parsedRates.base,
        exchangeRates: parsedRates.rates,
        //currencyOptions,
        quotationForm: {
          ...this.state.quotationForm,
          QuoteCurrency: parsedRates.base, // set default currency for new form
        },
      });
    } else {
      console.warn("Currency configuration not found or invalid.");
    }
  } catch (err) {
    console.error("Error loading currency configuration:", err);
  }
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

private submitQuotation = async (): Promise<void> => {
  const { selectedFile } = this.state;
  const form = this.state.quotationForm;
  const {
    Title,
    QuoteID,
    QuoteDate,
    QuoteRevisionNumber,
    QuoteRevenueQuoted,
    QuoteBusinessSize,
    QuoteTentativeDecisionDate,
    QuoteAmount,
    QuoteCurrency,
    QuoteComments
  } = form;

  if (!Title) {
    alert("Please select an Opportunity ID.");
    return;
  }

  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items.filter(`Title eq '${Title}'`)
      .top(1)();

    if (items.length === 0) {
      alert("No Opportunity found with this ID.");
      return;
    }

    const item = items[0];
    const itemId = item.Id;
    const fieldsToUpdate: any = {};
    let isUpdate = false;
    let logAction: "Created" | "Updated" = "Created";

    for (let i = 1; i <= 5; i++) {
      const suffix = i === 1 ? "" : i.toString();
      const existingQuoteID = item[`QuoteID${suffix}`];

      if (existingQuoteID === QuoteID) {
        fieldsToUpdate[`QuoteDate${suffix}`] = QuoteDate;
        fieldsToUpdate[`QuoteRevisionNumber${suffix}`] = QuoteRevisionNumber;
        fieldsToUpdate[`QuoteRevenueQuoted${suffix}`] = QuoteRevenueQuoted;
        fieldsToUpdate[`QuoteBusinessSize${suffix}`] = QuoteBusinessSize;
        fieldsToUpdate[`QuoteTentativeDecisionDate${suffix}`] = QuoteTentativeDecisionDate;
        fieldsToUpdate[`QuoteAmount${suffix}`] = QuoteAmount;
        // fieldsToUpdate[`QuoteCurrency${suffix}`] = QuoteCurrency;
        fieldsToUpdate[`QuoteComments${suffix}`] = QuoteComments;
        isUpdate = true;
        logAction = "Updated";
        break;
      }
    }

    if (!isUpdate) {
      for (let i = 1; i <= 5; i++) {
        const suffix = i === 1 ? "" : i.toString();
        if (!item[`QuoteID${suffix}`]) {
         const newQuoteID = QuoteID || await this.generateNextQuoteID();
          fieldsToUpdate[`QuoteID${suffix}`] = newQuoteID;
          fieldsToUpdate[`QuoteDate${suffix}`] = QuoteDate;
          fieldsToUpdate[`QuoteRevisionNumber${suffix}`] = QuoteRevisionNumber;
          fieldsToUpdate[`QuoteRevenueQuoted${suffix}`] = QuoteRevenueQuoted;
          fieldsToUpdate[`QuoteBusinessSize${suffix}`] = QuoteBusinessSize;
          fieldsToUpdate[`QuoteTentativeDecisionDate${suffix}`] = QuoteTentativeDecisionDate;
          fieldsToUpdate[`QuoteAmount${suffix}`] = QuoteAmount;
          // fieldsToUpdate[`QuoteCurrency${suffix}`] = QuoteCurrency;
          fieldsToUpdate[`QuoteComments${suffix}`] = QuoteComments;
          isUpdate = true;
          form.QuoteID = newQuoteID; // ensure we store it
          break;
        }
      }
    }

    if (!isUpdate) {
      alert("No available or matching slot to update/add quotation.");
      return;
    }

    // Perform update
    await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items.getById(itemId)
      .update(fieldsToUpdate);

    // 🔒 Log audit entry
    await this.logAuditEntry(
      Title,
      logAction,
      {
        QuoteID: form.QuoteID,
        QuoteDate,
        QuoteRevisionNumber,
        QuoteRevenueQuoted,
        QuoteBusinessSize,
        QuoteTentativeDecisionDate,
        QuoteAmount,
        QuoteCurrency
      },
      "Quotation submission"
    );

    alert("Quotation saved successfully.");

    // File upload (optional)
    if ((selectedFile ?? []).length > 0) {
      const folderName = `${form.QuoteID}`;
     // await this.sp.web.folders.addUsingPath(`Shared Documents/${folderName}`);
        try{
      await this.sp.web.folders.addUsingPath(`Shared Documents/${Title}/${folderName}`);
    } catch (error) {
      if (!error.message.includes("already exists")) {
        console.error("Error creating folder:", error);
        throw error;
      }
      console.log("Folder already exists, skipping creation.");
    }
    for (const file of selectedFile ?? []) {
      const fileBuffer = await file.arrayBuffer();
      await this.sp.web
        .getFolderByServerRelativePath(`Shared Documents/${Title}/${folderName}`)
        .files.addUsingPath(file.name, fileBuffer, { Overwrite: true });

      alert("File uploaded to document library.");
    }
  }
    // Optionally refresh quotations
    await this.fetchAllQuotations();
    this.setState({ showSales: false, quotationForm: {}, selectedOpportunity: null });

  } catch (err) {
    console.error("Error saving quotation", err);
    alert("Failed to save quotation.");
  }
};

private deleteQuotation = async (quoteIdToDelete: string): Promise<void> => {
  const confirmDelete = window.confirm(`Are you sure you want to delete quotation "${quoteIdToDelete}"?`);
  if (!confirmDelete) return;

  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items.filter(`QuoteID eq '${quoteIdToDelete}' or QuoteID2 eq '${quoteIdToDelete}' or QuoteID3 eq '${quoteIdToDelete}' or QuoteID4 eq '${quoteIdToDelete}' or QuoteID5 eq '${quoteIdToDelete}'`)
      .top(1)();

    if (items.length === 0) {
      alert("Quotation not found.");
      return;
    }

    const item = items[0];
    const itemId = item.Id;
    const fieldsToUpdate: any = {};

    // Remove the quotation from one of the 5 slots
    for (let i = 1; i <= 5; i++) {
      const suffix = i === 1 ? "" : i.toString();
      if (item[`QuoteID${suffix}`] === quoteIdToDelete) {
        fieldsToUpdate[`QuoteID${suffix}`] = null;
        fieldsToUpdate[`QuoteDate${suffix}`] = null;
        fieldsToUpdate[`QuoteRevisionNumber${suffix}`] = null;
        fieldsToUpdate[`QuoteRevenueQuoted${suffix}`] = null;
        fieldsToUpdate[`QuoteBusinessSize${suffix}`] = null;
        fieldsToUpdate[`QuoteTentativeDecisionDate${suffix}`] = null;
         fieldsToUpdate[`QuoteAmount${suffix}`] = null;
        // fieldsToUpdate[`QuoteCurrency${suffix}`] = null;
        fieldsToUpdate[`QuoteComments${suffix}`] = null;
        break;
      }
    }

    await this.sp.web.lists.getByTitle("CWSalesRecords").items.getById(itemId).update(fieldsToUpdate);

    alert("Quotation deleted successfully.");
    await this.fetchAllQuotations();
  } catch (error) {
    console.error("Failed to delete quotation:", error);
    alert("Failed to delete quotation.");
  }
};

private loadOpportunityFiles = async (quoteId: string,ParentFolder:string) => {
  const folderName = `${quoteId}`;
const ParentFolderName = `${ParentFolder}`;
  try {
    const files = await this.sp.web
      .getFolderByServerRelativePath(`Shared Documents/${ParentFolderName}/${folderName}`)
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
const ParentFolderName = `${this.state.quotationForm?.Title}`;
  const folderName = `${this.state.selectedOpportunity.QuoteID}`;
  const confirm = window.confirm(`Delete file "${fileName}" from SharePoint?`);

  if (!confirm) return;

  try {
    const files = await this.sp.web
      .getFolderByServerRelativePath(`Shared Documents/${ParentFolderName}/${folderName}`)
      .files.filter(`Name eq '${fileName}'`)();

    if (files.length > 0) {
      await this.sp.web.getFileByServerRelativePath(files[0].ServerRelativeUrl).recycle(); // OR .delete() for permanent
      alert("File deleted.");
      this.loadOpportunityFiles(folderName,ParentFolderName);
    } else {
      alert("File not found.");
    }
  } catch (err) {
    console.error("Error deleting file:", err);
    alert("Failed to delete the file.");
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

  private fetchCurrencyForOpportunity = async (opportunityID: string): Promise<string | null> => {
  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesRecords")
      .items
      .filter(`Title eq '${opportunityID}'`)
      .select("Currency")
      .top(1)();
this.setState({ OPPCurrency: items[0]?.Currency || "" });
    return items[0]?.Currency || null;
  } catch (error) {
    console.error("Error fetching opportunity currency:", error);
    return null;
  }
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
    this.fetchAllQuotations();
    this.fetchAllPOs();
    void this.fetchOpportunitiesOption();
    void this.loadCurrencyConfig();
    this.loadDateFormatConfig();
    this.loadCurrencySeparatorFormat();
    // this.getGraphFileWebUrl("Q1745316883257");
    const savedColumnConfig = localStorage.getItem('editQuoteItem');
    
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
  onRowClick = async (item?: Quotation, _index?: number, _ev?: Event) => {
    if (item) {
      this.setState({ selectedOpportunity: item, panelOpen: true });
      // const fileUrl = await this.getGraphFileWebUrl("Q1745316883257");
      this.setState({
        selectedOpportunity: item,
        panelOpen: true,
        // filePreviewUrl: fileUrl,
      });
    }
  };

  onSearchChange = (_ev: any, newValue?: string) => {
    const searchQuery = newValue || "";
    const lowerValue = searchQuery.toLowerCase();
    console.log("Search Query:", this.state.allQuotations);
const filtered = this.state.allQuotations.filter(
      (opp) =>
        opp.Title.toLowerCase().includes(lowerValue) ||
         opp.QuoteComments.toLowerCase().includes(lowerValue)||
        opp.QuoteID.includes(searchQuery) 
        // opp.QuoteRevisionNumber.toLowerCase().includes(lowerValue)
    );
    this.setState({ searchQuery, filteredOpportunities: filtered });
  };
  getPagedItems = (): Quotation[] => {
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
    XLSX.utils.book_append_sheet(workbook, worksheet, "Quotation");
    XLSX.writeFile(workbook, "Quotation.xlsx");
  };
 getColumns = ():IColumn[] => [
  {
    key: "QuoteID",
    name: "Quotation Number",
    fieldName: "QuoteID",
    minWidth: 100,
    isResizable: true,
  },
  { key: "QuoteDate", name: "Quotation Date", fieldName: "QuoteDate", minWidth: 100, onRender: (item: Quotation) => {
  if (!item.QuoteDate) return "";
  return formatDate(item.QuoteDate, this.dateFormat);
}, },
  {
    key: "QuoteRevisionNumber",
    name: "Revision Number",
    fieldName: "QuoteRevisionNumber",
    minWidth: 90,
    isResizable: true,
  },
  { key: "QuoteAmount", name: "Amount", fieldName: "QuoteAmount", minWidth: 100 ,onRender: (item) => formatCurrency(item.QuoteAmount || 0, this.currencySeparator),isResizable: true,

},
  { key: "Title", name: "Opportunity", fieldName: "Title", minWidth: 100 ,isResizable: true,},
  {
        key: "actions",
        name: "",
        fieldName: "actions",
        minWidth: 40,
        maxWidth: 40,
        isResizable: false,
        onRender: (item: Quotation) => {
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
                      this.loadOpportunityFiles(item.QuoteID,item.Title);
                      console.log("Edit clicked for item:", item);
                      this.setState({
                        selectedOpportunity: item,
                        quotationForm: { ...item },
                        showSales: true,
                      });
                    },
                  },
                  {
            key: "delete",
            text: "Delete",
            iconProps: { iconName: "Delete" },
            onClick: () => this.deleteQuotation(item.QuoteID)
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
        text: "New Quotation",
        iconProps: { iconName: "Add" },
        onClick: async () => {
  const newQuoteID = await this.generateNextQuoteID(); // assuming this is defined
  this.setState({
    showSales: true,
    quotationForm: { QuoteID: newQuoteID },
  });
},

        styles: { primaryButtonStyles },
      },

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

    //  if (this.state.showSales) {
    //       return (
    //         <Sales
    //           description={this.props.salesProps.description}
    //           isDarkTheme={false}
    //           environmentMessage={this.props.salesProps.environmentMessage}
    //           hasTeamsContext={false}
    //           userDisplayName={this.props.salesProps.userDisplayName}
    //           sp={this.props.salesProps.sp}
    //           context={this.props.salesProps.context}
    //           View="Quotation"
    //         />
    //       );
    //     }

    return (
       <div style={{ width: "100%", height: "100vh" }} id="sales-webpart-root">
        <MessageBar messageBarType={MessageBarType.info}>
          Welcome to the Quotation Viewer
        </MessageBar>
        <TextField
          placeholder="Search Quotation..."
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
          //   this.onRowClick(item as Quotation, index, ev)
          // }
          selectionMode={SelectionMode.single}
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
            localStorageKey='editQuoteItem'
          />
        )}
        </div>

        <Panel
          isOpen={panelOpen}
          onDismiss={() => this.setState({ panelOpen: false,quotationForm: {}, selectedOpportunity: null, selectedFile: [],existingFiles: [] })}
          headerText={selectedOpportunity?.QuoteID}
          type={PanelType.medium}
        >
          {selectedOpportunity && (
            <div>
              <p>
                <strong>Quotation Number:</strong> {selectedOpportunity.QuoteID}
              </p>
              <p>
                <strong>Quotation Date:</strong> {selectedOpportunity.QuoteDate}
              </p>
              <p>
                <strong>Revision Number:</strong>{" "}
                {selectedOpportunity.QuoteRevisionNumber}
              </p>
              <p>
                <strong>Revenue:</strong>{" "}
                {selectedOpportunity.QuoteRevenueQuoted}
              </p>
              {this.state.filePreviewUrl && (
                <iframe
                  src={this.state.filePreviewUrl}
                  width="100%"
                  height="600px"
                  frameBorder="0"
                  title="Document Preview"
                />
              )}
            </div>
          )}
        </Panel>
        <Panel
          isOpen={this.state.showSales}
          onDismiss={() => this.setState({ showSales: false,selectedOpportunity: null, quotationForm: {}, selectedFile: [], existingFiles: [], filePreviewUrl: null, })}
          headerText={!selectedOpportunity?.QuoteID?"Create Quotation":"Edit Quotation"}
          type={PanelType.extraLarge}
          styles={{
            main: {
              backgroundColor: "#f5f5f5",
            },
          }}
        >
          <Stack horizontal tokens={{ childrenGap: 24 }} styles={{ root: { width: "100%" } }}>
                    <Stack
                      tokens={{ childrenGap: 12 }}
                      styles={{ root: { width: "50%" } }}
                    >
          <div style={{ padding: 16 }}>
            {/* <div style={columnGridStyle.root}> */}
              <TextField
                label="Quote ID"
                disabled
                value={this.state.quotationForm.QuoteID}
                onChange={(_, v) => this.handleQuotationChange("QuoteID", v)}
              />
              <Dropdown
                label="Opportunity ID"
                options={this.state.opportunityOptions}
                defaultSelectedKey={
                  this.state.quotationForm.Title || ""}
                onChange={(_, option) =>
                  this.handleQuotationChange("Title", option?.key)
                }
              />
              <Dropdown
                label="Business Size"
                options={businessSizes}
                defaultSelectedKey={
                  this.state.quotationForm?.QuoteBusinessSize || ""
                }
                onChange={(_, option) =>
                  this.handleQuotationChange("QuoteBusinessSize", option?.key)
                }
              />
                            <Stack
                              horizontal
                              tokens={{ childrenGap: 12 }}
                              styles={{ root: { width: "100%" } }}
                            >
              <DatePicker
                label="Quote Date"
                 value={this.state.quotationForm?.QuoteDate ? new Date(this.state.quotationForm?.QuoteDate) : undefined}
                onSelectDate={(date) =>
                  this.handleQuotationChange("QuoteDate", date)
                }
                 styles={{ root: { width: "50%" } }}
              />
              <DatePicker
                label="Tentative Decision Date"
            value={
                  this.state.quotationForm?.QuoteTentativeDecisionDate? new Date(this.state.quotationForm?.QuoteTentativeDecisionDate) : undefined}
                onSelectDate={(date) =>
                  this.handleQuotationChange("QuoteTentativeDecisionDate", date)
                }
                styles={{ root: { width: "50%" } }}
              />
              </Stack>
              <TextField
                label="Revision Number"
                value={this.state.quotationForm?.QuoteRevisionNumber}
                onChange={(_, v) =>
                  this.handleQuotationChange("QuoteRevisionNumber", v)
                }
                styles={{ root: { width: "50%" } }}
              />
                                         <Stack
                              horizontal
                              tokens={{ childrenGap: 12 }}
                              styles={{ root: { width: "100%" } }}
                            >
              <TextField
                label={`Amount${this.state?.defaultCurrency ? ` (${this.state.defaultCurrency})` : ""}`}
                type="number"
                value={this.state.quotationForm?.QuoteAmount}
                onChange={(_, v) => this.handleQuotationChange("QuoteAmount", v)}
                styles={{ root: { width: "50%" } }}
              />
              {/* <Dropdown
                label="Currency"
                options={currencies}
                defaultSelectedKey={this.state.quotationForm?.QuoteCurrency || ""}
                onChange={(_, option) =>
                  this.handleQuotationChange("QuoteCurrency", option?.key)
                }
              /> */}
              <TextField
                label={`Revenue Quoted${this.state?.OPPCurrency ? ` (${this.state.OPPCurrency})` : ""}`}
                disabled
                value={this.state.quotationForm?.QuoteRevenueQuoted}
                onChange={(_, v) =>
                  this.handleQuotationChange("QuoteRevenueQuoted", v)
                }
                styles={{ root: { width: "50%" } }}
              />
              </Stack>
              <TextField
                label="Comments"
                multiline
                rows={3}
                value={this.state.quotationForm.QuoteComments || ""}
                // value={selectedOpportunity?.comments || ""}
                onChange={(_, v) => this.handleQuotationChange("QuoteComments", v)}
          
             />
                {/* 
                <input type="file" onChange={this.handleQuotationFileChange} /> */}
                {/* <Label>Attachment</Label> */}
                                  {this.state.quotationForm.QuoteID && <Stack tokens={{ childrenGap: 6 }}>
  {this.state.existingFiles.length > 0 && (
    <div>
      <label><b>Existing Attachments:</b></label>
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
        onClick={() => this.deleteFileFromLibrary(file.name)}
      />
    </div>
  ))}


      </ul>
    </div>
  )}

  {/* DropZoneUploader for new files */}
  <DropZoneUploader
    //selectedFiles={this.state.selectedFile}
    onFilesSelected={(files) => this.setState({ selectedFile: files })}
  />
</Stack>
}
                 <Stack tokens={{ childrenGap: 8 }}>
                {!this.state.quotationForm.QuoteID&&<DropZoneUploader
                 onFilesSelected={(files) => {
    // const previews = files.map(file => ({
    //   file,
    //   previewUrl: URL.createObjectURL(file)
    // }));
       this.setState((prevState) => {
  const newFiles = [...(prevState.selectedFile || []), ...files];
  const newPreviews = newFiles.map(file => ({
    file,
    previewUrl: URL.createObjectURL(file)
  }));
  return {
    selectedFile: newFiles,
    filePreviews: newPreviews,
    isViewerOpen: false
  };
});
  }}
                />}
{this.state.selectedFile && this.state.selectedFile.map((file, index) => (
  <div key={index} style={{ display: "flex", alignItems: "center", border: "1px solid #ccc", borderRadius: 6, padding: "8px 12px", marginTop: 12, backgroundColor: "#fff", maxWidth: 300 }}>
    <Icon iconName="Page" style={{ fontSize: 16, marginRight: 8 }} />
    <span style={{ flexGrow: 1, cursor: "pointer", color: "#0078d4", textDecoration: "underline" }}
      onClick={() => this.setState({ isViewerOpen: true, selectedFileUrl: URL.createObjectURL(file) , selectedPreviewFile: file,selectedFileName: file.name })}>
      {file.name}
    </span>
    <Icon iconName="Delete" style={{ fontSize: 16, cursor: "pointer", color: "#a80000", marginLeft: 8 }}
      onClick={() => {
        const newFiles = [...(this.state.selectedFile || [])];
        newFiles.splice(index, 1);
        this.setState({ selectedFile: newFiles });
      }} />
  </div>
))}
                {/* {this.state.isViewerOpen && (
                  <div>
                    <DocumentViewer
                      url={this.state.selectedFileUrl || ""}
                      isOpen={this.state.isViewerOpen}
                      onDismiss={() => this.setState({ isViewerOpen: false })}
 fileName={this.state.selectedFileName || ""}
                    />
                  </div>
                )} */}
              </Stack>
            {/* </div> */}

            {/* Submit Button */}
            <Stack horizontalAlign="start" styles={{ root: { marginTop: 20 } }}>
              <PrimaryButton
                text="Submit"
                onClick={this.submitQuotation}
                styles={primaryButtonStyles}
              />
            </Stack>
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
                                ):<DocumentViewer
                        url={this.state.selectedFileUrl || ""}
                        isOpen={!!this.state.isViewerOpen}
                        onDismiss={() => this.setState({ isViewerOpen: false })}
                        fileName={this.state.selectedFileName || ""}
                      />}
                              
                    </Stack>
                    </Stack>
        </Panel>
      </div>
    );
  }
}

export default QuotationViewer;
