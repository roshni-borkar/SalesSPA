/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-unused-vars */
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
  IOppViewerState,
  Opportunity,
  ExchangeRates,
} from "./IViewOpportunitiesState";
// import Sales from '../Sales';
import {
  ComboBox,
  DatePicker,
  DefaultButton,
  Dropdown,
  getTheme,
  IButtonStyles,
  // IComboBoxOption,
  // Icon,
  IconButton,
  IDropdownOption,
  Label,
  PrimaryButton,
  Stack,
} from "@fluentui/react";
// import AdaptiveCardRenderer from "../AdaptiveCard/AdaptiveCardRenderer";
// import { opportunityCard } from "../AdaptiveCard/opportunityCardTemplate";
// import DropZoneUploader from "../DropZoneUploader";
// import DocumentViewer from "../DocumentViewer";
import ColumnsViewPanel from "../ColumnsViewSettings/ColumnsViewPanel";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import { Quotation } from "../ViewQuotation/IViewQuotationState";
// import styles from "../Sales.module.scss";
// import * as $ from "jquery";
// import { WebPartContext } from '@microsoft/sp-webpart-base';
// import { ISalesProps } from '../ISalesProps';
// import Sales from '../Sales';
// import { WebPartContext } from '@microsoft/sp-webpart-base';

// Define Opportunity type
// type Opportunity = {
//   Id: number;
//   OpportunityID: string;
//   Title: string;
//   // Account: string;
//   // Stage: string;
//   // Owner: string;
//   // CloseDate: string;
//   // Probability: number;
//   // Revenue: string;
//   Business: string;
//   BusinessUnit: string;
//   OEM: string;
//   EndCustomer: string;
//   KeyContact: string;
//   DecisionMaker: string;
//   TentativeStartDate: string;
//   TentativeDecisionDate: string;
//   Currency: string;
//   RiskLevel: string;
//   Strategic: string;
//   OpportunityStatus: string;
//   OppComments: string;
//   OppAmount: number;
// };

const mockOpportunities: Opportunity[] = [];
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
// const gridStyle = {
//   display: "grid",
//   gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))",
//   gap: "16px",
// };

const riskLevels: IDropdownOption[] = [
  { key: "Low", text: "Low" },
  { key: "Medium", text: "Medium" },
  { key: "High", text: "High" },
];

const strategicOptions: IDropdownOption[] = [
  { key: "Yes", text: "Yes" },
  { key: "No", text: "No" },
];
// const currencies: IDropdownOption[] = [
//   { key: "EUR", text: "EUR" },
//   { key: "USD", text: "USD" },
//   { key: "GBP", text: "GBP" },
// ];

const statusOptions: IDropdownOption[] = [
  { key: "Open", text: "Open" },
  { key: "Won", text: "Won" },
  { key: "Lost", text: "Lost" },
  { key: "On Hold", text: "On Hold" },
];
interface OpportunityViewerProps {
  context: any; // Replace 'any' with the specific type of your context if available
  salesProps: any;
}
// type ExchangeRates = {
//   [currencyCode: string]: number;
// };

const convertToDefaultCurrency = (
  amount: number,
  fromCurrency: string,
  defaultCurrency: string,
  exchangeRates: ExchangeRates
): number => {
  if (fromCurrency === defaultCurrency) return amount;

  const rateFrom = exchangeRates[fromCurrency];
  const rateTo = exchangeRates[defaultCurrency];

  if (!rateFrom || !rateTo) {
    console.warn("Invalid exchange rate configuration");
    return amount;
  }

  const convertedAmount = (amount / rateFrom) * rateTo;
  return parseFloat(convertedAmount.toFixed(2));
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

class OpportunityViewer extends React.Component<
  OpportunityViewerProps,
  IOppViewerState
  // {
  //   selectedOpportunity: Opportunity | null;
  //   panelOpen: boolean;
  //   searchQuery: string;
  //   filteredOpportunities: Opportunity[];
  //   showSales: boolean;
  //   opportunityForm: any;
  //   currentPage: number;
  //   pageSize: number;
  //   selectedFile: File[];
  //   showForm: boolean;
  //   isViewerOpen: boolean;
  //   selectedFileUrl: string | null;
  //   filePreviews?: { file: File; previewUrl: string }[];
  //   selectedPreviewFile: File | null;
  //   selectedFileName: string | null;
  //   isColumnsPanelOpen: boolean;
  //   visibleColumns: IColumn[];
  //   isCustomerPanelOpen: boolean;
  //   customerForm: {
  //     Customer: string;
  //     PersonResponsible: string | null;
  //     City: string;
  //   };
  //   customerOptions: IDropdownOption[]; // your dropdown options
  //   // allOpportunities: Opportunity[];
  //   isKeyContactPanelOpen: boolean;
  //   keyContactForm: {
  //     Customer: string;
  //     Contact: string;
  //     Email: string;
  //     Address: string;
  //     City: string;
  //     BusinessPhone: string;
  //     MobileNumber: string;
  //     Designation: string;
  //     Department: string;
  //   };
  //   keyContactOptions: IDropdownOption[];
  //   existingFiles: { name: string; url: string }[];
  //   defaultCurrency: string; // fallback currency
  //   exchangeRates: ExchangeRates; // exchange rates configuration
  //   currencyOptions: IComboBoxOption[];
  //   allOpportunities: Opportunity[];
  // }
> {
  private sp: SPFI;
  private _columns: IColumn[];
  _peoplePickerContext: {
    absoluteUrl: any;
    msGraphClientFactory: any;
    spHttpClient: any;
  };
  constructor(props: OpportunityViewerProps) {
    super(props);
    this._columns = this.getColumns();
    this._peoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };
    this.state = {
      selectedOpportunity: null,
      panelOpen: false,
      searchQuery: "",
      filteredOpportunities: [],
      showSales: false,
      opportunityForm: {},
      currentPage: 1,
      pageSize: 10,
      selectedFile: [],
      showForm: false,
      isViewerOpen: false,
      selectedFileUrl: null,
      filePreviews: [],
      selectedPreviewFile: null,
      selectedFileName: null,
      isColumnsPanelOpen: false,
      visibleColumns: this.getColumns(),
      isCustomerPanelOpen: false,
      customerForm: {
        Customer: "",
        PersonResponsible: null,
        City: "",
      },
      customerOptions: [],
      isKeyContactPanelOpen: false,
      keyContactForm: {
        Customer: "",
        Contact: "",
        Email: "",
        Address: "",
        City: "",
        BusinessPhone: "",
        MobileNumber: "",
        Designation: "",
        Department: "",
      },
      keyContactOptions: [],
      //filePreviews: [], // for local uploads with preview
      existingFiles: [], // for already uploaded SharePoint files
      defaultCurrency: "EUR", // fallback
      exchangeRates: {
        EUR: 1.0,
        USD: 1.07,
        GBP: 0.85,
      },
      currencyOptions: [],
      allOpportunities: [],
    };
    this.sp = spfi().using(SPFx(this.props.context));
  }

  private fetchOpportunities = async (): Promise<void> => {
    try {
      const items = await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.select("*")(); // optional if you use 'RecordType'
      console.log(items);
      const opportunities = items.map(
        (item: {
          OpportunityID: any;
          Id: { toString: () => any };
          Title: any;
          Customer: any;
          OpportunityStatus: any;
          ReportDate: any;
          OppAmount: any;
          Business: any;
          BusinessUnit: any;
          OEM: any;
          EndCustomer: any;
          KeyContact: any;
          DecisionMaker: any;
          TentativeStartDate: any;
          TentativeDecisionDate: any;
          Currency: any;
          RiskLevel: any;
          Strategic: any;
          OppComments: any;
        }) => ({
          Id: item.Id.toString(),
          OpportunityID: item.OpportunityID || item.Id.toString(),
          Title: item.Title || "",
          // Account: item.Customer,
          // Stage: item.OpportunityStatus,
          // Owner: "Unknown", // optionally extend with owner logic
          // CloseDate: item.ReportDate,
          // Probability: 50, // static or custom logic
          // Revenue: item.AmountEUR ? `€${item.AmountEUR}` : "€0",
          OpportunityStatus: item.OpportunityStatus,
          Business: item.Business,
          BusinessUnit: item.BusinessUnit,
          OEM: item.OEM,
          EndCustomer: item.EndCustomer,
          Customer: item.Customer,
          KeyContact: item.KeyContact,
          DecisionMaker: item.DecisionMaker,
          TentativeStartDate: item.TentativeStartDate,
          TentativeDecisionDate: item.TentativeDecisionDate,
          Currency: item.Currency,
          RiskLevel: item.RiskLevel,
          Strategic: item.Strategic,
          OppComments: item.OppComments,
          OppAmount: item.OppAmount ? parseFloat(item.OppAmount) : 0,
          ReportDate: item.ReportDate || new Date().toISOString(),
        })
      );
      console.log("Fetched opportunities:", opportunities);
      mockOpportunities.push(...opportunities);
      this.setState({
        filteredOpportunities: opportunities,
        allOpportunities: opportunities,
      });
    } catch (err) {
      console.error("Failed to fetch opportunities", err);
    }
  };
  // private generateNextOpportunityID = async (): Promise<string> => {
  //   const currentYear = new Date().getFullYear();

  //   // Get the latest Opportunity with matching prefix
  //   const items = await this.sp.web.lists
  //     .getByTitle("CWSalesRecords")
  //     .items.filter(`startswith(Title, 'OPP-${currentYear}')`)
  //     .orderBy("Created", false)
  //     .top(1)
  //     .select("Title")();

  //   let nextNumber = 1;
  //   if (items.length > 0) {
  //     const lastTitle = items[0].Title; // e.g., OPP-2025-0012
  //     const parts = lastTitle.split("-");
  //     const number = parseInt(parts[2]);
  //     if (!isNaN(number)) {
  //       nextNumber = number + 1;
  //     }
  //   }

  //   const padded = nextNumber.toString().padStart(4, "0");
  //   return `OPP-${currentYear}-${padded}`;
  // };

  private generateNextOpportunityID = async (): Promise<string> => {
  const currentYear = new Date().getFullYear();
  let prefix = "OPP"; // default

  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'OPPConfig'")
      .top(1)();

    if (configItems.length > 0) {
      const config = JSON.parse(configItems[0].MultiValue || "{}");
      prefix = config.prefix || prefix;
    }
  } catch (err) {
    console.warn("OPP prefix config not found or invalid, using default.");
  }

  const items = await this.sp.web.lists
    .getByTitle("CWSalesRecords")
    .items.filter(`startswith(Title, '${prefix}-${currentYear}')`)
    .orderBy("Created", false)
    .top(1)
    .select("Title")();

  let nextNumber = 1;
  if (items.length > 0) {
    const lastTitle = items[0].Title;
    const parts = lastTitle.split("-");
    const number = parseInt(parts[2]);
    if (!isNaN(number)) {
      nextNumber = number + 1;
    }
  }

  const padded = nextNumber.toString().padStart(4, "0");
  return `${prefix}-${currentYear}-${padded}`;
};

  private handleOpportunityChange = (field: string, value: any): void => {
    this.setState((prevState) => ({
      opportunityForm: {
        ...prevState.opportunityForm,
        [field]: value,
      },
    }));
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

  private submitOpportunity = async (): Promise<void> => {
    const { selectedOpportunity, opportunityForm } = this.state;
    try {
      let itemId: number;
      const user = this.props.context.pageContext.user.displayName;
      const now = new Date();

      const metadata = {
        LastModifiedBy: user,
        LastModifiedDate: now.toISOString(),
      };

      const updatedForm = {
        ...opportunityForm,
        ...metadata,
      };

      if (selectedOpportunity) {
        // Update existing
        const item = await this.sp.web.lists
          .getByTitle("CWSalesRecords")
          .items.getById(selectedOpportunity.Id);
        await item.update(opportunityForm);
        itemId = selectedOpportunity.Id;

        await this.logAuditEntry(itemId.toString(), "Updated", updatedForm);
      } else {
        const newID = await this.generateNextOpportunityID();
        const newForm = {
          ...opportunityForm,
          Title: newID, // Save to Title
          // Save to OpportunityID field
        };
        // Create new
        const added = await this.sp.web.lists
          .getByTitle("CWSalesRecords")
          .items.add(newForm);

        itemId = added.Id;

        await this.logAuditEntry(itemId.toString(), "Created", newForm);
                try {
          await this.sp.web.folders.addUsingPath(
            `Shared Documents/${newID}`
          );
        } catch (error) {
          if (!error.message.includes("already exists")) {
            console.error("Error creating folder:", error);
            throw error;
          }
          console.log("Folder already exists, skipping creation.");
        }
      }

      // Upload file(s)

        // const folderName = this.state.opportunityForm.Title;
        //  await this.ensureFolderExists(folderName);




        alert("Opportunity uploaded successfully.");
      
      await this.fetchOpportunities();
      this.setState({
        selectedOpportunity: null,
        opportunityForm: {},
        selectedFile: [],
        showSales: false,
        showForm: false,
      });
    } catch (error) {
      console.error("Error saving Opportunity:", error);
      alert("Failed to save Opportunity.");
    }
  };
  // private ensureFolderExists = async (folderPath: string) => {
  //   try {
  //     // Try to get the folder
  //     await this.sp.web.getFolderByServerRelativePath(`Shared Documents/${folderPath}`)();
  //     console.log("Folder already exists:", folderPath);
  //   } catch (error) {
  //     if (
  //       error.message &&
  //       error.message.includes("does not exist")
  //     ) {
  //       // Folder doesn't exist — safe to create
  //       await this.sp.web.folders.addUsingPath(`Shared Documents/${folderPath}`);
  //       console.log("Folder created:", folderPath);
  //     } else {
  //       console.error("Error checking folder existence:", error);
  //       throw error;
  //     }
  //   }
  // };

  private deleteOpportunity = async (id: number): Promise<void> => {
    const confirmDelete = window.confirm(
      "Are you sure you want to delete this opportunity?"
    );
    if (!confirmDelete) return;

    try {
      await this.sp.web.lists
        .getByTitle("CWSalesRecords")
        .items.getById(id)
        .recycle();
      alert("Opportunity moved to recycle bin.");
      await this.fetchOpportunities();
    } catch (error) {
      console.error("Failed to delete opportunity:", error);
      alert("Failed to delete the opportunity.");
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
    this.setState({
      visibleColumns: visibleColumns.filter(
        (col): col is IColumn => col !== undefined
      ),
    });
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
  capitalizeWord(word: string): string {
    const trimmed = word.trim(); // Remove whitespace
    if (!trimmed) return ""; // Handle empty string case
    return trimmed.charAt(0).toUpperCase() + trimmed.slice(1).toLowerCase();
  }
  private saveCustomerToList = async () => {
    const { Customer, PersonResponsible, City } = this.state.customerForm;

    if (!Customer.trim()) {
      alert("Customer name is required");
      return;
    }
    console.log("Saving customer:", Customer, PersonResponsible, City);
    try {
      await this.sp.web.lists.getByTitle("CWSalesCustomer").items.add({
        Title: Customer,
        Customer:
          Customer.charAt(0).toUpperCase() + Customer.slice(1).toLowerCase(),
        City: City || "",
      });

      alert("Customer saved successfully!");

      this.setState({
        isCustomerPanelOpen: false,
        customerForm: {
          Customer: "",
          PersonResponsible: null,
          City: "",
        },
      });

      // Reload dropdown options
      this.loadCustomerOptions();
    } catch (error) {
      console.error("Error saving customer:", error);
      alert("Failed to save customer.");
    }
  };
  private loadCustomerOptions = async () => {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesCustomer")
      .items.select("Id", "Customer")();

    const options = items.map((item) => ({
      key: item.Customer,
      text: item.Customer,
    }));

    this.setState({ customerOptions: options });
  };
  private saveKeyContact = async () => {
    const form = this.state.keyContactForm;

    if (!form.Contact.trim()) {
      alert("Contact name is required.");
      return;
    }

    try {
      await this.sp.web.lists.getByTitle("CWSalesKeyContact").items.add({
        Title: form.Contact,
        Customer: form.Customer,
        Contact: form.Contact,
        Email: form.Email,
        Address: form.Address,
        City: form.City,
        BusinessPhone: form.BusinessPhone,
        MobileNumber: form.MobileNumber,
        Designation: form.Designation,
        Department: form.Department,
      });

      alert("Key Contact added successfully!");

      this.setState({
        isKeyContactPanelOpen: false,
        keyContactForm: {
          Customer: "",
          Contact: "",
          Email: "",
          Address: "",
          City: "",
          BusinessPhone: "",
          MobileNumber: "",
          Designation: "",
          Department: "",
        },
      });

      this.loadKeyContactOptions(); // Reload dropdown
    } catch (error) {
      console.error("Error adding key contact:", error);
      alert("Failed to add key contact.");
    }
  };
  private loadKeyContactOptions = async () => {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesKeyContact")
      .items.select("Id", "Contact")();

    const options = items.map((item) => ({
      key: item.Contact,
      text: item.Contact,
    }));

    this.setState({ keyContactOptions: options });
  };
  private loadOpportunityFiles = async (opportunityId: string) => {
    const folderName = `${opportunityId}`;

    try {
      const files = await this.sp.web
        .getFolderByServerRelativePath(`Shared Documents/${folderName}`)
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
  // private deleteFileFromLibrary = async (fileName: string) => {
  //   if (!this.state.selectedOpportunity) return;

  //   const folderName = `${this.state.selectedOpportunity.OpportunityID}`;
  //   const confirm = window.confirm(
  //     `Delete file "${fileName}" from SharePoint?`
  //   );

  //   if (!confirm) return;

  //   try {
  //     const files = await this.sp.web
  //       .getFolderByServerRelativePath(`Shared Documents/${folderName}`)
  //       .files.filter(`Name eq '${fileName}'`)();

  //     if (files.length > 0) {
  //       await this.sp.web
  //         .getFileByServerRelativePath(files[0].ServerRelativeUrl)
  //         .recycle(); // OR .delete() for permanent
  //       alert("File deleted.");
  //       this.loadOpportunityFiles(this.state.selectedOpportunity.OpportunityID);
  //     } else {
  //       alert("File not found.");
  //     }
  //   } catch (err) {
  //     console.error("Error deleting file:", err);
  //     alert("Failed to delete the file.");
  //   }
  // };
  private defaultCurrency = "EUR";

  // private exchangeRates: ExchangeRates = {
  //   USD: 1.07,
  //   EUR: 1.0,
  //   GBP: 0.85,
  // };
  private loadCurrencyConfig = async () => {
    try {
      const configItems = await this.sp.web.lists
        .getByTitle("CWSalesConfiguration")
        .items.select("Title", "DefaultCurrency", "MultiValue")();

      const config = configItems.find((item) => item.Title === "Currency");

      if (config) {
        const parsedRates = JSON.parse(config.MultiValue);
        const currencyOptions: IDropdownOption[] = Object.keys(
          parsedRates.rates
        ).map((code) => ({
          key: code,
          text: code,
        }));

        this.setState({
          defaultCurrency: parsedRates.base,
          exchangeRates: parsedRates.rates,
          currencyOptions,
          opportunityForm: {
            ...this.state.opportunityForm,
            Currency: parsedRates.base,
          },
        });
        // this.setState({
        //   defaultCurrency: config.DefaultCurrency,
        //   exchangeRates: parsedRates,
        // });
      } else {
        console.warn("Currency config not found. Using default.");
      }
    } catch (err) {
      console.error("Error loading currency configuration:", err);
    }
  };
  // private async loadCurrenciesFromAPI(): Promise<void> {
  //   try {
  //     const response = await fetch("https://api.exchangerate-api.com/v4/latest/EUR");
  //     const data = await response.json();

  //     const currencyOptions: IDropdownOption[] = Object.keys(data.rates).map(
  //       (code) => ({
  //         key: code,
  //         text: code,
  //       })
  //     );

  //     this.setState({
  //       defaultCurrency: data.base,
  //       exchangeRates: data.rates,
  //       currencyOptions, // you need to add this to state
  //     });
  //   } catch (err) {
  //     console.error("Failed to load currency list:", err);
  //   }
  // }
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

  componentDidMount() {
    this.fetchOpportunities();
    console.log(this.props.context);
    console.log(this.props.context.pageContext.web.absoluteUrl);
    console.log(this.props.salesProps);
    console.log(this.props.children);
    this.loadCustomerOptions();
    this.loadKeyContactOptions();
    this.loadCurrencyConfig();
    this.loadDateFormatConfig();
    // this.loadCurrenciesFromAPI();
    // this.fetchAllQuotations();
    // this.fetchAllPOs();
    // $(".editOpportunityItem").hide();
    // Load the column configuration from local storage
    const savedColumnConfig = localStorage.getItem("editOpportunityItem");
console.log("Saved Column Config:", savedColumnConfig);
    // If there is a saved configuration, update the state
    if (savedColumnConfig) {
      const parsedConfig = JSON.parse(savedColumnConfig);
      const visibleColumns = this.getConfiguredColumns(parsedConfig);
      this.setState({
        visibleColumns: visibleColumns.filter(
          (col): col is IColumn => col !== undefined
        ),
      });
    } else {
      // Provide a default configuration if no saved configuration is found
      this.setState({ visibleColumns: this._columns });
    }
    console.log("Visible Columns");
    console.log(this.state.visibleColumns);
  }
  componentWillUnmount() {
    if (this.state.filePreviews) {
      this.state.filePreviews.forEach((p) => URL.revokeObjectURL(p.previewUrl));
    }
  }

  private onSearchChange = (_ev: any, newValue?: string) => {
    const searchQuery = newValue || "";
    const lowerValue = searchQuery.toLowerCase();

    const filtered = this.state.allOpportunities.filter(
      (opp) =>
        opp.Title?.toLowerCase().includes(lowerValue) ||
        opp.OpportunityID?.toLowerCase().includes(lowerValue) ||
        opp.Business?.toLowerCase().includes(lowerValue) ||
        opp.EndCustomer?.toLowerCase().includes(lowerValue)
    );
    console.log("Filtered Opportunities:", filtered);
    this.setState({ searchQuery, filteredOpportunities: filtered });
  };

  getPagedItems = (): Opportunity[] => {
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
    XLSX.utils.book_append_sheet(workbook, worksheet, "Opportunities");
    XLSX.writeFile(workbook, "Opportunities.xlsx");
  };
  getColumns = (): IColumn[] => [
    {
      key: "OpportunityID",
      name: "Opportunity",
      fieldName: "OpportunityID",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "Title",
      name: "Opportunity ID",
      fieldName: "Title",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "TentativeStartDate",
      name: "Tentative Start Date",
      fieldName: "TentativeStartDate",
      minWidth: 70,
      isResizable: true,
      // onRender: (item: Opportunity) => {
      //   const rawDate = item.TentativeStartDate;
      //   if (!rawDate) return "";
      //   const date = new Date(rawDate);
      //   return date.toLocaleDateString("en-GB", {
      //     day: "2-digit",
      //     month: "short",
      //     year: "numeric",
      //   }); // Output: 08-Jun-2025
      // },
      onRender: (item: Opportunity) => {
  if (!item.TentativeStartDate) return "";
  return formatDate(item.TentativeStartDate, this.dateFormat);
}
    },
    {
      key: "Business",
      name: "Business",
      fieldName: "Business",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "BusinessUnit",
      name: "Business Unit",
      fieldName: "BusinessUnit",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "Customer",
      name: "Customer",
      fieldName: "Customer",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "DecisionMaker",
      name: "Decision Maker",
      fieldName: "DecisionMaker",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "OpportunityStatus",
      name: "Opportunity Status",
      fieldName: "OpportunityStatus",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "OEM",
      name: "OEM",
      fieldName: "OEM",
      minWidth: 70,
      isResizable: true,
    },
    {
      key: "actions",
      name: "",
      fieldName: "actions",
      minWidth: 40,
      maxWidth: 40,
      isResizable: false,
      onRender: (item: Opportunity) => {
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
                    this.loadOpportunityFiles(item.OpportunityID);
                    console.log("Edit clicked for item:", item);
                    this.setState({
                      selectedOpportunity: item,
                      opportunityForm: { ...item },
                      showSales: true,
                    });
                    console.log(this.state.opportunityForm);
                  },
                },
                {
                  key: "delete",
                  text: "Delete",
                  iconProps: { iconName: "Delete" },
                  onClick: () => this.deleteOpportunity(item.Id),
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
      //panelOpen,
      searchQuery,
      // filteredOpportunities,
    } = this.state;

    const commandBarItems = [
      {
        key: "new",
        text: "New Opportunity",
        iconProps: { iconName: "Add" },
        onClick: () =>
          this.setState({
            showSales: true,
            opportunityForm: {},
            selectedOpportunity: null,
          }),
        styles: { primaryButtonStyles },
      },
      // {
      //   key: "edit",
      //   text: "Edit Opportunity",
      //   iconProps: { iconName: "Edit" },
      //   //disabled: !selectedOpportunity,
      //   onClick: () => {
      //     this.setState({
      //       showSales: true,
      //       opportunityForm: selectedOpportunity,

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
        key: "columns",
        text: "Edit Columns",
        iconProps: { iconName: "ColumnOptions" },
        onClick: this.openColumnsPanel,
        styles: { primaryButtonStyles },
      },
    ];

    // if (this.state.showSales) {
    //   return (
    //     <Sales
    //       description={this.props.salesProps.description}
    //       isDarkTheme={false}
    //       environmentMessage={this.props.salesProps.environmentMessage}
    //       hasTeamsContext={false}
    //       userDisplayName={this.props.salesProps.userDisplayName}
    //       sp={this.props.salesProps.sp}
    //       context={this.props.salesProps.context}
    //       View="Opportunity"
    //     />
    //   );
    // }
    const { OppAmount, Currency } = this.state.opportunityForm;
    const { defaultCurrency, exchangeRates } = this.state;

    const convertedAmount = convertToDefaultCurrency(
      OppAmount || 0,
      Currency || defaultCurrency,
      defaultCurrency,
      exchangeRates
    );
    return (
      <div style={{ width: "100%", height: "100vh" }} id="sales-webpart-root">
        <MessageBar messageBarType={MessageBarType.info}>
          Welcome to the Opportunity Viewer
        </MessageBar>
        <TextField
          placeholder="Search opportunities..."
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
          //onItemInvoked={(item) => this.onEdit(item as Opportunity)}
          selectionMode={SelectionMode.none}
          // selectionPreservedOnEmptyClick={true}

          //onItemInvoked={this.onRowClick}
          //onActiveItemChanged={(item) => this.onRowClick(item as Opportunity)}
        />
        </div>
</div>
        {console.log("Filtered Opportunities:", this.state.opportunityForm)}
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
              localStorageKey="editOpportunityItem"
            />
          )}
        </div>

        <Panel
          isOpen={this.state.showSales}
          onDismiss={() =>
            this.setState({
              showSales: false,
              selectedOpportunity: null,
              opportunityForm: {},
              selectedFile: [],
              filePreviews: [],
              selectedFileName: null,
              selectedPreviewFile: null,
              isViewerOpen: false,
              isCustomerPanelOpen: false,
              isKeyContactPanelOpen: false,
              existingFiles: [],
            })
          }
          headerText={
            !selectedOpportunity ? "Create Opportunity" : "Edit Opportunity"
          }
          type={PanelType.extraLarge}
          styles={{
            main: {
              backgroundColor: "#f5f5f5",
            },
          }}
        >

            <Stack
              tokens={{ childrenGap: 12 }}
              styles={{ root: { width: "100%" } }}
            >
              <div style={{ padding: 16 }}>
                {/* <div style={gridStyle}> */}
                <TextField
                  label="Opportunity"
                  value={this.state.opportunityForm?.OpportunityID || ""}
                  onChange={(_, v) =>
                    this.handleOpportunityChange("OpportunityID", v)
                  }
                />
                <Stack
                  horizontal
                  tokens={{ childrenGap: 12 }}
                  styles={{ root: { width: "100%" } }}
                >
                  <TextField
                    label="Business"
                    value={this.state.opportunityForm?.Business || ""}
                    // value={selectedOpportunity?.business || ""}
                    onChange={(_, v) =>
                      this.handleOpportunityChange("Business", v)
                    }
                    styles={{ root: { width: "50%" } }}
                  />
                  <TextField
                    label="Business Unit"
                    value={this.state.opportunityForm?.BusinessUnit || ""}
                    // value={selectedOpportunity?.businessUnit || ""}
                    onChange={(_, v) =>
                      this.handleOpportunityChange("BusinessUnit", v)
                    }
                    styles={{ root: { width: "50%" } }}
                  />
                </Stack>
                <Stack
                  horizontal
                  tokens={{ childrenGap: 12 }}
                  styles={{ root: { width: "100%" } }}
                >
                  <TextField
                    label="OEM"
                    value={this.state.opportunityForm?.OEM || ""}
                    // value={selectedOpportunity?.oem || ""}
                    onChange={(_, v) => this.handleOpportunityChange("OEM", v)}
                    styles={{ root: { width: "25%" } }}
                  />
                  {/* <TextField
                label="End Customer"
                value={this.state.opportunityForm?.EndCustomer || ""}
                // value={selectedOpportunity?.endCustomer || ""}
                onChange={(_, v) =>
                  this.handleOpportunityChange("EndCustomer", v)
                }
                styles={{ root: { width: "20%" } }}
              /> */}
                  {/* <TextField
                label="Customer"
                value={this.state.opportunityForm?.EndCustomer || ""}
                // value={selectedOpportunity?.endCustomer || ""}
                onChange={(_, v) => this.handleOpportunityChange("Customer", v)}
              /> */}
                  {/* <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 4 }}>
                <Dropdown
                  label="Customer"
                  options={this.state.customerOptions}
                  selectedKey={this.state.opportunityForm.Customer}
                  onChange={(_, option) =>
                    this.handleOpportunityChange("Customer", option?.key)
                  }
                  styles={{
                    root: { width: "100%" }, // make dropdown fill the container width
                  }}
                />
                <DefaultButton
                  iconProps={{ iconName: "Add" }}
                  onClick={() => this.setState({ isCustomerPanelOpen: true })}
                />
              </Stack> */}
                  <Stack styles={{ root: { width: "25%" } }}>
                    <Stack
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 4 }}
                    >
                      <Label>Customer</Label>
                      <IconButton
                        iconProps={{ iconName: "Add" }}
                        title="Add Customer"
                        ariaLabel="Add Customer"
                        onClick={() =>
                          this.setState({ isCustomerPanelOpen: true })
                        }
                        styles={{
                          root: {
                            padding: 2,
                            height: 24,
                            width: 24,
                          },
                          icon: {
                            fontSize: 12,
                          },
                        }}
                      />
                    </Stack>

                    <Dropdown
                      options={this.state.customerOptions}
                      selectedKey={this.state.opportunityForm.Customer}
                      onChange={(_, option) =>
                        this.handleOpportunityChange("Customer", option?.key)
                      }
                      styles={{ root: { width: "100%" } }}
                    />
                  </Stack>
                  {/* <TextField
                label="Key Contact"
                value={this.state.opportunityForm?.KeyContact || ""}
                // value={selectedOpportunity?.keyContact || ""}
                onChange={(_, v) =>
                  this.handleOpportunityChange("KeyContact", v)
                }
              /> */}
                  {/* <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 4 }}>
                <Dropdown
                  label="Key Contact"
                  options={this.state.keyContactOptions}
                  selectedKey={this.state.opportunityForm.KeyContact}
                  onChange={(_, option) =>
                    this.handleOpportunityChange("KeyContact", option?.key)
                  }
                  styles={{
                    root: { width: "100%" }, // make dropdown fill the container width
                  }}
                />
                <DefaultButton
                  iconProps={{ iconName: "Add" }}
                  onClick={() => this.setState({ isKeyContactPanelOpen: true })}
                />
              </Stack> */}
                  <Stack styles={{ root: { width: "25%" } }}>
                    <Stack
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 4 }}
                    >
                      <Label>Key Contact</Label>
                      <IconButton
                        iconProps={{ iconName: "Add" }}
                        title="Add Key Contact"
                        ariaLabel="Add Key Contact"
                        onClick={() =>
                          this.setState({ isKeyContactPanelOpen: true })
                        }
                        styles={{
                          root: {
                            padding: 2,
                            height: 24,
                            width: 24,
                          },
                          icon: {
                            fontSize: 12,
                          },
                        }}
                      />
                    </Stack>

                    <Dropdown
                      options={this.state.keyContactOptions}
                      selectedKey={this.state.opportunityForm.KeyContact}
                      onChange={(_, option) =>
                        this.handleOpportunityChange("KeyContact", option?.key)
                      }
                      styles={{
                        root: { width: "100%" },
                      }}
                    />
                  </Stack>

                  <Dropdown
                    label="Decision Maker"
                    options={this.state.keyContactOptions}
                    selectedKey={this.state.opportunityForm.DecisionMaker}
                    onChange={(_, option) =>
                      this.handleOpportunityChange("DecisionMaker", option?.key)
                    }
                    styles={{ root: { width: "25%" } }}
                  />
                </Stack>
                {/* <TextField
                label="Decision Maker"
                value={this.state.opportunityForm?.DecisionMaker || ""}
                // value={selectedOpportunity?.decisionMaker || ""}
                onChange={(_, v) =>
                  this.handleOpportunityChange("DecisionMaker", v)
                }
              /> */}
                <Stack
                  horizontal
                  tokens={{ childrenGap: 12 }}
                  styles={{ root: { width: "100%" } }}
                >
                  <DatePicker
                    label="Report Date"
                    // value={
                    //   this.state.opportunityForm.CloseDate
                    //     ? new Date(this.state.opportunityForm.CloseDate)
                    //     : undefined
                    // }
                    value={
                      this.state.opportunityForm.ReportDate
                        ? new Date(this.state.opportunityForm.ReportDate)
                        : undefined
                    }
                    onSelectDate={(d) =>
                      this.handleOpportunityChange("ReportDate", d)
                    }
                    styles={{ root: { width: "33%" } }}
                  />
                  <DatePicker
                    label="Tentative Start Date"
                    value={
                      this.state.opportunityForm.TentativeStartDate
                        ? new Date(
                            this.state.opportunityForm.TentativeStartDate
                          )
                        : undefined
                    }
                    // value={
                    //   this.state.opportunityForm.TentativeStartDate
                    //     ? new Date(this.state.opportunityForm.TentativeStartDate)
                    //     : undefined
                    // }
                    onSelectDate={(d) =>
                      this.handleOpportunityChange("TentativeStartDate", d)
                    }
                    styles={{ root: { width: "33%" } }}
                  />
                  <DatePicker
                    label="Tentative Decision Date"
                    value={
                      this.state.opportunityForm.TentativeDecisionDate
                        ? new Date(
                            this.state.opportunityForm.TentativeDecisionDate
                          )
                        : undefined
                    }
                    onSelectDate={(d) =>
                      this.handleOpportunityChange("TentativeDecisionDate", d)
                    }
                    styles={{ root: { width: "33%" } }}
                  />
                </Stack>
                <Stack
                  horizontal
                  tokens={{ childrenGap: 12 }}
                  styles={{ root: { width: "100%" } }}
                >
                  <Stack.Item grow styles={{ root: { width: "33%" } }}>
                    <TextField
                      label={`Amount (${this.state.opportunityForm.Currency})`}
                      type="number"
                      value={
                        this.state.opportunityForm.OppAmount
                          ? this.state.opportunityForm.OppAmount.toString()
                          : ""
                      }
                      // value={selectedOpportunity?.revenue || ""}
                      onChange={(_, v) =>
                        this.handleOpportunityChange(
                          "OppAmount",
                          parseFloat(v || "0")
                        )
                      }
                      // styles={{ root: { width: "33%" } }}
                    />
                  </Stack.Item>
                  <Stack.Item grow styles={{ root: { width: "33%" } }}>
                    <ComboBox
                      label="Currency"
                      selectedKey={
                        this.state.opportunityForm.Currency || undefined
                      }
                      options={this.state.currencyOptions}
                      onChange={(_, o) =>
                        this.handleOpportunityChange("Currency", o?.key)
                      }
                      allowFreeform={false}
                      autoComplete="on"
                      //  styles={{ root: { width: "33%" } }}
                      calloutProps={{
                        styles: {
                          root: {
                            width: 100, // <-- this controls dropdown menu width
                          },
                        },
                      }}
                    />
                  </Stack.Item>
                  <Stack.Item grow styles={{ root: { width: "33%" } }}>
                    <TextField
                      label={`Converted Amount (${this.defaultCurrency})`}
                      value={convertedAmount.toString()}
                      disabled
                          onChange={(_, v) =>
                        this.handleOpportunityChange(
                          "ConvertedAmount",
                          parseFloat(v || "0")
                        )
                      }
                      // styles={{ root: { width: "33%" } }}
                    />
                  </Stack.Item>
                </Stack>
                <Stack
                  horizontal
                  tokens={{ childrenGap: 12 }}
                  styles={{ root: { width: "100%" } }}
                >
                  <Dropdown
                    label="Risk Level"
                    selectedKey={
                      this.state.opportunityForm.RiskLevel || undefined
                    }
                    // selectedKey={selectedOpportunity?.riskLevel || undefined}
                    options={riskLevels}
                    onChange={(_, o) =>
                      this.handleOpportunityChange("RiskLevel", o?.key)
                    }
                    styles={{ root: { width: "33%" } }}
                  />
                  <Dropdown
                    label="Strategic"
                    selectedKey={
                      this.state.opportunityForm.Strategic || undefined
                    }
                    // selectedKey={selectedOpportunity?.strategic || undefined}
                    options={strategicOptions}
                    onChange={(_, o) =>
                      this.handleOpportunityChange("Strategic", o?.key)
                    }
                    styles={{ root: { width: "33%" } }}
                  />
                  <Dropdown
                    label="Opportunity Status"
                    selectedKey={
                      this.state.opportunityForm.OpportunityStatus || undefined
                    }
                    // selectedKey={selectedOpportunity?.stage || undefined}
                    options={statusOptions}
                    onChange={(_, o) =>
                      this.handleOpportunityChange("OpportunityStatus", o?.key)
                    }
                    styles={{ root: { width: "33%" } }}
                  />
                </Stack>
                <TextField
                  label="Comments"
                  multiline
                  rows={3}
                  value={this.state.opportunityForm.OppComments || ""}
                  // value={selectedOpportunity?.comments || ""}
                  onChange={(_, v) =>
                    this.handleOpportunityChange("OppComments", v)
                  }
                />
                <div>
                  {/* <input
                  type="file"
                  id="attachmentUpload"
                  onChange={(e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                      this.setState({ selectedFile: file });
                    }
                  }}
                /> */}

                </div>
              </div>
              {/* </div> */}
            </Stack>

        
          <Stack
            horizontal
            tokens={{ childrenGap: 12 }}
            style={{ marginTop: 24 }}
          >
            <PrimaryButton
              text="Submit"
              onClick={this.submitOpportunity}
              styles={primaryButtonStyles}
            />
            {/* <PrimaryButton
                      text="View Opportunities"
                      onClick={() => this.setState({ tab: 'View Opportunities' })}
                      styles={primaryButtonStyles}
                    /> */}
          </Stack>
          <Panel
            isOpen={this.state.isCustomerPanelOpen}
            onDismiss={() => this.setState({ isCustomerPanelOpen: false })}
            headerText="Add New Customer"
            type={PanelType.smallFixedFar}
            styles={{
              main: {
                backgroundColor: "#f5f5f5",
              },
            }}
          >
            <TextField
              label="Customer"
              value={this.state.customerForm.Customer}
              onChange={(_, val) =>
                this.setState((prev) => ({
                  customerForm: { ...prev.customerForm, Customer: val || "" },
                }))
              }
            />
            <PeoplePicker
              context={this._peoplePickerContext}
              titleText="Person Responsible"
              personSelectionLimit={1}
              showtooltip={true}
              required={false}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]} // Only people
              resolveDelay={300}
              onChange={(items: any[]) => {
                this.setState((prev) => ({
                  customerForm: {
                    ...prev.customerForm,
                    PersonResponsible: items[0]?.loginName || null,
                  },
                }));
                // console.log("Selected Person:", items[0]);
              }}
              ensureUser
            />
            <TextField
              label="City"
              value={this.state.customerForm.City}
              onChange={(_, val) =>
                this.setState((prev) => ({
                  customerForm: { ...prev.customerForm, City: val || "" },
                }))
              }
            />
            <PrimaryButton
              text="Save"
              onClick={this.saveCustomerToList}
              styles={primaryButtonStyles}
            />
          </Panel>
          {/* <Panel
          isOpen={this.state.isKeyContactPanelOpen}
          onDismiss={() => this.setState({ isKeyContactPanelOpen: false })}
          headerText="Add New Key Contact"
          type={PanelType.smallFixedFar}
        >
          <Stack tokens={{ childrenGap: 10 }}>
            {(
              [
                "Customer",
                "Contact",
                "Email",
                "Address",
                "City",
                "BusinessPhone",
                "MobileNumber",
                "Designation",
                "Department",
              ] as (keyof typeof this.state.keyContactForm)[]
            ).map((field) => (
              <TextField
                key={field}
                label={field}
                value={this.state.keyContactForm[field]}
                onChange={(_, val) =>
                  this.setState((prev) => ({
                    keyContactForm: {
                      ...prev.keyContactForm,
                      [field]: val || "",
                    },
                  }))
                }
              />
            ))}

            <PrimaryButton
              text="Save"
              onClick={this.saveKeyContact}
              styles={primaryButtonStyles}
            />
          </Stack>
        </Panel> */}
          <Panel
            isOpen={this.state.isKeyContactPanelOpen}
            onDismiss={() => this.setState({ isKeyContactPanelOpen: false })}
            headerText="Add New Key Contact"
            type={PanelType.smallFixedFar}
            styles={{
              main: {
                backgroundColor: "#f5f5f5",
              },
            }}
          >
            <Stack tokens={{ childrenGap: 10 }}>
              {(
                [
                  "Customer",
                  "Contact",
                  "Email",
                  "Address",
                  "City",
                  "BusinessPhone",
                  "MobileNumber",
                  "Designation",
                  "Department",
                ] as (keyof typeof this.state.keyContactForm)[]
              ).map((field) =>
                field === "Customer" ? (
                  <Dropdown
                    key={field}
                    label={field}
                    options={this.state.customerOptions}
                    selectedKey={this.state.keyContactForm.Customer}
                    onChange={(_, option) =>
                      this.setState((prev) => ({
                        keyContactForm: {
                          ...prev.keyContactForm,
                          Customer: option?.key as string,
                        },
                      }))
                    }
                    styles={{ root: { width: "100%" } }}
                  />
                ) : (
                  <TextField
                    key={field}
                    label={field}
                    value={this.state.keyContactForm[field]}
                    onChange={(_, val) =>
                      this.setState((prev) => ({
                        keyContactForm: {
                          ...prev.keyContactForm,
                          [field]: val || "",
                        },
                      }))
                    }
                  />
                )
              )}

              <PrimaryButton
                text="Save"
                onClick={this.saveKeyContact}
                styles={primaryButtonStyles}
              />
            </Stack>
          </Panel>
        </Panel>
      </div>
    );
  }
}

export default OpportunityViewer;
