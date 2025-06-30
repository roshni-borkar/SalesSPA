/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
// IPOViewerState.ts
import { IDropdownOption, IColumn } from "@fluentui/react";

export interface PO {
  POID: string;
  OpportunityID: string;
  POReceivedDate: Date | string;
  POStatus: string;
  POAmount: string;
  POCurrency: string;
  CustomerPONumber: string;
  lineItemsJSON: string;
  POQuoteID: string;
  POComments: string;
  IsChildPO: boolean | string;
  lineItems?: any;
}

export interface IPOViewerState {
  selectedOpportunity: PO | null;
  panelOpen: boolean;
  searchQuery: string;
  filteredOpportunities: PO[];
  purchaseOrderForm: any;
  quoteOptions: IDropdownOption[];
  opportunityOptions: IDropdownOption[];
  lineItems: any[];
  purchaseOrders: any[];
  showSales: boolean;
  currentPage: number;
  pageSize: number;
  selectedFile: File[];
  selectedFileUrl: string;
  isViewerOpen: boolean;
  filePreviews?: { file: File; previewUrl: string }[];
  selectedPreviewFile?: File;
  selectedFileName?: string;
  allOpportunities: PO[];
  isColumnsPanelOpen: boolean;
  visibleColumns: IColumn[];
  existingFiles: { name: string; url: string }[];
  parentPOOptions: IDropdownOption[];
exchangeRates: { [currency: string]: number };
defaultCurrency: string;
OPPCurrency: string;

}
