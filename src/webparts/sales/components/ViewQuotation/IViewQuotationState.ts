/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
// IQuotationViewerState.ts
import { IDropdownOption } from "@fluentui/react";
import { IColumn } from "@fluentui/react/lib/DetailsList";

export interface Quotation {
  QuoteID: string;
  Title: string;
  QuoteDate: string;
  QuoteRevisionNumber: string;
  QuoteRevenueQuoted: string;
  QuoteBusinessSize: string;
  QuoteAmount: number;
  QuoteCurrency: string;
  QuoteTentativeDecisionDate: Date;
  QuoteComments: string;
}

export interface IQuotationViewerState {
  selectedOpportunity: Quotation | null;
  panelOpen: boolean;
  searchQuery: string;
  filteredOpportunities: Quotation[];
  showSales: boolean;
  quotationForm: any;
  opportunityOptions: IDropdownOption[];
  quotationFile: any;
  currentPage: number;
  pageSize: number;
  filePreviewUrl: string | null;
  selectedFile?: File[];
  selectedFileUrl?: string | null;
  isViewerOpen?: boolean;
  filePreviews?: { file: File; previewUrl: string }[];
  selectedPreviewFile?: File | null;
  selectedFileName?: string | null;
  isColumnsPanelOpen: boolean;
  visibleColumns: IColumn[];
  existingFiles: { name: string; url: string }[];
  allQuotations: Quotation[];
  exchangeRates?: { [key: string]: number };
  QuoteAmount: "";         // EUR
  defaultCurrency: string; 
  OPPCurrency: string;    
RevenueInCustomerCurrency: "" // converted amount

}
