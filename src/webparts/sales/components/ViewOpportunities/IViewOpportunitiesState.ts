/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @rushstack/no-new-null */
// IOppViewerState.ts

import { IColumn, IDropdownOption, IComboBoxOption } from "@fluentui/react";

export interface Opportunity {
  Id: number;
  OpportunityID: string;
  Title: string;
  Business: string;
  BusinessUnit: string;
  OEM: string;
  EndCustomer: string;
  Customer: string;
  KeyContact: string;
  DecisionMaker: string;
  TentativeStartDate: string;
  TentativeDecisionDate: string;
  Currency: string;
  RiskLevel: string;
  Strategic: string;
  OpportunityStatus: string;
  OppComments: string;
  OppAmount: number;
}

export interface ExchangeRates {
  [currencyCode: string]: number;
}

export interface IOppViewerState {
  selectedOpportunity: Opportunity | null;
  panelOpen: boolean;
  searchQuery: string;
  filteredOpportunities: Opportunity[];
  showSales: boolean;
  opportunityForm: any;
  currentPage: number;
  pageSize: number;
  selectedFile: File[];
  showForm: boolean;
  isViewerOpen: boolean;
  selectedFileUrl: string | null;
  filePreviews?: { file: File; previewUrl: string }[];
  selectedPreviewFile: File | null;
  selectedFileName: string | null;
  isColumnsPanelOpen: boolean;
  visibleColumns: IColumn[];
  isCustomerPanelOpen: boolean;
  customerForm: {
    Customer: string;
    PersonResponsible: string | null;
    City: string;
  };
  customerOptions: IDropdownOption[];
  isKeyContactPanelOpen: boolean;
  keyContactForm: {
    Customer: string;
    Contact: string;
    Email: string;
    Address: string;
    City: string;
    BusinessPhone: string;
    MobileNumber: string;
    Designation: string;
    Department: string;
  };
  keyContactOptions: IDropdownOption[];
  existingFiles: { name: string; url: string }[];
  defaultCurrency: string;
  exchangeRates: ExchangeRates;
  currencyOptions: IComboBoxOption[];
  allOpportunities: Opportunity[];
}
