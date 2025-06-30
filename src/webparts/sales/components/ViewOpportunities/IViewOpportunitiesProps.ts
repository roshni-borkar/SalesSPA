/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IViewOpportunitiesProps {
  context: WebPartContext;
  sp: SPFI;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  fetchOpportunities: () => void;
  componentDidMount: () => void;
  onRowClick: (rowId: string) => void;
  onSearchChange: (searchTerm: string) => void;
  exportToExcel?: () => void; // Add missing properties
  render?: () => React.ReactNode;
  setState?: (state: any) => void;
  forceUpdate?: () => void;
}