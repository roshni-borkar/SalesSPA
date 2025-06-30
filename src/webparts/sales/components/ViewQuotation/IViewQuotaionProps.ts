/* eslint-disable @typescript-eslint/no-explicit-any */
// IQuotationViewerProps.ts
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IQuotationViewerProps {
  context: WebPartContext; // Replace 'any' with more specific type if possible
  salesProps: any;         // You can replace this with a better type if available
}
