/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-use-before-define */
import * as React from "react";
import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import {
  DetailsList,
  IColumn,
  Stack,
  Text,
  Label,
} from "@fluentui/react";
import { ISalesProps } from "../ISalesProps";

interface IUnifiedRecord {
  Id: number;
  Title?: string;
  OpportunityID: string;
  ReportDate?: string;
  Customer?: string;
  QuoteBusinessSize?: string;
  OpportunityStatus?: string;
  QuoteID?: string;
  POID?: string;
  QuoteRevenueQuoted?: number;
  POAmount?: number;
  POQuoteID?: string;
}

const SalesDashboard: React.FC<ISalesProps> = ({ context }) => {
  const [records, setRecords] = React.useState<IUnifiedRecord[]>([]);
  const [allPOs, setAllPOs] = React.useState<any[]>([]);
const [allQuotations, setAllQuotations] = React.useState<any[]>([]);

  const [selectedOppId, setSelectedOppId] = React.useState<string | null>(null);
  const [expandedQuoteId, setExpandedQuoteId] = React.useState<string | null>(null);

  const sp: SPFI = spfi().using(SPFx(context));

 React.useEffect(() => {
  const fetchData = async () => {
    try {
      const items = await sp.web.lists.getByTitle("CWSalesRecords").items.select("*")();

      setRecords(items);
      const quotes: any[] = [];

items.forEach(item => {
  for (let i = 1; i <= 5; i++) {
    const suffix = i === 1 ? "" : i.toString();
    const QuoteId = item[`QuoteID${suffix}`];
    if (QuoteId) {
      quotes.push({
        QuoteID: QuoteId,
        OpportunityID: item.OpportunityID,
        QuoteRevisionNumber: item[`QuoteRevisionNumber${suffix}`],
        QuoteRevenueQuoted: item[`QuoteRevenueQuoted${suffix}`],
        QuoteAmount: item[`QuoteAmount${suffix}`],
        BusinessSize: item[`BusinessSize${suffix}`],
        QuoteDate: item[`QuoteDate${suffix}`],
        QuoteTentativeDecisionDate: item[`QuoteTentativeDecisionDate${suffix}`],
        QuoteCurrency: item[`Currency${suffix}`],
        QuoteComments: item[`QuoteComments${suffix}`],
      });
    }
  }
});

setAllQuotations(quotes);


      const pos: any[] = [];
      items.forEach((item) => {
        for (let i = 1; i <= 5; i++) {
          const suffix = i === 1 ? "" : i.toString();
          const POID = item[`POID${suffix}`];
          if (POID) {
            pos.push({
              POID,
              OpportunityID: item.OpportunityID,
              POReceivedDate: item[`POReceivedDate${suffix}`],
              POStatus: item[`POStatus${suffix}`],
              POAmount: item[`POAmount${suffix}`],
              Currency: item[`POCurrency${suffix}`],
              CustomerPONumber: item[`CustomerPONumber${suffix}`],
              lineItems: item[`LineItemsJSON${suffix}`],
              POQuoteID: item[`POQuoteID${suffix}`],
              POComments: item[`POComments${suffix}`],
            });
          }
        }
      });

      setAllPOs(pos); // ✅ store flattened POs
    } catch (err) {
      console.error("Failed to load sales data", err);
    }
  };

  fetchData();
}, []);


  const opportunities = React.useMemo(() => {
    const seen = new Set();
    return records.filter(item => {
      if (!seen.has(item.OpportunityID)) {
        seen.add(item.OpportunityID);
        return true;
      }
      return false;
    });
  }, [records]);

  const formatDate = (raw?: string) => {
    if (!raw) return "-";
    const d = new Date(raw);
    return d.toLocaleDateString();
  };

  const renderOpportunityDetails = (oppId: string) => {
    const match = records.find(r => r.OpportunityID === oppId);
    if (!match) return null;

    return (
      <Stack horizontal tokens={{ childrenGap: 32 }} style={{ marginTop: 12 }}>
        <Stack>
          <Label>Report Date</Label>
          <Text>{formatDate(match.ReportDate)}</Text>
        </Stack>
        <Stack>
          <Label>Customer</Label>
          <Text>{match.Customer || "-"}</Text>
        </Stack>
        <Stack>
          <Label>Business Size</Label>
          <Text>{match.QuoteBusinessSize || "-"}</Text>
        </Stack>
        <Stack>
          <Label>Status</Label>
          <Text>{match.OpportunityStatus || "-"}</Text>
        </Stack>
      </Stack>
    );
  };

const renderQuotationList = (oppId: string) => {
  const quotations = allQuotations.filter(q => q.OpportunityID === oppId);

  if (quotations.length === 0) {
    return <Text style={{ paddingLeft: 24 }}>No quotations found for this opportunity</Text>;
  }

  const columns: IColumn[] = [
    { key: 'q1', name: 'Quote ID', fieldName: 'QuoteID', minWidth: 100 },
    { key: 'q2', name: 'Quoted Amount', fieldName: 'QuoteAmount', minWidth: 120 },
    { key: 'q3', name: 'Quote Date', fieldName: 'QuoteDate', minWidth: 100 },
    { key: 'q4', name: 'Currency', fieldName: 'QuoteCurrency', minWidth: 80 },
  ];

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ paddingLeft: 24, marginTop: 12 }}>
      <Label style={{ fontWeight: "bold" }}>Quotations</Label>
      <DetailsList
        items={quotations}
        columns={columns}
        selectionMode={0}
        onActiveItemChanged={(item: any) => {
          setExpandedQuoteId(item.QuoteID === expandedQuoteId ? null : item.QuoteID || null);
        }}
      />

      {expandedQuoteId && renderPurchaseOrders(expandedQuoteId)}
    </Stack>
  );
};


const renderPurchaseOrders = (quoteId: string) => {
  const pos = allPOs.filter(po => po.POQuoteID === quoteId);

  if (pos.length === 0) {
    return <Text style={{ paddingLeft: 24 }}>No Purchase Orders linked to Quote ID: {quoteId}</Text>;
  }

  return (
    <Stack tokens={{ childrenGap: 8 }} style={{ paddingLeft: 48, marginTop: 12 }}>
      <Label style={{ fontWeight: "bold" }}>Linked Purchase Orders</Label>
      <DetailsList
        items={pos}
        columns={[
          { key: 'p1', name: 'PO ID', fieldName: 'POID', minWidth: 100 },
          { key: 'p2', name: 'Amount', fieldName: 'POAmount', minWidth: 100 },
          { key: 'p3', name: 'Quote Ref', fieldName: 'POQuoteID', minWidth: 120 },
          { key: 'p4', name: 'Comments', fieldName: 'POComments', minWidth: 150 },
        ]}
        selectionMode={0}
      />
    </Stack>
  );
};


  return (
    <div style={{ width: "100vh", height: "100vh" }} id="sales-webpart-root">
    <Stack tokens={{ childrenGap: 24 }}>
      <Text variant="xxLarge">Sales Dashboard</Text>

      <DetailsList
        items={opportunities}
        columns={[
          { key: 'oppid', name: 'Opportunity ID', fieldName: 'OpportunityID', minWidth: 100, isResizable: true },
        //   { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 150, isResizable: true },
        ]}
        selectionMode={0}
        onActiveItemChanged={(item: IUnifiedRecord) => {
          const newOpp = item.OpportunityID;
          setSelectedOppId(newOpp === selectedOppId ? null : newOpp);
          setExpandedQuoteId(null); // reset PO view
        }}
      />

      {selectedOppId && (
        <Stack tokens={{ childrenGap: 12 }}>
          <Text variant="large" style={{ marginTop: 12 }}>
            Details for Opportunity: <strong>{selectedOppId}</strong>
          </Text>
          {renderOpportunityDetails(selectedOppId)}
          {renderQuotationList(selectedOppId)}
        </Stack>
      )}
    </Stack>
    </div>
  );
};

export default SalesDashboard;
