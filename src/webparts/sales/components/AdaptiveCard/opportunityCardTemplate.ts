export const opportunityCard = {
  type: "AdaptiveCard",
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
  version: "1.5",
  body: [
    { type: "TextBlock", text: "Opportunity Details", weight: "Bolder", size: "Medium" },
    { type: "Input.Text", id: "Title", label: "Title", isRequired: true },
    { type: "Input.Text", id: "Customer", label: "Customer" },
    { type: "Input.Text", id: "Business", label: "Business" },
    { type: "Input.Text", id: "OEM", label: "OEM" },
    { type: "Input.Text", id: "RiskLevel", label: "Risk Level" },
    { type: "Input.Text", id: "AmountEUR", label: "Amount EUR" },
    { type: "Input.Date", id: "TentativeStartDate", label: "Tentative Start Date" },
    { type: "Input.Text", id: "Comments", label: "Comments", isMultiline: true }
  ],
  actions: [
    { type: "Action.Submit", title: "Save Opportunity" }
  ]
};
