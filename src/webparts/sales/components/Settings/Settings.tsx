/* eslint-disable no-sequences */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable dot-notation */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable react/jsx-key */
import * as React from "react";
// import styles from '../Sales.module.scss';
import { ISettingsProps } from "./ISettingsProps";
import { ISettingsState, pageSettings } from "./ISettingsState";
// import { spfi, SPFI, SPFx } from '@pnp/sp';
import { DefaultButton, PrimaryButton, Dialog, DialogFooter, DialogType, Dropdown, FontIcon, IconButton, Pivot, PivotItem, Toggle, TextField, Panel, PanelType, Stack, Checkbox } from "@fluentui/react";
import "@pnp/sp/webs";
import "@pnp/sp/context-info";
import "@pnp/sp/profiles";
import "@pnp/sp/sites";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items/list";
import "@pnp/sp/sputilities";
import "@pnp/sp/regional-settings/web";
 import { TaskContext } from "../Sales";
import { SPFI, } from "@pnp/sp";
import { ISettings } from "../Sales";
// import { ISettings, ROLE } from '../../Types/GlobalTypes';
// import CustomToolTip from '../../CommonComponents/CustomToolTip/CustomToolTip';

export class Settings extends React.Component<ISettingsProps, ISettingsState> {
private onSettingsChange: any;
private settingsItem: any;
private settingsContext = TaskContext;
  private sp: SPFI;
  private settings: ISettings;
  static contextType = TaskContext;
  contextObject: any;
  constructor(props: ISettingsProps) {
    super(props);
    //this.sp = spfi().using(SPFx(this.settingsContext._currentValue.context));
    console.log("Settings constructor", this.settingsContext);
    // this.currentUserRole = this.settingsContext._currentValue.user.role;
console.log("Settings component mounted",this.settings);
     this.settings = this.settingsContext._currentValue.config.settings;
    this.settingsItem = this.settingsContext._currentValue.config.item;
    this.onSettingsChange = this.settingsContext._currentValue.config.onChange;

    this.state = {
      selectedKey: "General",

      hideCommandBar: false,
      hideCommentsWrapper: false,
      hidePageTitle: false,
      hideSideAppBar: false,
      hideSiteHeader: false,
      hideO365BrandNavbar: false,
      hideSharepointHubNavbar: false,
 isPOPrefixPanelOpen: false,
  poPrefix: "", // default
  qtnPrefix: "QTN",
  isQTNPrefixPanelOpen: false,
  oppPrefix: "OPP",
  isOPPPrefixPanelOpen: false,
      dateFormat: "MM/DD/YYYY", // default date format
      isDateFormatDialogOpen: false,
  currencySeparator: "International",
  isCurrencyDialogOpen: false,
    
    };
// this.sp = spfi().using(SPFx(this.props.context));

console.log("Settings constructor", this.props);
    this.setSettings = this.setSettings.bind(this);
    this.togglePageElements = this.togglePageElements.bind(this);
    this.updateConfiguration = this.updateConfiguration.bind(this);
    this.getURLParameters = this.getURLParameters.bind(this);
  }

  public async componentDidMount() {
    this.getURLParameters();
    this.setSettings();
    
    console.log("Settings component mounted",this.sp);

this.sp = this.props.context.sp as SPFI;

    this.settings = this.props.context.config.settings;
    this.settingsItem = this.props.context.config.item; 
    this.onSettingsChange = this.props.context.config.onChange;
    //  this.settingsItem = this.settingsContext._currentValue.config.item;
    //  this.onSettingsChange = this.settingsContext._currentValue.config.onChange;
console.log("Settings component mounted",this.props.context);
    window.addEventListener("popstate", this.getURLParameters);
    const poConfigItem = await this.sp.web.lists.getByTitle("CWSalesConfiguration").items
  .filter("Title eq 'POConfig'")
  .top(1)();

if (poConfigItem.length > 0) {
  const config = JSON.parse(poConfigItem[0].MultiValue || "{}");
  this.setState({ poPrefix: config.prefix || "PO" });
}
try {
  const qtnConfigItems = await this.sp.web.lists
    .getByTitle("CWSalesConfiguration")
    .items.filter("Title eq 'QTNConfig'")
    .top(1)();

  if (qtnConfigItems.length > 0) {
    const values = JSON.parse(qtnConfigItems[0].MultiValue || "{}");
    this.setState({ qtnPrefix: values.prefix || "QTN" });
  }
} catch (err) {
  console.warn("Failed to load QTN prefix from config:", err);
}
const oppConfigItems = await this.sp.web.lists
  .getByTitle("CWSalesConfiguration")
  .items.filter("Title eq 'OPPConfig'")
  .top(1)();

if (oppConfigItems.length > 0) {
  const config = JSON.parse(oppConfigItems[0].MultiValue || "{}");
  this.setState({ oppPrefix: config.prefix || "OPP" });
}
const dateConfigItems = await this.sp.web.lists
  .getByTitle("CWSalesConfiguration")
  .items.filter("Title eq 'DateFormat'")
  .top(1)();

if (dateConfigItems.length > 0) {
  const format = JSON.parse(dateConfigItems[0].MultiValue || "{}").format;
  this.setState({ dateFormat: format || "DD-MMM-YYYY" });
}
const currencyFormatItems = await this.sp.web.lists
  .getByTitle("CWSalesConfiguration")
  .items.filter("Title eq 'CurrencyFormat'")
  .top(1)();

if (currencyFormatItems.length > 0) {
  const format = JSON.parse(currencyFormatItems[0].MultiValue || "{}").format;
  this.setState({ currencySeparator: format || "International" });
}


  }

  public componentWillUnmount(): void {
    window.removeEventListener("popstate", this.getURLParameters);
  }

  private getURLParameters() {
    const hashValue = window.location.hash;

    if (hashValue) {
      const hash = hashValue.split("#")[1];
      this.setState(
        {
          selectedKey: ["General", "Page", "Navigation"].includes(hash)
            ? hash
            : "General",
        },
        () => {
          history.pushState({ page: "new" }, "", `#${this.state.selectedKey}`);
        }
      );
    } else {
      this.setState({ selectedKey: "General" }, () => {
        history.pushState({ page: "new" }, "", `#${this.state.selectedKey}`);
      });
    }
  }

  private setSettings() {
    const hideCommandBar = this.props.context.config.settings.hideCommandBar,
      hideSideAppBar = this.props.context.config.settings.hideSideAppBar,
      hidePageTitle = this.props.context.config.settings.hidePageTitle,
      hideSiteHeader = this.props.context.config.hideSiteHeader,
      hideCommentsWrapper = this.props.context.config.settings.hideCommentsWrapper,
      hideO365BrandNavbar = this.props.context.config.settings.hideO365BrandNavbar,
      hideSharepointHubNavbar = this.props.context.config.settings.hideSharepointHubNavbar,
      isHoursTrackable = this.props.context.config.settings.isHoursTrackable,
      roleBasedAccess = this.props.context.config.settings.roleBasedAccess


    this.setState({
        hideCommandBar,
        hideSideAppBar,
        hidePageTitle,
        hideSiteHeader,
        hideCommentsWrapper,
        hideO365BrandNavbar,
        hideSharepointHubNavbar,
        isHoursTrackable,
        roleBasedAccess,

    });
  }

  private togglePageElements(query: string, hide: boolean) {
    const HTMLElement: any = document.querySelector(`${query}`);
    if (HTMLElement) {
      HTMLElement.style.setProperty(
        "display",
        hide ? "none" : "block",
        "important"
      );
    }
  }
private saveOPPPrefixConfig = async () => {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'OPPConfig'")
      .top(1)();

    const payload = {
      Title: "OPPConfig",
      MultiValue: JSON.stringify({ prefix: this.state.oppPrefix || "OPP" }),
    };

    if (configItems.length > 0) {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.getById(configItems[0].Id).update(payload);
    } else {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.add(payload);
    }

    alert("Opportunity Prefix updated.");
    this.setState({ isOPPPrefixPanelOpen: false });
  } catch (err) {
    console.error("Failed to save OPP Prefix:", err);
    alert("Error saving OPP Prefix.");
  }
};
private saveQTNPrefixConfig = async () => {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'QTNConfig'")
      .top(1)();

    const payload = {
      Title: "QTNConfig",
      MultiValue: JSON.stringify({ prefix: this.state.qtnPrefix || "QTN" }),
    };

    if (configItems.length > 0) {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.getById(configItems[0].Id).update(payload);
    } else {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.add(payload);
    }

    alert("QTN Prefix updated.");
    this.setState({ isQTNPrefixPanelOpen: false });
  } catch (err) {
    console.error("Failed to save QTN Prefix:", err);
    alert("Error saving QTN Prefix.");
  }
};
private savePOPrefixConfig = async () => {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'POConfig'")
      .top(1)();

    const payload = {
      Title: "POConfig",
      MultiValue: JSON.stringify({ prefix: this.state.poPrefix || "PO" }),
    };

    if (configItems.length > 0) {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.getById(configItems[0].Id).update(payload);
    } else {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.add(payload);
    }

    alert("PO Prefix updated.");
    this.setState({ isPOPrefixPanelOpen: false });
  } catch (err) {
    console.error("Failed to save PO Prefix:", err);
    alert("Error saving PO Prefix.");
  }
};

  private async updateConfiguration() {
    await this.sp.web.lists.getByTitle("CWSalesConfiguration").items.getById(this.settingsItem?.Id).update({
      MultiValue: JSON.stringify(this.settings)
    }).then(() => {
      this.onSettingsChange(this.settings);
      console.log("Settings updated successfully", this.settings);
    });
     try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.select("Id", "Title")();

    const config = configItems.find((item) => item.Title === "Currency");

    if (config) {
      await this.sp.web.lists
        .getByTitle("CWSalesConfiguration")
        .items.getById(config.Id)
        .update({
          DefaultCurrency: this.state.defaultCurrency,
  
        });
    } else {
      await this.sp.web.lists
        .getByTitle("CWSalesConfiguration")
        .items.add({
          Title: "Currency",
          DefaultCurrency: this.state.defaultCurrency,

        });
    }

 
    this.setState({ isCurrencyPanelOpen: false });
  } catch (err) {
    console.error("Error saving currency config:", err);
    alert("Failed to update currency config.");
  }
    this.setState({ isDefaultCurrencyDialogOpen: false });

  }
  private saveDateFormatConfig = async () => {
  try {
    const items = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'DateFormat'")
      .top(1)();

    const payload = {
      Title: "DateFormat",
      MultiValue: JSON.stringify({ format: this.state.dateFormat || "DD-MMM-YYYY" }),
    };

    if (items.length > 0) {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.getById(items[0].Id).update(payload);
    } else {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.add(payload);
    }

    alert("Date format updated.");
    this.setState({ isDateFormatDialogOpen: false });
  } catch (err) {
    console.error("Error saving date format:", err);
    alert("Failed to save date format.");
  }
};

private async loadCurrencyOptionsFromConfig() {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.select("Title", "MultiValue")();

    const currencyConfig = configItems.find((item) => item.Title === "Currency");

    if (currencyConfig?.MultiValue) {
      const json = JSON.parse(currencyConfig.MultiValue);

      const rates = json.rates;
      const currencyOptions = Object.keys(rates).map((key) => ({
        key,
        text: key,
      }));

      this.setState({ currencyOptions });
    }
  } catch (err) {
    console.error("Failed to load currency config:", err);
  }
}
private saveCurrencySeparatorFormat = async () => {
  try {
    const configItems = await this.sp.web.lists
      .getByTitle("CWSalesConfiguration")
      .items.filter("Title eq 'CurrencyFormat'")
      .top(1)();

    const payload = {
      Title: "CurrencyFormat",
      MultiValue: JSON.stringify({ format: this.state.currencySeparator }),
    };

    if (configItems.length > 0) {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.getById(configItems[0].Id).update(payload);
    } else {
      await this.sp.web.lists.getByTitle("CWSalesConfiguration")
        .items.add(payload);
    }

    alert("Currency separator format saved.");
    this.setState({ isCurrencyDialogOpen: false });
  } catch (err) {
    console.error("Failed to save currency format:", err);
    alert("Error saving currency format.");
  }
};


  render() {
    // if (this.currentUserRole !== ROLE.ADMIN) {
    //   return <div><p style={{ textAlign: "center", fontSize: "2em" }}>You don't have access to this page.</p></div>;
    // }

    return (
       <div style={{ width: "100%", height: "100vh" }} id="sales-webpart-root">
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: "0.75em",
            marginLeft: "0.75em",
            fontSize: "1.25rem",
          }}
        >
          <FontIcon aria-label="settings" iconName="Settings" />
          <p style={{ margin: "0.5em 0" }}>Settings</p>
        </div>
        <Pivot
          aria-label="Settings"
          selectedKey={this.state.selectedKey}
          style={{ alignSelf: "flex-start" }}
          onLinkClick={(item) => {
              if (item && item.props.itemKey !== this.state.selectedKey) {
                const selectedKey = item.props.itemKey ?? "General";
                this.setState({ selectedKey }, () => {
                  history.pushState({ page: "new" }, "", `#${selectedKey}`);
                });
              }
          }}
        >
          {/* <PivotItem itemKey='General' headerText='General' itemIcon='CaseSetting'>
          <div style={{ paddingLeft: "0.75em" }}>
            {
              ...generalSettings.map(item => {
                return <Toggle
                  inlineLabel
                  onText='on'
                  offText='off'
                //   label={<div style={{ display: "flex", alignItems: "center", gap: "0.25em" }}><label>{item.label}</label>{<CustomToolTip id={item.stateVariable} tooltip={item.tooltip} />}</div>}
                  checked={this.state[item.stateVariable]}
                  styles={{ root: { justifyContent: "space-between" } }}
                //   onChange={(_, checked) => {
                //     this.setState({ [item.stateVariable]: checked }, async () => {
                //       this.settings[item.stateVariable] = checked;
                //       await this.updateConfiguration();
                //     });
                //   }}
                />;
              })
            }
          </div>
        </PivotItem> */}
          <PivotItem
            itemKey="Page"
            headerText="Page"
            itemIcon="TextDocumentSettings"
          >
            <div style={{ paddingLeft: "0.75em" }}>
              {...pageSettings.map((item) => {
                return (
                  <Toggle
                    inlineLabel
                    onText="Hide"
                    offText="Show"
                    label={
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: "0.25em",
                        }}
                      >
                        <label>{item.label}</label>
                        {/* {<CustomToolTip id={item.stateVariable} tooltip={item.tooltip} />} */}
                      </div>
                    }
                    checked={this.state[item.stateVariable]}
                    styles={{
                      root: { justifyContent: "space-between" },
                      text: { width: "50px" },
                    }}
                    onChange={(_, checked) => {
                      const isChecked = checked ?? false;
                      this.setState({ [item.stateVariable]: isChecked }, async () => {
                        this.settings[item.stateVariable] = isChecked;
                        await this.updateConfiguration()
                          .then(() => {
                            this.togglePageElements(item.sharepointElement, isChecked);
                          });
                      });
                    }}
                  />
                );
              })}
            </div>
          </PivotItem>
          <PivotItem itemKey="General" headerText="General" itemIcon="GlobalNavButton">
            <div style={{ marginTop: 16 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <label><b>Default Currency:</b> {this.state.defaultCurrency}</label>
                <IconButton
                  iconProps={{ iconName: "Edit" }}
                  title="Edit Default Currency"
                  ariaLabel="Edit Default Currency"
                  onClick={() => {
                    this.loadCurrencyOptionsFromConfig();
                    this.setState({ isDefaultCurrencyDialogOpen: true });
                  }}
                />
              </div>
            </div>
            <div style={{ marginTop: 16 }}>
  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
    <label><b>Currency Separator:</b> {this.state.currencySeparator}</label>
    <IconButton
      iconProps={{ iconName: "Edit" }}
      title="Edit Currency Format"
      onClick={() => this.setState({ isCurrencyDialogOpen: true })}
    />
  </div>
</div>

            <div style={{ marginTop: 16 }}>
  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
    <label><b>PO ID Prefix:</b> {this.state.poPrefix}</label>
    <IconButton
      iconProps={{ iconName: "Edit" }}
      title="Edit PO Prefix"
      ariaLabel="Edit PO Prefix"
      onClick={() => this.setState({ isPOPrefixPanelOpen: true })}
    />
  </div>
</div>
<Dialog
  hidden={!this.state.isCurrencyDialogOpen}
  onDismiss={() => this.setState({ isCurrencyDialogOpen: false })}
  dialogContentProps={{
    type: DialogType.normal,
    title: "Currency Separator",
  }}
>
  <Stack horizontal tokens={{ childrenGap: 12 }}>
    {["International", "India", "None"].map((format) => (
      <Checkbox
        key={format}
        label={format === "None" ? "No separator" : format}
        checked={this.state.currencySeparator === format}
        onChange={() => this.setState({ currencySeparator: format as any })}
      />
    ))}
  </Stack>
  <DialogFooter>
    <PrimaryButton text="Save" onClick={this.saveCurrencySeparatorFormat} />
    <DefaultButton text="Cancel" onClick={() => this.setState({ isCurrencyDialogOpen: false })} />
  </DialogFooter>
</Dialog>

<div style={{ marginTop: 16 }}>
  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
    <label><b>QTN ID Prefix:</b> {this.state.qtnPrefix}</label>
    <IconButton
      iconProps={{ iconName: "Edit" }}
      title="Edit QTN Prefix"
      ariaLabel="Edit QTN Prefix"
      onClick={() => this.setState({ isQTNPrefixPanelOpen: true })}
    />
  </div>
</div>
<div style={{ marginTop: 16 }}>
  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
    <label><b>Opportunity ID Prefix:</b> {this.state.oppPrefix}</label>
    <IconButton
      iconProps={{ iconName: "Edit" }}
      title="Edit Opportunity Prefix"
      ariaLabel="Edit Opportunity Prefix"
      onClick={() => this.setState({ isOPPPrefixPanelOpen: true })}
    />
  </div>
</div>
<div style={{ marginTop: 16 }}>
  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
    <label><b>Date Format:</b> {this.state.dateFormat}</label>
    <IconButton
      iconProps={{ iconName: "Edit" }}
      title="Edit Date Format"
      ariaLabel="Edit Date Format"
      onClick={() => this.setState({ isDateFormatDialogOpen: true })}
    />
  </div>
</div>


          </PivotItem>
          {/* <PivotItem itemKey='Navigation' headerText='Navigation' itemIcon='SecondaryNav'>
          <div>
            <div>
              <Dropdown
                label='Planner Visible To:'
                placeholder='Select roles'
                multiSelect
                selectedKeys={this.state['plannersVisibleTo']}
                options={[
                  { key: "Leader", text: "Leader" },
                  { key: "Member", text: "Member" },
                  { key: "Guest", text: "Guest" }
                ]}
                onChange={(_, opt) => {
                  if (opt) {
                    let plannersVisibleTo = [...this.state['plannersVisibleTo']];

                    if (opt.selected) {
                      plannersVisibleTo.push(opt.key);
                    } else {
                      plannersVisibleTo = plannersVisibleTo.filter(key => key !== opt.key);
                    }

                    this.setState({ plannersVisibleTo }, () => {
                    //   this.settings['plannersVisibleTo'] = plannersVisibleTo;
                    //   this.updateConfiguration();
                    //   this.onSettingsChange(this.settings);
                    });
                  }
                }}
              />
            </div>
          </div>
        </PivotItem> */}
        </Pivot>
        <Dialog
          hidden={!this.state.isDefaultCurrencyDialogOpen}
          onDismiss={() => this.setState({ isDefaultCurrencyDialogOpen: false })}
          dialogContentProps={{
            type: DialogType.normal,
            title: "Edit Default Currency",
            closeButtonAriaLabel: "Close"
          }}
        >

          <Dropdown
            label="Default Currency"
            selectedKey={this.state.defaultCurrency}
            options={this.state.currencyOptions}
            onChange={(_, option) => this.setState({ defaultCurrency: option?.key as string })}
          />
          <DialogFooter>
            <DefaultButton onClick={this.updateConfiguration} text="Save" />
            <DefaultButton onClick={() => this.setState({ isDefaultCurrencyDialogOpen: false })} text="Cancel" />
          </DialogFooter>
        </Dialog>
        <Dialog
  hidden={!this.state.isDateFormatDialogOpen}
  onDismiss={() => this.setState({ isDateFormatDialogOpen: false })}
  dialogContentProps={{
    type: DialogType.normal,
    title: "Change Date Format",
    subText: "Choose a format for displaying dates in the webpart.",
  }}
>
  <Dropdown
    label="Select Date Format"
    selectedKey={this.state.dateFormat}
    options={[
      { key: "DD-MMM-YYYY", text: "08-Jun-2025" },
      { key: "YYYY-MM-DD", text: "2025-06-08" },
      { key: "MM/DD/YYYY", text: "06/08/2025" },
      { key: "DD/MM/YYYY", text: "08/06/2025" },
    ]}
    onChange={(_, option) => this.setState({ dateFormat: option?.key as string })}
  />
  <DialogFooter>
    <PrimaryButton text="Save" onClick={this.saveDateFormatConfig} />
    <DefaultButton text="Cancel" onClick={() => this.setState({ isDateFormatDialogOpen: false })} />
  </DialogFooter>
</Dialog>

                  <Panel
  headerText="Edit PO Prefix"
  isOpen={this.state.isPOPrefixPanelOpen}
  onDismiss={() => this.setState({ isPOPrefixPanelOpen: false })}
  closeButtonAriaLabel="Close"
  isLightDismiss
  type={PanelType.smallFixedFar}
>
  <TextField
    label="PO Prefix"
    value={this.state.poPrefix}
    onChange={(_, val) => this.setState({ poPrefix: val || "" })}
  />

  <DialogFooter>
    <DefaultButton text="Save" onClick={this.savePOPrefixConfig} />
    <DefaultButton text="Cancel" onClick={() => this.setState({ isPOPrefixPanelOpen: false })} />
  </DialogFooter>
</Panel>
<Panel
  headerText="Edit QTN Prefix"
  isOpen={this.state.isQTNPrefixPanelOpen}
  onDismiss={() => this.setState({ isQTNPrefixPanelOpen: false })}
  closeButtonAriaLabel="Close"
  isLightDismiss
  type={PanelType.smallFixedFar}
>
  <TextField
    label="QTN Prefix"
    value={this.state.qtnPrefix}
    onChange={(_, val) => this.setState({ qtnPrefix: val || "" })}
  />
  <DialogFooter>
    <DefaultButton text="Save" onClick={this.saveQTNPrefixConfig} />
    <DefaultButton text="Cancel" onClick={() => this.setState({ isQTNPrefixPanelOpen: false })} />
  </DialogFooter>
</Panel>
<Panel
  headerText="Edit Opportunity Prefix"
  isOpen={this.state.isOPPPrefixPanelOpen}
  onDismiss={() => this.setState({ isOPPPrefixPanelOpen: false })}
  closeButtonAriaLabel="Close"
  isLightDismiss
  type={PanelType.smallFixedFar}
>
  <TextField
    label="Opportunity Prefix"
    value={this.state.oppPrefix}
    onChange={(_, val) => this.setState({ oppPrefix: val || "" })}
  />
  <DialogFooter>
    <DefaultButton text="Save" onClick={this.saveOPPPrefixConfig} />
    <DefaultButton text="Cancel" onClick={() => this.setState({ isOPPPrefixPanelOpen: false })} />
  </DialogFooter>
</Panel>

      </div>
    );
  }
}

export default Settings;
