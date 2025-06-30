/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SalesWebPartStrings';
import Sales from './components/Sales';
import { ISalesProps } from './components/ISalesProps';

import { spfi, SPFI, SPFx } from "@pnp/sp";
// import OpportunityViewer from './components/ViewOpportunities/ViewOpportunies';
// import OpportunityViewer from './components/ViewOpportunities/ViewOpportunies';
export interface ISalesWebPartProps {
  description: string;
}

export default class SalesWebPart extends BaseClientSideWebPart<ISalesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp: SPFI;
  props: any;
  sp: SPFI;

  public componentDidMount(): void {
    this._sp = spfi().using(getSPFxContext(this.context));
    console.log(this._sp)
  }
  public render(): void {
    const element: React.ReactElement<ISalesProps> = React.createElement(
      Sales,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this.sp,
        context: this.context,
        View: "Opportunity",
        settings: {}, // Provide appropriate settings object
        onConfigChange: () => {}, // Provide appropriate callback function
       // isDeployedToMainSite: false, // Add appropriate value
       // ProjectsSite: "", // Add appropriate value
       // ProjectName: "", // Add appropriate value
        //ComponentDropdown: "", // Add appropriate value
        //TimeSite: "", // Add appropriate value
       // ComponentHeight: 0, // Add appropriate value
       // Planner: "", // Add appropriate value
       // siteOption: "", // Add appropriate value
        // Add other missing properties with appropriate values
      }
    );

    ReactDom.render(element, this.domElement);
    // const element2: React.ReactElement<{}> = React.createElement(
    //   OpportunityViewer
    // );

    // ReactDom.render(element2, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
      this.sp = spfi().using(SPFx(this.context));
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}


function getSPFxContext(context: any): import("@pnp/core").TimelinePipe {
  throw new Error('Function not implemented.');
}

