import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneLink,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

import * as strings from "DvRequestFormWebPartStrings";
import DvRequestForm from "./components/DvRequestForm";
import { IDvRequestFormProps } from "./components/IDvRequestFormProps";
import { SPHttpClient } from "@microsoft/sp-http";
import { IServiceProvider } from "./service/IServiceProvider";
import { SPFI } from "@pnp/sp";
import { getSP } from "./service/pnpjsConfig";
import SharePointServiceProvider from "../dvRequestForm/service/SharepointServiceProvider";
import {
  PropertyFieldListPicker,
  PropertyFieldListPickerOrderBy,
} from "@pnp/spfx-property-controls";
import { Link } from "@fluentui/react";

export interface IDvRequestFormWebPartProps {
  description: string;
  RequestFormId: string;
  formDescription: string;
  PAURL: string;
}

export default class DvRequestFormWebPart extends BaseClientSideWebPart<IDvRequestFormWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";
  private _serviceProvider: IServiceProvider | any;
  private formdesID: any;
  public render(): void {
    const element: React.ReactElement<IDvRequestFormProps> =
      React.createElement(DvRequestForm, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email,
        currentWebUrl: this.context.pageContext.web.absoluteUrl,
        spHttpClient: this.context.spHttpClient as SPHttpClient,
        context: this.context,
        provider: this._serviceProvider,
        RequestFormId: this.properties.RequestFormId,
        formDescription: this.properties.formDescription,
        PAURL: this.properties.PAURL
      });

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._serviceProvider = new SharePointServiceProvider(this.context);
    getSP(this.context);
    this._serviceProvider
      .fetchFormDescription(this.properties.formDescription)
      .then((res: any) => {
        this.formdesID = res[0].ID;
      });
    return this._getEnvironmentMessage().then((message) => {
      this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app
        .getContext()
        .then((context) => {
          let environmentMessage: string = "";
          switch (context.app.host.name) {
            case "Office": // running in Office
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOffice
                : strings.AppOfficeEnvironment;
              break;
            case "Outlook": // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentOutlook
                : strings.AppOutlookEnvironment;
              break;
            case "Teams": // running in Teams
            case "TeamsModern":
              environmentMessage = this.context.isServedFromLocalhost
                ? strings.AppLocalEnvironmentTeams
                : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(
      this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentSharePoint
        : strings.AppSharePointEnvironment,
    );
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null,
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null,
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker("RequestFormId", {
                  label: "Select Request Form ID",
                  selectedList: this.properties.RequestFormId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyFieldListPicker("formDescription", {
                  label: "Select Form Description List",
                  selectedList: this.properties.formDescription,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                }),
                PropertyPaneLink("", {
                  target: "_blank",
                  href: `${this.context.pageContext.web.absoluteUrl}/_layouts/15/listform.aspx?PageType=6&ListId=${this.properties.formDescription}&ID=${this.formdesID}`,
                  text: "Edit Form Description",
                }),
                PropertyPaneTextField("PAURL", {
                  label: "Enter Trigger URL for Power Automate",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
