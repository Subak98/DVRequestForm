import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IServiceProvider } from "../service/IServiceProvider";

export interface IDvRequestFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  currentWebUrl: string;
  spHttpClient: SPHttpClient;
  userEmail: string;
  context: WebPartContext;
  provider: IServiceProvider;
  RequestFormId: string;
  formDescription: string;
  PAURL: string;
}
