import { SPHttpClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IServiceProvider } from "../../service/IServiceProvider";

export interface IApproveFormProps {
  currentWebUrl: string;
  spHttpClient: SPHttpClient;
  context: WebPartContext;
  provider: IServiceProvider;
  requestId: string;
  listid: string;
  useremail: string;
}
