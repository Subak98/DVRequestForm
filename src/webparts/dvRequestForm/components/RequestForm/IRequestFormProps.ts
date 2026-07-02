import { SPHttpClient } from "@microsoft/sp-http";
import { IServiceProvider } from "../../service/IServiceProvider";

export interface IRequestFormProps {
  description: string;
  currentWebUrl: string;
  spHttpClient: SPHttpClient;
  userDisplayName: string;
  onCancel?: () => void;
  onNavigateToTab?: (tab: "raise" | "review" | "myapproval") => void;
  provider: IServiceProvider;
  props: any; // Add this line to include the props object
}
