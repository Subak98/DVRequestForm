import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp, pnp logging system, and any other selective imports needed
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { ConsoleListener, LogLevel, Logger } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/security";
let _sp: SPFI = new SPFI();

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== null && context !== undefined) {
    // Configure logging with ConsoleListener
    // Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Warning;

    _sp = spfi().using(SPFx(context));
  }
  return _sp;
};
