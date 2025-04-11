import { WebPartContext } from "@microsoft/sp-webpart-base";
import { FormCustomizerContext, FieldCustomizerContext, ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
import { BrowserFetchWithRetry,CacheNever } from "@pnp/queryable";
import { PageContext } from "@microsoft/sp-page-context";

// import pnp, pnp logging system, and any other selective imports needed
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { graphfi, GraphFI, SPFx as gSPFx } from "@pnp/graph";
import "@pnp/sp/batching";
import "@pnp/logging";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";

import {TimelinePipe} from "@pnp/core";
import {Queryable,InjectHeaders} from "@pnp/queryable";
import { DefaultHeaders} from "@pnp/graph";

let _sp: SPFI;
let _graph: GraphFI;

export const getSP = (context?: WebPartContext | FormCustomizerContext | FieldCustomizerContext | ListViewCommandSetContext | ApplicationCustomizerContext): SPFI => {
  if (context) {
    _sp = spfi()
    .using(SPFx(context))
    .using(PnPLogging(LogLevel.Warning))
    .using(BrowserFetchWithRetry())
    .using(CahceNoStore());
  }
  return _sp;
};
export const getSPFromPageContext = (pageContext: PageContext): SPFI => {
  _sp = spfi()
  .using(SPFx({ pageContext: pageContext }))
  .using(PnPLogging(LogLevel.Warning))
  .using(BrowserFetchWithRetry())
  .using(CahceNoStore());
  return _sp;
};
export const getGraph = (context?: WebPartContext | FormCustomizerContext | FieldCustomizerContext | ListViewCommandSetContext): GraphFI => {
  if (context) {
    _graph = graphfi()
    .using(gSPFx(context))
    .using(PnPLogging(LogLevel.Warning))
  }
  return _graph;
}

export const getSPCacheNever = (context?: WebPartContext | FormCustomizerContext | FieldCustomizerContext | ListViewCommandSetContext | ApplicationCustomizerContext): SPFI => {
  if (context) {
    _sp = spfi()
    .using(SPFx(context))
    .using(PnPLogging(LogLevel.Warning))
    .using(BrowserFetchWithRetry())
    .using(CacheNever())
    .using(CahceNoStore());
  }
  return _sp;
};

export const getSPFromPageContextCacheNever = (pageContext: PageContext): SPFI => {
  _sp = spfi()
  .using(SPFx({ pageContext: pageContext }))
  .using(PnPLogging(LogLevel.Warning))
  .using(BrowserFetchWithRetry())
  .using(CacheNever())
  .using(CahceNoStore());
  return _sp;
};

export function CahceNoStore(): TimelinePipe<Queryable> {

  return (instance: Queryable) => {
      instance.using(
          // use the default headers
          DefaultHeaders(),          
          // inject our special header to all requests
          InjectHeaders({
              "Cache-Control": "no-store",
          }));
      return instance;
  };
}