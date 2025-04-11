import { TokenCredential } from "@azure/identity";

import { spfi, SPFI } from "@pnp/sp/index.js";
import { AzureIdentity } from "@pnp/azidjsclient";
import { GraphDefault, SPDefault } from "@pnp/nodejs";

import { graphfi, GraphFI } from '@pnp/graph';
import { logger } from "./Logger.js";
import { AppSettings } from "../AppSettings.js";
import { PnPLogging } from "@pnp/logging";
import { Caching } from "@pnp/queryable/behaviors/caching.js";
import { authHelper } from "./AuthHelper.js";
import { TimelinePipe } from "@pnp/core/index.js";

export interface IPnpjsService {
  Init: (noCache?: boolean) => Promise<void>;
  sp(siteRelativeUrl: string): SPFI;
  graph(behaviors?:TimelinePipe[]): GraphFI;
}

export class PnpjsService implements IPnpjsService {
  private LOG_SOURCE = "PnpjsService";
  private _ready: boolean = false;

  private _sps: { [siteRelativeUrl: string]: SPFI } = {};
  private _credential: TokenCredential;
  private _noCache: boolean;

  public constructor() { }

  public async Init(noCache?: boolean): Promise<void> {
    try {
      this._credential = await authHelper.getCredential();
      this._noCache = noCache;

      const requiredAppSettings = !AppSettings.IsLocalEnvironment ?
        ['TenantName', 'TenantId', 'ClientId', 'CertificateName', 'KeyVaultUrl'] :
        ['TenantName', 'TenantId', 'ClientId', 'CertificatePath'];
      requiredAppSettings.forEach(setting => {
        if (!AppSettings[setting]) {
          throw new Error(`App setting ${setting} is required`);
        }
      });

      this._ready = true;
      logger.trackTrace({
        message: 'Init success',
        properties: {
          source: this.LOG_SOURCE,
          method: "Init"
        },
        severity: "Verbose"
      });
    } catch (err) {
      logger.trackException({
        exception: err,
        properties: { source: this.LOG_SOURCE, method: "Init" }
      });
      throw new Error(err);
    }
  }

  public get ready(): boolean {
    return this._ready;
  }

  public sp(siteRelativeUrl: string): SPFI {
    if (this._sps[siteRelativeUrl] == null) {
      this._sps[siteRelativeUrl] = this._noCache ?
        spfi(`https://${AppSettings.TenantName}.sharepoint.com${siteRelativeUrl}`).using(
          SPDefault({}),
          PnPLogging(AppSettings.PnPLogLevel),
          AzureIdentity(this._credential, [`https://${AppSettings.TenantName}.sharepoint.com/.default`], null)
        ) :
        spfi(`https://${AppSettings.TenantName}.sharepoint.com${siteRelativeUrl}`).using(
          SPDefault({}),
          AzureIdentity(this._credential, [`https://${AppSettings.TenantName}.sharepoint.com/.default`], null),
          PnPLogging(AppSettings.PnPLogLevel),
          Caching()
        );
    }
    return this._sps[siteRelativeUrl];
  }

  public graph(behaviors?:TimelinePipe[]): GraphFI {
    
    return !behaviors ? graphfi().using(
      GraphDefault({}),
      PnPLogging(AppSettings.PnPLogLevel),
      AzureIdentity(this._credential, [`https://graph.microsoft.com/.default`], null)
    ) :
    graphfi().using(
      GraphDefault({}),
      PnPLogging(AppSettings.PnPLogLevel),
      AzureIdentity(this._credential, [`https://graph.microsoft.com/.default`], null),
      ...behaviors
    );
  }

}

export const pnpjs: IPnpjsService = new PnpjsService();