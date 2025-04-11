import { SPFI } from "@pnp/sp";
import { ISPUtilityService } from "./ISPUtilityService";
import { IListInfo } from "@pnp/sp/lists";
import axios, { AxiosInstance, AxiosRequestConfig } from "axios";
import { AppSettings } from "../AppSettings.js";
import { PnPClientStorage, getHashCode } from "@pnp/core";
import { IAzureIdentityToken } from "@pnp/azidjsclient";
import { AccessToken, DefaultAzureCredential, TokenCredential } from "@azure/identity";
import { logger } from "./Logger.js";
import "@pnp/sp/sites/index.js";
import "@pnp/sp/webs/index.js";
import "@pnp/sp/lists/index.js";
import "@pnp/sp/items/index.js";
import "@pnp/sp/fields/index.js";
import "@pnp/sp/site-users/web.js";
import "@pnp/sp/batching.js";
import "@pnp/queryable/behaviors/caching.js";
import { authHelper } from "./AuthHelper";
import { DmsRole } from "../model";
import { PermissionKind } from "@pnp/sp/security/index.js";

const LOG_SOURCE = "SPUtilityService";
export class SPUtilityService implements ISPUtilityService {
    private _httpClient: AxiosInstance;
    constructor(private _sp: SPFI) {
        this._httpClient = axios.create();
        this._httpClient.interceptors.request.use(this._addAuthHeader.bind(this));
    }
    private async _addAuthHeader(config: AxiosRequestConfig): Promise<AxiosRequestConfig> {
        const credential = await authHelper.getCredential();
        const scopes = [`https://${AppSettings.TenantName}.sharepoint.com/.default`];
        const key = `AzureIdentityCredential${Math.abs(getHashCode(scopes.join()))}`;
        const storage = new PnPClientStorage();

        let token: string;
        let expires: Date;

        const tokenStore = storage.session.get(key);
        if (tokenStore) {
            const storedToken: IAzureIdentityToken = JSON.parse(tokenStore);
            if (new Date(storedToken.expires) > (new Date())) {
                token = storedToken.token;
            }
        }
        if (token == null) {
            logger.trackTrace({
                message: 'Renewing token for http client',
                severity: "Verbose",
                properties: { source: LOG_SOURCE, method: '_addAuthHeader' }
            });
            const aiToken: AccessToken = await credential.getToken(scopes);
            // Set expiration date equal to the expiration timestamp minus 5 minutes for buffer
            expires = new Date((new Date()).getMilliseconds() + (aiToken.expiresOnTimestamp - 300000));
            token = aiToken.token;
            const newToken: IAzureIdentityToken = { token, expires };
            storage.session.put(key, JSON.stringify(newToken));
        }

        config.headers.Authorization = `Bearer ${token}`;
        return config;
    }

    public async getListFromName(listName: string, select?: string, expand?: string, isDocLib?: boolean): Promise<IListInfo> {
        logger.trackTrace({
            message: `Getting list by nam. Parameters: ${listName}, ${select}, ${expand}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE, method: 'getListFromName' }
        });
        const webUrl = (await this._sp.web()).ServerRelativeUrl;
        const url = !isDocLib ? `${webUrl}/Lists/${listName}` : `${webUrl}/${listName}`;
        const lists = await this._sp.web.lists
            .filter(`RootFolder/ServerRelativeUrl eq '${url}'`)
            .select(select)
            .expand(expand)();
        if (lists.length === 0) {
            return null;
        }
        return lists[0];
    }
    public async setPropertypagValues(values: { [propertyName: string]: string }): Promise<void> {
        logger.trackTrace({
            message: 'Setting property bag values: ' + JSON.stringify(values)?.replace(/\n/g, "'"),
            severity: "Verbose",
            properties: { source: LOG_SOURCE, method: 'setPropertypagValues' }
        });
        const client = this._httpClient;
        const site = await this._sp.site.select('Id')();
        const web = await this._sp.web.select('Id', 'Url')();
        const methods = Object.entries(values).map(([propertyName, propertyValue], index) => (
            `<Method Name="SetFieldValue" Id="${index + 4}" ObjectPathId="3">
                    <Parameters>
                        <Parameter Type="String">${propertyName}</Parameter>
                        <Parameter Type="String">${propertyValue}</Parameter>
                    </Parameters>
                </Method>`
        )).join('');
        const body = `
            <Request AddExpandoFieldTypeSuffix="true" SchemaVersion = "15.0.0.0" LibraryVersion = "15.0.0.0" ApplicationName = ".NET Library" xmlns = "http://schemas.microsoft.com/sharepoint/clientquery/2009" >
                <Actions>
                    ${methods}
                    <Method Name="Update" Id="1" ObjectPathId ="2" />
                </Actions>
                <ObjectPaths>
                    <Property Id="3" ParentId="2" Name="AllProperties" />
                    <Identity Id="2" Name = "85525e9f-f0af-b000-5242-902ff09fcdba|740c6a0b-85e2-48a0-a494-e0f1759d4aa7:site:${site.Id}:web:${web.Id}" />
                </ObjectPaths>
            </Request >`;
        const response = await client.post(`${web.Url}/_vti_bin/client.svc/ProcessQuery`, body, {
            headers: {
                "content-type": "text/xml"
            }
        });
        if (response.status >= 400) {
            throw new Error(response.statusText);
        } else {
            const responseObj: {
                SchemaVersion: string;
                LibraryVersion: string;
                ErrorInfo?: {
                    "ErrorMessage": string;
                    "ErrorValue": string;
                    "TraceCorrelationId": string;
                    "ErrorCode": number;
                    "ErrorTypeName": string;
                };
                TraceCorrelationId: string;
            }[] = response.data;
            if (responseObj.map(r => r.ErrorInfo).filter(e => e).length > 0) {
                throw new Error(responseObj.map(r => r.ErrorInfo).filter(e => e).map(e => e?.ErrorMessage).join('\n'));
            }
        }
    }

    public async getDmsRoles(email: string): Promise<DmsRole[]> {
        logger.trackTrace({
            message: 'Getting DMS roles: ' + email,
            severity: "Verbose",
            properties: { source: LOG_SOURCE, method: 'getDmsRoles' }
        });
        const user = await this._sp.web.ensureUser(email);
        const list = await this.getListFromName(AppSettings.CustomPermissionListName, 'Id');
        if (!list) {
            throw new Error(`List ${AppSettings.CustomPermissionListName} not found`);
        }
        const items = await this._sp.web.lists.getById(list.Id).items.select('Title', 'Id')();
        const roles = await Promise.all(
            items.map(i => {
                return this._sp.web.lists.getById(list.Id).items.getById(i.Id)
                    .userHasPermissions(user.LoginName, PermissionKind.ViewListItems)
                    .then(hasPermission => ({ role: i.Title, hasPermission }));
            })
        );
        return roles.filter(r => r.hasPermission).map(r => r.role as DmsRole);
    }
    public async checkWebPermission(permission: PermissionKind, email: string): Promise<boolean> {
        logger.trackTrace({
            message: `Checking web permission: ${permission}, user: ${email}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE, method: 'checkWebPermission' }
        });
        const user = await this._sp.web.ensureUser(email);
        return this._sp.web.userHasPermissions(user.LoginName, permission);
    }
    public async canCallFnApp(user: string, siteUrl: string): Promise<boolean> {
        logger.trackTrace({
            message: `Checking if user can call function app: ${user}, ${siteUrl}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE, method: 'canCallFnApp' }
        });
        const allowedRoles = AppSettings.AllowedDmsRoles;
        if (!allowedRoles || allowedRoles.length === 0) {
            return true;
        }
        const roles = await this.getDmsRoles(user);
        return roles.some(r => allowedRoles.includes(r));
    }
    public async moveDocument(libraryId: string, itemId: number, targetFolder: string, rootFolder: string): Promise<void> {
        logger.trackTrace({
            message: `Target Folder ${targetFolder}`,
            properties: { source: LOG_SOURCE },
            severity: 'Information'
        });
        const itemFolder = await this._sp.web.lists.getById(libraryId).items.getById(itemId).folder.select('ServerRelativeUrl')();
        const currentFolder = itemFolder.ServerRelativeUrl.substring(0, itemFolder.ServerRelativeUrl.lastIndexOf('/'));
        const currentName = itemFolder.ServerRelativeUrl.substring(itemFolder.ServerRelativeUrl.lastIndexOf('/') + 1);
        if (currentFolder === targetFolder) {
            logger.trackTrace({
                message: `Already in target folder`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
            return;
        }
        const spTargetFolder = await this._sp.web.getFolderByServerRelativePath(targetFolder).select('Exists')();
        logger.trackTrace({
            message: `Target Folder Exists: ${spTargetFolder.Exists}`,
            properties: { source: LOG_SOURCE },
            severity: 'Information'
        });
        if (!spTargetFolder.Exists) {
            await this._ensureFolder(this._sp, rootFolder, targetFolder);
            logger.trackTrace({
                message: `Target Folder Created`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
        }
        await this._sp.web.lists.getById(libraryId).items.getById(itemId).folder.moveByPath(`${targetFolder}/${currentName}`, {
            RetainEditorAndModifiedOnMove: true,
            ShouldBypassSharedLocks: true
        });
    }
    
    private async _ensureFolder(sp: SPFI, rootFolder: string, targetFolder: string): Promise<void> {
        const pathComponents = targetFolder.replace(rootFolder, '').split('/').filter(p => !!p);
        let path = rootFolder;
        while (pathComponents.length > 0) {
            path += `/${pathComponents.splice(0, 1)[0]}`;
            let spFolder = await sp.web.getFolderByServerRelativePath(path).select('Exists')();
            if (!spFolder.Exists) {
                try {
                    await sp.web.folders.addUsingPath(path);
                } catch (error) {
                    if (error.message.includes('-2130575257')) {
                        spFolder = await sp.web.getFolderByServerRelativePath(path).select('Exists')();
                        if (!spFolder.Exists) throw new Error(error)
                    }
                }
                logger.trackTrace({
                    message: `Folder Created: ${path}`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
            }
        }
    }
}