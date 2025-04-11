import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists/web";
import { IListInfo } from "@pnp/sp/lists";
import { CustomPermissionListName } from "../constants";
import { getSPFromPageContext } from "../pnpjs-config";

export enum DmsRole {
    DocumentController = 'Document Controller',
    Contributor = 'Contributor',
    Viewer = 'Viewer',
    ProjectAdmin = 'Project Administrator'
}
export interface IPermissionsService {
    getCurrentUserRoles(): Promise<DmsRole[]>;
    currentUserHasRole(role: DmsRole): Promise<boolean>;
}
export class PermissionsService implements IPermissionsService {

    public static readonly serviceKey: ServiceKey<IPermissionsService> =
        ServiceKey.create<IPermissionsService>('dms:IPermissionsService', PermissionsService);
    private _sp: SPFI;
    private _context: PageContext;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._context = serviceScope.consume(PageContext.serviceKey);

            this._sp = getSPFromPageContext(this._context);
        });
    }
    public async getCurrentUserRoles(): Promise<DmsRole[]> {
        const list = await this._getListFromName(CustomPermissionListName, 'Id');
        if (!list) {
            throw new Error(`List ${CustomPermissionListName} not found`);
        }
        const items = await this._sp.web.lists.getById(list.Id).items.select('Title', 'Id')();
        return items.map(item => item.Title as DmsRole);
    }
    public async currentUserHasRole(role: DmsRole): Promise<boolean> {
        return (await this.getCurrentUserRoles()).includes(role);
    }

    private async _getListFromName(listName: string, select?: string, expand?: string, isDocLib?: boolean): Promise<IListInfo | undefined> {
        const webUrl = (await this._sp.web()).ServerRelativeUrl;
        const url = !isDocLib ? `${webUrl}/Lists/${listName}` : `${webUrl}/${listName}`;
        let listRequest = this._sp.web.lists.filter(`RootFolder/ServerRelativeUrl eq '${url}'`);
        if (select) {
            listRequest = listRequest.select(select);
        }
        if (expand) {
            listRequest = listRequest.expand(expand);
        }
        const lists = await listRequest();
        if (lists.length === 0) {
            return undefined;
        }
        return lists[0];
    }
}