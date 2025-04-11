import { SPFI } from "@pnp/sp";
import { DmsRole, IPermissionsService } from "./IPermissionsService";
import { AppSettings } from "../AppSettings.js";
import { SPUtilityService } from "az-common/dist/index.mjs";
import { logger } from "az-common/dist/index.mjs";
import { PermissionKind } from "@pnp/sp/security/index.js";

const LOG_SOURCE = "PermissionsService";
export class PermissionsService implements IPermissionsService {
    private _spUtility: SPUtilityService;

    constructor(private _sp: SPFI) {
        this._spUtility = new SPUtilityService(_sp);
    }

    public async getDmsRoles(email: string): Promise<DmsRole[]> {
        logger.trackTrace({
            message: 'Getting DMS roles: ' + email,
            severity: "Verbose",
            properties: { source: LOG_SOURCE, method: 'getDmsRoles' }
        });
        const user = await this._sp.web.ensureUser(email);
        const list = await this._spUtility.getListFromName(AppSettings.CustomPermissionListName, 'Id');
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

}