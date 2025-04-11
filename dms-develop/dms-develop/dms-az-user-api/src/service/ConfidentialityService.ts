import { logger, pnpjs } from "az-common/dist/index.mjs";
import { IRoleAssignmentInfo, PermissionKind } from "@pnp/sp/security/index.js";


export interface IConfidentialityService {
    checkUserPermissions(siteUrl: string, libraryId: string, itemId: number, user: string, permission: PermissionKind): Promise<boolean>;
    getEmailsFromPeopleColumn(siteUrl: string, listId: string, itemId: number, columnName: string): Promise<string[]>;
    getUserPermissionsOnSite(siteUrl: string, userLoginName: string): Promise<PermissionKind[]>;
    setItemPermissions(siteUrl: string, libraryId: string, itemId: number, permissions: { [user: string]: PermissionKind[] }): Promise<void>;
    removeUniquePermissions(siteUrl: string, libraryId: string, itemId: number): Promise<void>;
}
const LOG_SOURCE = "ConfidentialityService";
export class ConfidentialityService implements IConfidentialityService {

    public async getConfidentialUsers(siteUrl: string, libraryId: string, itemId: number): Promise<{ Id: number; Email: string; Title: string; }[]> {
        const sp = pnpjs.sp(siteUrl);
        // Check if the item has unique role assignments
        const items = await sp.web.lists.getById(libraryId).items.select("HasUniqueRoleAssignments").filter(`ID eq ${itemId}`)();
        if (!items || items.length === 0) {
            throw new Error(`Item with ID ${itemId} not found in library ${libraryId}`);
        }
        const item = items[0];
        if (!item.HasUniqueRoleAssignments) {
            return [];
        }
        const roleAssignments: IRoleAssignmentInfo & {
            Member: { Id: number; Title: string; Email: string; PrincipalType: number; }
        }[] = await sp.web.lists.getById(libraryId).items.getById(itemId).roleAssignments.select("Member/Id", "Member/Email", "Member/Title", "Member/PrincipalType").expand("Member")();
        return roleAssignments.filter(ra => ra.Member.PrincipalType === 1).map(ra => ra.Member);
    }

    public async checkUserPermissions(siteUrl: string, libraryId: string, itemId: number, userLoginName: string, permission: PermissionKind): Promise<boolean> {
        try {
            const user = await pnpjs.sp(siteUrl).web.ensureUser(userLoginName);
            const loginName = user.LoginName;
            const item = pnpjs.sp(siteUrl).web.lists.getById(libraryId).items.getById(itemId);
            const permissions = await item.getUserEffectivePermissions(loginName);
            return (permissions.High & permission) !== 0 || (permissions.Low & permission) !== 0;
        } catch (err) {
            logger.trackException({
                exception: err as Error,
                properties: { source: LOG_SOURCE, method: "checkUserPermissions" }
            });
            console.error("Error checking user permissions", err);
            throw new Error((err as Error).message);
        }
    }
    public async getEmailsFromPeopleColumn(siteUrl: string, listId: string, itemId: number, columnName: string): Promise<string[]> {
        try {
            // Log the request details for debugging
            logger.trackTrace({
                message: `Getting emails from People: ${JSON.stringify({ siteUrl, listId, itemId, columnName })}`,
                properties: {
                    source: LOG_SOURCE,
                    method: "getEmailsFromPeopleColumn"
                },
                severity: "Verbose"
            });
            console.log(`Getting emails from People column for siteUrl: ${siteUrl}, listId: ${listId}, itemId: ${itemId}, columnName: ${columnName}`);

            // Use $expand and $select to get user details
            const item = await pnpjs.sp(siteUrl).web.lists.getById(listId).items.getById(itemId).select(`${columnName}/EMail`).expand(columnName)();
            const users = item[columnName];
            const emails = users.map((user: any) => user.EMail); // Ensure to handle any type correctly
            return emails;
        } catch (err) {
            // Log the error details
            logger.trackException({
                exception: err as Error,
                properties: { source: LOG_SOURCE, method: "getEmailsFromPeopleColumn" }
            });
            console.error(`Error in getEmailsFromPeopleColumn:`, err);
            throw new Error(`Error getting emails from People column: ${err.message}`);
        }
    }

    public async getUserPermissionsOnSite(siteUrl: string, userLoginName: string): Promise<PermissionKind[]> {
        try {
            const user = await pnpjs.sp(siteUrl).web.ensureUser(userLoginName);
            const loginName = user.LoginName;
            const sitePermissions = await pnpjs.sp(siteUrl).web.getUserEffectivePermissions(loginName);
            const permissions: PermissionKind[] = [];

            if ((sitePermissions.High & PermissionKind.ViewListItems) !== 0 || (sitePermissions.Low & PermissionKind.ViewListItems) !== 0) {
                permissions.push(PermissionKind.ViewListItems);
            }
            if ((sitePermissions.High & PermissionKind.EditListItems) !== 0 || (sitePermissions.Low & PermissionKind.EditListItems) !== 0) {
                permissions.push(PermissionKind.EditListItems);
            }
            // Add other permissions as needed

            return permissions;
        } catch (err) {
            logger.trackException({
                exception: err as Error,
                properties: { source: LOG_SOURCE, method: "getUserPermissionsOnSite" }
            });
            console.error("Error getting user permissions on site", err);
            throw new Error((err as Error).message);
        }
    }

    public async setItemPermissions(siteUrl: string, libraryId: string, itemId: number, permissions: { [user: string]: PermissionKind[] }): Promise<void> {
        try {
            const sp = pnpjs.sp(siteUrl);
            const item = sp.web.lists.getById(libraryId).items.getById(itemId);

            // Break role inheritance
            await item.breakRoleInheritance(false);

            // Retrieve and log available role definitions for the site
            const roleDefinitions = await sp.web.roleDefinitions();
            logger.trackTrace({
                message: `Available role definitions ${JSON.stringify(roleDefinitions || {})}`,
                properties: {
                    source: LOG_SOURCE,
                    method: "setItemPermissions"
                },
                severity: "Verbose"
            });
            console.log(`Available role definitions:`, roleDefinitions);

            // Create a mapping of PermissionKind to role definition name
            const permissionToRoleDefMap: { [key: number]: string } = {
                [PermissionKind.ViewListItems]: "Read",
                [PermissionKind.EditListItems]: "Edit",
                [PermissionKind.FullMask]: "Full Control",
                // [PermissionKind.Contribute]: "Contribute"
                // Add other mappings as necessary
            };

            const keepUsers: string[] = [];
            const addedUsers: string[] = [];

            for (const [userLoginName, perms] of Object.entries(permissions)) {
                const user = await sp.web.ensureUser(userLoginName);
                const userId = user.Id;
                keepUsers.push(user.LoginName);

                // Log the userId and permissions
                logger.trackTrace({
                    message: `Adding role assignments ${JSON.stringify({ userLoginName, userId, perms })}`,
                    properties: {
                        source: LOG_SOURCE,
                        method: "setItemPermissions"
                    },
                    severity: "Verbose"
                });

                for (const perm of perms) {
                    const roleDefName = permissionToRoleDefMap[perm];
                    if (!roleDefName) {
                        logger.trackTrace({
                            message: `Invalid permission ${perm} for userId: ${userId}`,
                            properties: {
                                source: LOG_SOURCE,
                                method: "setItemPermissions"
                            },
                            severity: "Warning"
                        });
                        console.warn(`Invalid permission for userId: ${userId}, permission: ${perm}`);
                        continue;
                    }

                    // Find the role definition ID by name
                    const roleDefinition = roleDefinitions.find(rd => rd.Name === roleDefName);
                    if (!roleDefinition) {
                        logger.trackTrace({
                            message: `Role definition not found ${roleDefName} for userId: ${userId}`,
                            properties: {
                                source: LOG_SOURCE,
                                method: "setItemPermissions"
                            },
                            severity: "Warning"
                        });
                        console.warn(`Role definition not found for name: ${roleDefName}`);
                        continue;
                    }

                    try {
                        await item.roleAssignments.add(userId, roleDefinition.Id);
                        addedUsers.push(userLoginName);
                        logger.trackTrace({
                            message: `Added role assignment ${roleDefName} for userId: ${userId}/${userLoginName}`,
                            properties: {
                                source: LOG_SOURCE,
                                method: "setItemPermissions"
                            },
                            severity: "Verbose"
                        });
                    } catch (addErr) {
                        logger.trackException({
                            exception: new Error(addErr),
                            properties: { source: LOG_SOURCE, method: "addRoleAssignment" }
                        });
                        console.error(`Error adding role assignment for userId: ${userId}, permission: ${perm}`, addErr);
                    }
                }
            }

            // Remove other permissions
            const removedUsers = await this.removeOtherPermissions(siteUrl, libraryId, itemId, keepUsers);

            logger.trackTrace({
                message: `Permissions update summary ${JSON.stringify({ addedUsers, removedUsers })}`,
                properties: {
                    source: LOG_SOURCE,
                    method: "setItemPermissionsAndRemoveOthers"
                },
                severity: "Verbose"
            });

            console.log(`Added users: ${addedUsers.join(', ')}`);
            console.log(`Removed users: ${removedUsers.join(', ')}`);

        } catch (err) {
            logger.trackException({
                exception: err as Error,
                properties: { source: LOG_SOURCE, method: "setItemPermissionsAndRemoveOthers" }
            });
            console.error(`Error in setItemPermissionsAndRemoveOthers:`, err);
            throw new Error(`Error setting item permissions and removing others: ${err.message}`);
        }
    }

    public async removeOtherPermissions(siteUrl: string, libraryId: string, itemId: number, keepUsers: string[]): Promise<string[]> {
        const removedUsers: string[] = [];
        try {
            const sp = pnpjs.sp(siteUrl);
            const item = sp.web.lists.getById(libraryId).items.getById(itemId);

            // Get all role assignments for the item
            const roleAssignments = await item.roleAssignments();
            logger.trackTrace({
                message: `Current role assignments ${JSON.stringify(roleAssignments || {})}`,
                properties: {
                    source: LOG_SOURCE,
                    method: "removeOtherPermissions"
                },
                severity: "Verbose"
            });

            for (const roleAssignment of roleAssignments) {
                const principal = await sp.web.siteUsers.getById(roleAssignment.PrincipalId)();
                if (!keepUsers.includes(principal.LoginName)) {
                    try {
                        const roles = await item.roleAssignments.getById(roleAssignment.PrincipalId).bindings();
                        for (const role of roles) {
                            await item.roleAssignments.remove(roleAssignment.PrincipalId, role.Id);
                        }
                        removedUsers.push(principal.LoginName);
                        logger.trackTrace({
                            message: `Removed permissions for user ${JSON.stringify(principal || {})}`,
                            properties: {
                                source: LOG_SOURCE,
                                method: "removeOtherPermissions"
                            },
                            severity: "Verbose"
                        });
                    } catch (removeErr) {
                        logger.trackException({
                            exception: new Error(removeErr),
                            properties: { source: LOG_SOURCE, method: "removeRoleAssignment" }
                        });
                        console.error(`Error removing role assignment for userId: ${principal.Id}, userLoginName: ${principal.LoginName}`, removeErr);
                    }
                }
            }
        } catch (err) {
            logger.trackException({
                exception: err as Error,
                properties: { source: LOG_SOURCE, method: "removeOtherPermissions" }
            });
            console.error(`Error in removeOtherPermissions:`, err);
            throw new Error(`Error removing other permissions: ${err.message}`);
        }

        return removedUsers;
    }


    public async removeUniquePermissions(siteUrl: string, libraryId: string, itemId: number): Promise<void> {
        try {
            const item = pnpjs.sp(siteUrl).web.lists.getById(libraryId).items.getById(itemId);
            await item.resetRoleInheritance();
        } catch (err) {
            logger.trackException({
                exception: err as Error,
                properties: { source: LOG_SOURCE, method: "removeUniquePermissions" }
            });
            console.error("Error removing unique permissions", err);
            throw new Error((err as Error).message);
        }
    }
}
export const confidentialityService = new ConfidentialityService();