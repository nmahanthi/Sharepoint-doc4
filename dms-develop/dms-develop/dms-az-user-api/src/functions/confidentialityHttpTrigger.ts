import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger } from "az-common/dist/index.mjs";
import { pnpjs } from "az-common/dist/index.mjs";
import { PermissionKind } from "@pnp/sp/security/index.js";
import { confidentialityService } from "../service/ConfidentialityService.js";
import { IConfidentialityRequest } from "../model/IConfidentialityRequest.js";
import { AppSettings } from "../AppSettings.js";
import { folderFromServerRelativePath } from "@pnp/sp/folders/types.js";

const LOG_SOURCE = "confidentialityHttpTrigger";
async function handleGetMethod(request: IConfidentialityRequest, loginName: string): Promise<HttpResponseInit> {
    const { siteUrl, libraryId, itemId: reqItemId, folderUrl } = request;
    try {
        if (!siteUrl || !libraryId || !(reqItemId || folderUrl)) {
            const missingParams = {
                siteUrl,
                libraryId,
                reqItemId,
                folderUrl
            };
            logger.trackTrace({
                message: `Missing required parameters ${JSON.stringify(missingParams || {})}`,
                properties: {
                    source: LOG_SOURCE,
                    method: "confidentialityGetHttpTrigger"
                },
                severity: "Error"
            });
            console.error('Missing required parameters', missingParams);
            return { status: 400, jsonBody: { message: "Missing required parameters" } };
        }
        let itemId = reqItemId;
        if (!itemId) {
            logger.trackTrace({
                message: `Getting item id from folder url ${folderUrl}`,
                properties: { source: LOG_SOURCE, method: "confidentialityGetHttpTrigger", },
                severity: "Info"
            });
            const sp = pnpjs.sp(siteUrl);
            const folder = folderFromServerRelativePath(sp.web, folderUrl);
            const folderItem = await (await folder.getItem('Id'))();
            itemId = folderItem.Id;
        }
        const hasReadPermission = await confidentialityService.checkUserPermissions(siteUrl, libraryId, itemId, loginName, PermissionKind.ViewListItems);
        if (!hasReadPermission) {
            logger.trackTrace({
                message: "User does not have read permissions",
                properties: { source: LOG_SOURCE, method: "checkUserPermissions" },
                severity: "Warning"
            });
            console.warn("User does not have read permissions");
            return { status: 403, jsonBody: { message: "User does not have read permissions" } };
        }
        const userEmails = request.users || [];
        if (userEmails.length === 0) {
            logger.trackTrace({
                message: "No users are defined",
                properties: { source: LOG_SOURCE, method: "getEmailsFromPeopleColumn" },
                severity: "Warning"
            });
            console.warn("No users are defined");
        }
        const users: { Id: number; Email: string; Title: string; }[] = await confidentialityService.getConfidentialUsers(siteUrl, libraryId, itemId);
        return { status: 200, body: JSON.stringify(users) };
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE, method: "GET" }
        });
        const message = error instanceof Error ? error.message : 'Invalid request from confidentialityTrigger';
        return { status: 400, body: message };
    }
}
async function handlePostMethod(request: IConfidentialityRequest, loginName: string): Promise<HttpResponseInit> {
    const { siteUrl, libraryId, itemId: reqItemId, folderUrl, confidentiality } = request;
    try{
        if (!siteUrl || !libraryId || !(reqItemId || folderUrl) || !confidentiality) {
            const missingParams = {
                siteUrl,
                libraryId,
                reqItemId,
                folderUrl,
                confidentiality
            };
            logger.trackTrace({
                message: `Missing required parameters ${JSON.stringify(missingParams || {})}`,
                properties: {
                    source: LOG_SOURCE,
                    method: "confidentialityHttpTrigger"
                },
                severity: "Error"
            });
            console.error('Missing required parameters', missingParams);
            return { status: 400, jsonBody: { message: "Missing required parameters" } };
        }
        let itemId = reqItemId;
        if (!itemId) {
            logger.trackTrace({
                message: `Getting item id from folder url ${folderUrl}`,
                properties: { source: LOG_SOURCE, method: "confidentialityHttpTrigger", },
                severity: "Info"
            });
            const sp = pnpjs.sp(siteUrl);
            const folder = folderFromServerRelativePath(sp.web, folderUrl);
            const folderItem = await (await folder.getItem('Id'))();
            itemId = folderItem.Id;
        }

        const hasEditPermission = await confidentialityService.checkUserPermissions(siteUrl, libraryId, itemId, loginName, PermissionKind.EditListItems);
        if (!hasEditPermission) {
            logger.trackTrace({
                message: "User does not have edit permissions",
                properties: { source: LOG_SOURCE, method: "checkUserPermissions" },
                severity: "Warning"
            });
            console.warn("User does not have edit permissions");
            return { status: 403, jsonBody: { message: "User does not have edit permissions" } };
        }

        if (confidentiality === AppSettings.ConfidentialValue) {
            const userEmails = request.users || [];
            if (userEmails.length === 0) {
                logger.trackTrace({
                    message: "No users are defined",
                    properties: { source: LOG_SOURCE, method: "getEmailsFromPeopleColumn" },
                    severity: "Warning"
                });
                console.warn("No users are defined");
            }
            else {
                const newPermissions: { [user: string]: PermissionKind[] } = {};
                const usersWithoutPermissions: string[] = [];
                for (const email of userEmails) {
                    try {
                        const userPermissions = await confidentialityService.getUserPermissionsOnSite(siteUrl, email);
                        if (userPermissions.length > 0) {
                            newPermissions[email] = userPermissions;
                        } else {
                            usersWithoutPermissions.push(email);
                        }
                    } catch (err) {
                        logger.trackException({
                            exception: new Error(err),
                            properties: { source: LOG_SOURCE, method: "getUserPermissionsOnSite" }
                        });
                        console.error("Error getting user permissions on site", err);
                    }
                }

                if (usersWithoutPermissions.length > 0) {
                    logger.trackTrace({
                        message: `Users without site permissions: ${usersWithoutPermissions.join(', ')}`,
                        properties: { source: LOG_SOURCE, method: "getUserPermissionsOnSite" },
                        severity: "Warning"
                    });
                    console.warn(`Users without site permissions: ${usersWithoutPermissions.join(', ')}`);
                }

                await confidentialityService.setItemPermissions(siteUrl, libraryId, itemId, newPermissions);
            }
        }
        else {
            logger.trackTrace({
                message: "confidentiality is not confidential",
                properties: { source: LOG_SOURCE, method: "ConfidentialityCheck" },
                severity: "Info"
            });
            console.log("confidentiality is not confidential");
            // Remove unique permissions and inherit from site
            await confidentialityService.removeUniquePermissions(siteUrl, libraryId, itemId);
        }
        return { status: 200, body: 'Success' };
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE, method: "POST" }
        });
        const message = error instanceof Error ? error.message : 'Invalid request from confidentialityTrigger';
        return { status: 400, body: message };
    }
}

export async function confidentialityHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    try {
        let requestBody: IConfidentialityRequest;
        if (request.method === 'GET') {
            requestBody =  {
                siteUrl: request.query.get('siteUrl'),
                libraryId: request.query.get('libraryId'),
                itemId: +request.query.get('itemId')
            } as IConfidentialityRequest;
        } else {
            requestBody = await request.json() as IConfidentialityRequest;
        }
        logger.trackEvent({
            name: 'confidentialityHttpTrigger',
            properties: { source: LOG_SOURCE, requestBody: JSON.stringify(requestBody) }
        });
        const loginName = request.headers.get('X-MS-CLIENT-PRINCIPAL-NAME');
        if (!loginName) throw new Error('Invalid request, user is required');
        logger.trackTrace({
            message: `Request from: ${loginName}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        await pnpjs.Init();
        switch (request.method) {
            case 'GET':
                return await handleGetMethod(requestBody, loginName);
            case 'POST':
                return await handlePostMethod(requestBody, loginName);
            default:
                return { status: 405, body: 'Method Not Allowed' };
        }

    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE, method: request.method }
        });
        const message = error instanceof Error ? error.message : 'Invalid request from confidentialityTrigger';
        return { status: 400, body: message };
    }
};

app.http('confidentialityHttpTrigger', {
    methods: ['POST', 'GET'],
    authLevel: 'function',
    route: 'confidentiality',
    handler: confidentialityHttpTrigger
});
