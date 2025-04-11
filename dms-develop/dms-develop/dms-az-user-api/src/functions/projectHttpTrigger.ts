import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger, ProjectPropertiesService } from "az-common/dist/index.mjs";
import { IProjectRequest } from "../model/IProjectRequest";
import { pnpjs } from "az-common/dist/index.mjs";
import { PermissionsService } from "../service/PermissionsService.js";
import { DmsRole } from "../service/IPermissionsService.js";
import { PermissionKind } from "@pnp/sp/security/index.js";
import { Constants } from "../Constants.js";
import { SPUtilityService } from "az-common/dist/index.mjs";

const LOG_SOURCE = "projectHttpTrigger";
export async function projectHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    try {
        const requestBody: IProjectRequest = await request.json() as IProjectRequest;
        logger.trackEvent({
            name: 'projectHttpTrigger',
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
        if (!requestBody.url) throw new Error('Invalid request, url is required');
        const sp = pnpjs.sp(requestBody.url);
        const projectPropertiesService = new ProjectPropertiesService(requestBody.url);
        const permissionsService = new PermissionsService(sp);
        const spUtilitySvc = new SPUtilityService(sp);
        const roles = await permissionsService.getDmsRoles(loginName);
        logger.trackTrace({
            message: `DMS roles: ${roles.join(', ')}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        if (roles.indexOf(DmsRole.ProjectAdmin) < 0) {
            const isAdministrator = await permissionsService.checkWebPermission(PermissionKind.ManageWeb, loginName);
            if (!isAdministrator)
                return { status: 403, body: 'Forbidden' };
        }

        //Get list Ids
        const lists = await Promise.all(
            Constants.docLibNames.map(
                libName => {
                    return spUtilitySvc.getListFromName(libName, 'Id', undefined, true);
                }
            )
        );

        if (requestBody.properties) {
            logger.trackTrace({
                message: `Updating properties: ${JSON.stringify(requestBody.properties)}`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });

            //Update site-level default values
            await Promise.all(requestBody.properties.filter(p => !!p.spFieldName).map(p => (
                sp.web.fields.getByInternalNameOrTitle(p.spFieldName).update({ DefaultValue: p.spValue })
            )));
            //Update list-level default values
            for (const field of requestBody.properties.filter(p => !!p.spFieldName)) {
                await (new Promise<void>((resolve) => { setTimeout(() => { resolve(); }, 200); }));
                await Promise.all(lists.map(list => () => (
                    sp.web.lists.getById(list.Id).fields.getByInternalNameOrTitle(field.spFieldName).update({ DefaultValue: field.spValue })
                )));
            }
            //Update property bag values
            await projectPropertiesService.setProperties(requestBody.properties.reduce((acc, p) => {
                acc[p.key] = p.value;
                return acc;
            }, {}));
        } else {
            logger.trackTrace({
                message: `No properties to update`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
        }
        if (requestBody.fields) {
            logger.trackTrace({
                message: `Updating fields: ${JSON.stringify(requestBody.fields)}`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
            //Update site-level fields
            await Promise.all(requestBody.fields.map(f => (
                sp.web.fields.getByInternalNameOrTitle(f.fieldName).update({
                    Hidden: f.hidden,
                    Title: f.displayName,
                    Choices: f.choices
                })
            )));
            //Update list-level fields
            for (const field of requestBody.fields) {
                logger.trackTrace({
                    message: `Updating field: ${field.fieldName}`,
                    severity: "Verbose",
                    properties: { source: LOG_SOURCE }
                });
                await (new Promise<void>((resolve) => { setTimeout(() => { resolve(); }, 200); }));
                await Promise.all(lists.map(list => (
                    sp.web.lists.getById(list.Id).fields.getByInternalNameOrTitle(field.fieldName).update({
                        Hidden: field.hidden,
                        Title: field.displayName,
                        Choices: field.choices
                    })
                )));
            }
        } else {
            logger.trackTrace({
                message: `No fields to update`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
        }
        return { status: 200, body: 'Success' };
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE, method: request.method }
        });
        const message = error instanceof Error ? error.message : 'Invalid request';
        return { status: 400, body: message };
    }
};

app.http('projectHttpTrigger', {
    methods: ['POST'],
    authLevel: 'function',
    route: 'project',
    handler: projectHttpTrigger
});
