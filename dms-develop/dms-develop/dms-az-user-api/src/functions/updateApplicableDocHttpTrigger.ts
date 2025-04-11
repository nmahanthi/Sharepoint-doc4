import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger } from "az-common/dist/index.mjs";
import { IUpdateApplicableDocRequest } from "../model/IUpdateApplicableDocRequest";
import { pnpjs } from "az-common/dist/index.mjs";
import { PermissionsService } from "../service/PermissionsService.js";
import { DmsRole } from "../service/IPermissionsService.js";
import { PermissionKind } from "@pnp/sp/security/index.js";
import { Constants } from "../Constants.js";
import { SPUtilityService } from "az-common/dist/index.mjs";

const LOG_SOURCE = "updateApplicableDocHttpTrigger";
export async function updateApplicableDocHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    // process.env.NODE_TLS_REJECT_UNAUTHORIZED='0';
    try {
        const requestBody: IUpdateApplicableDocRequest = await request.json() as IUpdateApplicableDocRequest;
        logger.trackEvent({
            name: 'updateApplicableDocHttpTrigger',
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
        if (!requestBody.siteUrl) throw new Error('Invalid request, url is required');
        const sp = pnpjs.sp(requestBody.siteUrl);
        const permissionsService = new PermissionsService(sp);
        const spUtilitySvc = new SPUtilityService(sp);
        const roles = await permissionsService.getDmsRoles(loginName);
        logger.trackTrace({
            message: `DMS roles: ${roles.join(', ')}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });

        // Check permissions for Admin and docController role
        const applicableRole = roles.indexOf(DmsRole.ProjectAdmin) > -1 || roles.indexOf(DmsRole.DocumentController) > -1;
        if (!applicableRole) {
            return { status: 403, body: 'Forbidden' };
        }

        //Get list Id
        const list = await spUtilitySvc.getListFromName(Constants.ApplicableDocuments, 'Id', undefined, true);

        if (requestBody.properties) {
            logger.trackTrace({
                message: `Updating properties: ${JSON.stringify(requestBody.properties)}`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
            // requestBody.properties ex:[{FieldName:"Title",FieldValue:"sekhar"},{FieldName:"Title",FieldValue:"sekhar"}...]
            // updating field Values
            const updateData = {};
            // requestBody.properties.map((data: any) => updateData[data.FieldName] = data.FieldValue);
            const data: { FieldName?: string; FieldValue?: string; }[] = requestBody.properties;
            logger.trackTrace({
                message: `Updating properties constructed Object: ${JSON.stringify(data)}`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
            // await sp.web.lists.getById(list.Id).items.getById(requestBody.docSetId).update(updateData);

            // updating Modified by User Value
            const modifiedByUser = {
                FieldName: "Editor",
                FieldValue: JSON.stringify([{ "Key": `i:0#.f|membership|${loginName}` }])
            };
            data.push(modifiedByUser);
            logger.trackTrace({
                message: `Updating Properties: ${JSON.stringify(data)}`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
            // await sp.web.lists.getById(list.Id).items.getById(requestBody.docSetId).validateUpdateListItem([modifiedByUser]);
            const resp = await sp.web.lists.getById(list.Id).items.getById(requestBody.docSetId).validateUpdateListItem(data)
                .then(validatedResponse => {
                    logger.trackTrace({
                        message: `response: ${JSON.stringify(validatedResponse)}`,
                        severity: "Verbose",
                        properties: { source: LOG_SOURCE }
                    });
                    return { status: 200, body: 'Success' };
                })
                .catch(error => {
                    logger.trackTrace({
                        message: error,
                        severity: "Verbose",
                        properties: { source: LOG_SOURCE }
                    })
                    return { status: 500, body: 'Internal server error, failed to update values ' + JSON.stringify(error) };

                });
            return resp;
        } else {
            logger.trackTrace({
                message: `No properties to update`,
                severity: "Verbose",
                properties: { source: LOG_SOURCE }
            });
            return { status: 500, body: `request body error ` + JSON.stringify(requestBody.properties) };
        }
        // return { status: 200, body: 'Success' };
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE, method: request.method }
        });
        const message = error instanceof Error ? error.message : 'Invalid request';
        return { status: 400, body: message };
    }
};

app.http('updateApplicableDocHttpTrigger', {
    methods: ['POST'],
    authLevel: 'function',
    route: 'applicableDocument',
    handler: updateApplicableDocHttpTrigger
});
