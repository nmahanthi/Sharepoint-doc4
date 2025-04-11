import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger } from "az-common/dist/index.mjs";
import { pnpjs } from "az-common/dist/index.mjs";
import { IFreezeBaselineRequest } from "../model/IFreezeBaselineRequest";
import moment from "moment";
import { PermissionKind } from "@pnp/sp/security/index.js";

const LOG_SOURCE = "freezeBaselineHttpTrigger";
export async function freezeBaselineHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    try {
        const requestBody: IFreezeBaselineRequest = await request.json() as IFreezeBaselineRequest;
        logger.trackEvent({
            name: 'freezeBaselineHttpTrigger',
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
        //check that all properties in IFreezeBaselineRequest are present
        if (!requestBody.siteUrl || !requestBody.listId || !requestBody.baselineId) {
            throw new Error('Invalid request, siteUrl, listId and baselineId are required');
        }
        const sp = pnpjs.sp(requestBody.siteUrl);

        logger.trackTrace({
            message: `Ensuring user: ${loginName}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        const requestor = await sp.web.ensureUser(loginName);

        //Check if user has edit permissions on the item
        logger.trackTrace({
            message: `Checking user permissions`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        const userHasPermissions = await sp.web.lists.getById(requestBody.listId).items.getById(requestBody.baselineId).userHasPermissions(requestor.LoginName, PermissionKind.EditListItems);
        if (!userHasPermissions) {
            return { status: 403, body: 'You do not have permission to edit the item' };
        }

        //Get existing role assignments
        logger.trackTrace({
            message: `Retrieving role assignments`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        const roleAssignments = await sp.web.lists.getById(requestBody.listId).items.getById(requestBody.baselineId).roleAssignments();

        //Replace existing role assignments with read permissions
        logger.trackTrace({
            message: `Replacing role assignments with read permissions`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        await sp.web.lists.getById(requestBody.listId).items.getById(requestBody.baselineId).breakRoleInheritance(false);
        await Promise.all(roleAssignments.map(ra => (
            sp.web.lists.getById(requestBody.listId).items.getById(requestBody.baselineId).roleAssignments.add(ra.PrincipalId, 1073741826)
        )));

        //Retrieve baseline title
        logger.trackTrace({
            message: `Retrieving baseline title`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        const baseline = await sp.web.lists.getById(requestBody.listId).items.getById(requestBody.baselineId).select('Title')();

        //Update baseline
        logger.trackTrace({
            message: `Updating baseline: ${requestBody.baselineId}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        const updateResult = await sp.web.lists.getById(requestBody.listId).items.getById(requestBody.baselineId).validateUpdateListItem([
            { FieldName: 'Editor', FieldValue: JSON.stringify([{ "Key": requestor.LoginName }]) },
            { FieldName: 'IsLatest', FieldValue: "1" },
            { FieldName: 'Modified', FieldValue: moment().format('M/d/YYYY hh:mm A') },
            { FieldName: 'BaselineStatus', FieldValue: 'Frozen' }
        ]);
        if (updateResult.some(r => r.HasException)) {
            throw new Error(updateResult.filter(r => r.HasException).map(r => `${r.FieldName}: ${r.ErrorMessage}`).join('\n'));
        }

        //Retrieve other baselines with same title and IsLatest = 1
        logger.trackTrace({
            message: `Retrieving baselines with same title: ${baseline.Title}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        const baselines = await sp.web.lists.getById(requestBody.listId).items
            // .filter(`(Title eq '${baseline.Title}') and (Id ne ${requestBody.baselineId}) and (IsLatest eq true)`)
            .filter(`(Title eq '${baseline.Title}') and (Id ne ${requestBody.baselineId}) and (IsLatest eq 1)`)
            .select('Id')();

        //Set IsLatest = 0 for other baselines
        logger.trackTrace({
            message: `Setting IsLatest = 0 for other baselines`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        await Promise.all(baselines.map(b => (
            sp.web.lists.getById(requestBody.listId).items.getById(b.Id).update({ IsLatest: false })
        )));
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

app.http('freezeBaselineHttpTrigger', {
    methods: ['POST'],
    authLevel: 'function',
    route: 'baseline/freeze',
    handler: freezeBaselineHttpTrigger
});
