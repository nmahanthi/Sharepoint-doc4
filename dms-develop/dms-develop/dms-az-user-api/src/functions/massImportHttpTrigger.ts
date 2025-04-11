import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger, IImportJob, MassImportService } from "az-common/dist/index.mjs";
import { pnpjs } from "az-common/dist/index.mjs";
import { PermissionsService } from "../service/PermissionsService.js";
import { DmsRole } from "../service/IPermissionsService.js";
import { PermissionKind } from "@pnp/sp/security/index.js";
import { IMassImportRequest } from "../model/IMassImportRequest";
import { AppSettings } from "../AppSettings.js";
import { ServiceBusClient, ServiceBusMessage } from "@azure/service-bus";
import { ISiteUserInfo } from "@pnp/sp/site-users/types.js";

const LOG_SOURCE = "massImportPostHttpTrigger";
async function handlePostRequest(requestBody: IMassImportRequest, user: ISiteUserInfo): Promise<HttpResponseInit> {
    const sbSender = new ServiceBusClient(AppSettings.ServiceBusConnectionString).createSender(AppSettings.MassImportTopicName);
    try {
        await sbSender.sendMessages({
            contentType: 'application/json',
            sessionId: requestBody.id,
            subject: AppSettings.NewMassImportSubject,
            body: {
                attempts: 0,
                documents: undefined,
                error: undefined,
                fileServerRelativeUrl: requestBody.fileServerRelativeUrl,
                id: requestBody.id,
            jobId:requestBody.id,
            requestorMail: user.Email,
                status: undefined,
                siteUrl: requestBody.siteUrl,
                spId: requestBody.spId,
                spListId: requestBody.spListId,
                type: requestBody.type
            } as IImportJob,
            applicationProperties: { subject: requestBody.type }
        } as ServiceBusMessage);
    } catch (error) {
        sbSender.close();
        throw new Error(error);
    }
    return { status: 200, body: 'Success' };
}
async function handleGetRequest(requestBody: { id: string | string [] }): Promise<HttpResponseInit> {
    const massImportService = new MassImportService();
    try {
        if(!requestBody.id) throw new Error('Invalid request, id is required');
        await massImportService.ensureTable();
        if(Array.isArray(requestBody.id)){
            const jobs : IImportJob[] = [];
            const ids = requestBody.id as string[];
            //Process ids in batches of 15 as per documetation limit
            for(let i = 0; i < ids.length; i += 15){
                const batchFilter = ids.slice(i, i + 15).map(id => `id eq '${id}'`).join(' or ');
                const batchJobs = await massImportService.getImportJobs(`(${batchFilter})`, 15);
                jobs.push(...batchJobs);
            }
            if (jobs.length === 0) return { status: 404, body: 'Not found' };
            return { status: 200, body: JSON.stringify(jobs) };
        }else{
            const job = await massImportService.getImportJob(requestBody.id);
            if (!job) return { status: 404, body: 'Not found' };
            return { status: 200, body: JSON.stringify(job) };
        }
    } catch (error) {
        return { status: 400, body: error.message };
    }finally{
        massImportService && massImportService.cleanup();
    }
}
export async function massImportPostHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    try {
        const requestBody: IMassImportRequest | { id: string; siteUrl: string } = request.method === 'POST' ?
            await request.json() as IMassImportRequest :
            {
                id: request.query.get('id'),
                siteUrl: request.query.get('siteUrl'),
            };
        logger.trackEvent({
            name: 'massImportPostHttpTrigger',
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
        const user = await sp.web.ensureUser(loginName);
        const roles = await permissionsService.getDmsRoles(user.LoginName);
        logger.trackTrace({
            message: `DMS roles: ${roles.join(', ')}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });
        if (roles.indexOf(DmsRole.DocumentController) < 0) {
            const isAdministrator = await permissionsService.checkWebPermission(PermissionKind.ManageWeb, loginName);
            if (!isAdministrator)
                return { status: 403, body: 'Forbidden' };
        }
        switch (request.method) {
            case 'POST':
                return await handlePostRequest(requestBody as IMassImportRequest, user);
            case 'GET':
                let id: string | string[] = (requestBody as { id: string })?.id;
                if(id?.includes(',')) id = id.split(',');
                return await handleGetRequest(requestBody as { id: string });
            default:
                return { status: 400, body: 'Invalid request' };
        }
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE, method: request.method }
        });
        const message = error instanceof Error ? error.message : 'Invalid request';
        return { status: 400, body: message };
    }
};

app.http('massImportPostHttpTrigger', {
    methods: ['POST', 'GET'],
    authLevel: 'function',
    route: 'massImport',
    handler: massImportPostHttpTrigger
});