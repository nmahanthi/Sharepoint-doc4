import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger, IImportJob, MassImportService, DocumentTemplateService } from "az-common/dist/index.mjs";
import { pnpjs } from "az-common/dist/index.mjs";
import { PermissionsService } from "../service/PermissionsService.js";
import { DmsRole } from "../service/IPermissionsService.js";
import { PermissionKind } from "@pnp/sp/security/index.js";
import { ISiteUserInfo } from "@pnp/sp/site-users/types.js";
import { IContentTypeResponse, IDocumentTemplateRequest } from "../model/IDocumentTemplateRequest.js";
import { graphfi, GraphFI } from '@pnp/graph';
import { SPFI } from "@pnp/sp";

const LOG_SOURCE = "documentTemplateHttpTrigger";
async function handlePostRequest(requestBody: IDocumentTemplateRequest, sp: SPFI): Promise<HttpResponseInit> {
    try {

        const service = new DocumentTemplateService();
        service.setGraph(pnpjs.graph());
        service.setSP(sp);
        const result = await service.createNewContentType(requestBody);
        return { status: 200, body: JSON.stringify(result) };
    } catch (error) {
        return { status: 400, body: error.message };
    }
}

async function handleDeleteRequest(requestBody: IDocumentTemplateRequest, sp: SPFI, cTypeId: string): Promise<HttpResponseInit> {
    try {
        const service = new DocumentTemplateService();
        service.setSP(sp);
        const result = await service.deleteContentType("name", cTypeId);
        return { status: 200, body: "Success" };
    } catch (error) {

        return { status: 400, body: error.message };
    }
}
async function handlePutRequest(requestBody: IDocumentTemplateRequest, sp: SPFI, cTypeId: string): Promise<HttpResponseInit> {
    try {
        const service = new DocumentTemplateService();
        service.setSP(sp);
        const result = await service.updateContentType(requestBody, cTypeId);
        return { status: 200, body: JSON.stringify(result) };
    } catch (error) {
        return { status: 400, body: error.message };
    }
}

async function handleGetRequest(sp: SPFI): Promise<HttpResponseInit> {
    try {
        const service = new DocumentTemplateService();
        service.setSP(sp);
        const cTypes: IContentTypeResponse[] = await service.getAllNativeContentTypes();
        return { status: 200, body: JSON.stringify(cTypes) };
    } catch (error) {
        return { status: 400, body: error.message };
    }
}
export async function documentTemplateHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    process.env.NODE_TLS_REJECT_UNAUTHORIZED='0';
    try {

        const requestBody: IDocumentTemplateRequest | {} = request.method !== 'GET' ?
            await request.json() as IDocumentTemplateRequest :
            {
            };
        logger.trackEvent({
            name: 'documentTemplateHttpTrigger',
            properties: { source: LOG_SOURCE, requestBody: JSON.stringify(requestBody) }
        });

        const siteUrl = request.query.get('siteUrl');
        const cTypeId = request.query.get('id');

        const loginName = request.headers.get('X-MS-CLIENT-PRINCIPAL-NAME');
        if (!loginName) throw new Error('Invalid request, user is required');
        logger.trackTrace({
            message: `Request from: ${loginName}`,
            severity: "Verbose",
            properties: { source: LOG_SOURCE }
        });

        await pnpjs.Init();
        if (!siteUrl) throw new Error('Invalid request, url is required');
        const sp = pnpjs.sp(siteUrl);
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
                return await handlePostRequest(requestBody as IDocumentTemplateRequest, sp);
            case 'PUT':
                if (!cTypeId) throw new Error('Invalid request, ContentType ID is required');
                return await handlePutRequest(requestBody as IDocumentTemplateRequest, sp, cTypeId);
            case 'DELETE':
                if (!cTypeId) throw new Error('Invalid request, ContentType ID is required');
                return await handleDeleteRequest(requestBody as IDocumentTemplateRequest, sp, cTypeId);
            case 'GET':
                return await handleGetRequest(sp);
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

app.http('documentTemplateHttpTrigger', {
    methods: ['POST', 'GET', 'PUT', 'DELETE'],
    authLevel: 'function',
    route: 'documentContentType',
    handler: documentTemplateHttpTrigger
});