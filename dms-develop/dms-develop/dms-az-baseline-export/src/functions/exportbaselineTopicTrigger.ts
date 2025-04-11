import { app, InvocationContext } from "@azure/functions";
import fs from 'fs';
import { SPFI } from "@pnp/sp";
import AppSettings from "../AppSettings.js";
import { IBaseLineExportEvent } from "../model/IBaselineRequest.js";
import { logger, pnpjs } from "az-common/dist/index.mjs";
import { IEmailProperties } from "@pnp/sp/sputilities/index.js";
import { getGUID } from "@pnp/core";
import { FileHelper } from "../service/FileHelper.js";
import path from "path";
import os from "os";

const LOG_SOURCE = 'baselineExportServicebusQueueTrigger';

let requestedUserName = "";
export async function exportBaselineTopicTrigger(message: IBaseLineExportEvent, context: InvocationContext): Promise<void> {
    logger.init(context);
    try {
        const siteName = message.SiteUrl?.split('/').pop() || 'unknown';
        const baselineId = `${siteName}-${message.RootBaseLine?.ItemId || 'unknown'}`;
        context.log('Service bus topic function processed message:', message);
        logger.init(context);
        logger.trackEvent({
            name: LOG_SOURCE,
            properties: { source: baselineId, requestBody: JSON.stringify(message), requestQuery: JSON.stringify(message) }
        });
        const hostUrl = `https://${AppSettings.TenantName}.sharepoint.com`;
        await pnpjs.Init(true);
        const sp = pnpjs.sp(message.SiteUrl);
        const fileHelper = new FileHelper({ sp });
        logger.trackTrace({
            message: `Initialized`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Information'
        });
        const fileMap = await fileHelper.getFileMapFromRequest(message, hostUrl);
        logger.trackTrace({
            message: `Created File Map`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Information'
        });
        logger.trackTrace({
            message: `File Map ${JSON.stringify(fileMap)}`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Verbose'
        });
        const archive = await fileHelper.getArchiveFromFileMap(fileMap);
        logger.trackTrace({
            message: `Created Archive`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Information'
        });

        const baselineName = AppSettings.FolderNameFields.map(f => fileMap[0].metadata[f]).join('-');
        const buffer = await fileHelper.getExcelFileFromMap(fileMap, baselineId, baselineName);

        logger.trackTrace({
            message: `Created Excel`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Information'
        });

        archive.append(buffer, { name: `Metadata.xlsx` });
        const newGuid = getGUID();
        const zipFileName = `${baselineName}_${newGuid}.zip`;
        const tempPath = path.join(AppSettings.TempZipFolder, zipFileName);
        const output = fs.createWriteStream(tempPath, { highWaterMark: 1024 * 1024 * 100 });
        output.on('error', (err) => {
            logger.trackException({
                exception: err,
                properties: { source: baselineId, method: 'Main' },
            });
        });
        output.on('close', async () => {
            logger.trackTrace({
                message: `Zip file saved locally`,
                properties: { source: baselineId, method: 'Main', request_type: 'NA' },
                severity: 'Information'
            });
        });
        archive.pipe(output);
        archive.finalize();
        await (new Promise<void>((resolve, reject) => {
            output.on('finish', () => {
                resolve();
            });
            output.on('error', (err) => {
                reject(err);
            });
        }));
        const uploadedFile = await fileHelper.uploadZipToSharepoint(tempPath, zipFileName);
        logger.trackTrace({
            message: `Zip file uploaded to sharepoint`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Information'
        });
        //delete temp file
        fs.unlinkSync(tempPath);
        logger.trackTrace({
            message: `Deleted temp file`,
            properties: { source: baselineId, method: 'Main', request_type: 'NA' },
            severity: 'Information'
        });

        requestedUserName = message.UserEmail;

        const emailBody = `The requested baseline export is available here : <a href="${hostUrl}${uploadedFile}"> ${uploadedFile.split('/').pop()} </a> >. <br/> 
        The export will be available for 3 days before auto removal.<br/><br/>Regards,<br/>DOC4A Team`;
        // `Export completed as per your request, please find the below url of the exported file.<br/><a href="${url}"> ${url} </a>`;
        sendEmail(sp, `DOC4A - Export Baseline ${baselineName}`, emailBody);
        logger.trackTrace({
            message: `Sent email.`,
            properties: { source: LOG_SOURCE },
            severity: 'Information'
        });
    }
    catch (err) {
        logger.trackException({
            exception: err,
            properties: { source: LOG_SOURCE }
        });
        context.log(err);
        throw new Error(err);
    }

}
//Send email after completeion or on any error
async function sendEmail(sp: SPFI, subject: string, message: string) {
    try {
        const emailProps: IEmailProperties = {
            To: [requestedUserName],
            // CC: ["user2@site.com", "user3@site.com"],
            //  BCC: ["user4@site.com", "user5@site.com"],
            Subject: subject,
            Body: message,
            AdditionalHeaders: {
                "content-type": "text/html"
            }
        };

        await sp.utility.sendEmail(emailProps);
    } catch (err) {
        logger.trackException({
            exception: err,
            properties: { source: LOG_SOURCE }
        });
        throw new Error(err);
    }
}

app.serviceBusTopic('exportBaselineTopicTrigger', {
    connection: 'dmsdev1_SERVICEBUS',
    topicName: 'baseline',
    subscriptionName: 'Baseline',
    handler: exportBaselineTopicTrigger
});