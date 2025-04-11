import { app, InvocationContext } from "@azure/functions";
import { logger, pnpjs, SPUtilityService, IImportJob, MassImportService } from "az-common/dist/index.mjs";
import ExcelJS from 'exceljs';
const LOG_SOURCE = "newImportJobTopicTrigger";
const MaxAttempts = parseInt(process.env['MAX_ATTEMPTS'] || '9');
export async function newImportJobTopicTrigger(message: IImportJob, context: InvocationContext): Promise<void> {
    context.log('Service bus topic function processed message:', message);
    const messageCopy = {...message};
    logger.init(context);
    const massImportService = new MassImportService();
    try {
        logger.trackEvent({
            name: 'newImportJobTopicTrigger',
            properties: { source: LOG_SOURCE, requestBody: `Trigger:${JSON.stringify(context.triggerMetadata)} | Body:${JSON.stringify(message)}` }
        });
        await massImportService.ensureTable();
        // Job ID must be the Unique ID of the SP file plus modified date
        logger.trackEvent({
            name: 'newImportJobTopicTrigger',
            properties: { source: LOG_SOURCE, requestBody: `Passing:${JSON.stringify(messageCopy.jobId)}` }
        });
        const existingJob = await massImportService.getImportJob(messageCopy.jobId);
        const job = existingJob || messageCopy;
        logger.trackEvent({
            name: 'newImportJobTopicTrigger',
            properties: { source: LOG_SOURCE, requestBody: `EffectiveJob:${JSON.stringify(job)}` }
        });
        job.attempts = +context.triggerMetadata.deliveryCount;

        if (existingJob && (
            +context.triggerMetadata.deliveryCount === 1
        )) {
            throw new Error(`Job ${job.id} already exists: ${existingJob.status}`);
        }

        try {
            await pnpjs.Init(true);
            const sp = pnpjs.sp(job.siteUrl);
            massImportService.setSP(sp);
            const spUtilitySvc = new SPUtilityService(sp);

            // Check if the user is allowed to call the function app. Dms Roles are specified in the app settings, SPUtilityService will get the roles from the custom permissions list
            const isAllowed = await spUtilitySvc.canCallFnApp(job.requestorMail, job.siteUrl);
            if (!isAllowed) {
                throw new Error(`User ${job.requestorMail} is not allowed to call the function app`);
            }
            if (!job.spId || !job.spLibraryId) {
                logger.trackTrace({ message: `Getting library & item id for ${job.fileServerRelativeUrl}`, properties: { source: LOG_SOURCE }, severity: 'Information' });
                const fileItem = await sp.web.getFileByServerRelativePath(job.fileServerRelativeUrl).getItem("Id");
                context.log('fileItem', fileItem);
                const list = await fileItem.list.select('Id')();
                job.spId = fileItem['Id'];
                job.spLibraryId = list['Id'];
            }
            //TODO: Check if running job for same file
            if (existingJob)
                await massImportService.updateJob(job, 'Queued');
            // load excel file
            logger.trackTrace({ message: `Loading file ${job.fileServerRelativeUrl}`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            const excelFile = await sp.web.getFileByServerRelativePath(job.fileServerRelativeUrl).getBuffer();
            logger.trackTrace({ message: `Parsing file ${excelFile?.byteLength} bytes`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(excelFile);
            let importedDocumentStatus: 'Draft' | 'Applicable' = 'Draft';
            if (job.type === 'MIApplicable') {
                importedDocumentStatus = 'Applicable';
            }
            logger.trackTrace({ message: `Getting document import jobs`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            const docJobs = massImportService.getDocumentImportJobs(workbook, job, importedDocumentStatus)
                .filter(j =>
                    !job.documents ||
                    (job.documents.find(d => d.id === j.id) === undefined) ||
                    (job.documents.find(d => d.id === j.id)?.status === 'Failed')
                );
            job.documents = docJobs;

            // Validate rows
            // Get default values
            logger.trackTrace({ message: `Validating fields`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            await massImportService.validateDocImportJobs(job);
            
            logger.trackTrace({ message: `Applying defaults`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            await massImportService.applyDefaultsToDocImportJobs(job);

            logger.trackTrace({ message: `Creating import job`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            await massImportService.createImportJob(job);
        } catch (error) {
            logger.trackTrace({ message: `Failed. Delivery Count: ${+context.triggerMetadata.deliveryCount}, Max attempts:${MaxAttempts}, IsLast:${+context.triggerMetadata.deliveryCount >= MaxAttempts}`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            await massImportService.failImportJob(job, error, +context.triggerMetadata.deliveryCount >= MaxAttempts, +context.triggerMetadata.deliveryCount);
            throw new Error(error);
        }
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE }
        });
        throw new Error(error);
    }finally{
        massImportService && massImportService.cleanup();
    }
}

app.serviceBusTopic('newImportJobTopicTrigger', {
    connection: 'SERVICE_BUS_CONNECTION_STRING',
    topicName: 'massImportEvents',
    subscriptionName: 'newImportJob',
    isSessionsEnabled: true,
    handler: newImportJobTopicTrigger
});
