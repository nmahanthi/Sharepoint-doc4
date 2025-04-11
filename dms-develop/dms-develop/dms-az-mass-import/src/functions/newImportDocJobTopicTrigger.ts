import { app, InvocationContext } from "@azure/functions";
import { logger, pnpjs, IImportDocumentJob, MassImportService } from "az-common/dist/index.mjs";
const LOG_SOURCE = "newImportDocJobTopicTrigger";

const MaxAttempts = parseInt(process.env['DOC_JOB_MAX_ATTEMPTS'] || '3');
export async function newImportDocJobTopicTrigger(message: unknown, context: InvocationContext): Promise<void> {
    context.log('Service bus topic function processed message:', message);
    logger.init(context);
    const massImportService = new MassImportService();
    try {
        logger.trackEvent({
            name: 'newImportDocJobTopicTrigger',
            properties: { source: LOG_SOURCE, requestBody: JSON.stringify(message) }
        });

        let docJob = message as IImportDocumentJob;
        try {
            await massImportService.ensureTable();
            const importJob = await massImportService.getImportJob(docJob.parentJobId);
            if (!importJob) {
                throw new Error(`Parent job ${docJob.parentJobId} not found`);
            }
            if (importJob.status === 'Failed') {
                throw new Error(`Parent job ${docJob.parentJobId} has failed`);
            }
            await pnpjs.Init(true);
            const sp = pnpjs.sp(importJob.siteUrl);
            massImportService.setSP(sp);
            const _graph = pnpjs.graph();
            massImportService.setGraph(_graph);

            if (importJob.status !== 'Processing') {
                await massImportService.updateJob(importJob, 'Processing');
            }
            const existingJob = importJob.documents.find(d => d.id === docJob.id);
            docJob = existingJob || docJob;

            docJob.attempts = +context.triggerMetadata.deliveryCount;
            //Process the import job
            await massImportService.updateImportDocJob(docJob, 'Processing');
            const item = await massImportService.processImportDocJob(importJob, docJob);
            //publish event
            const itemMetadata = await item.select('Id', 'ParentList/Id').expand('ParentList')();
            switch (importJob.type) {
                case 'MIApplicable':
                    await massImportService.publishDocumentChangeEvent(itemMetadata.ParentList.Id, itemMetadata.Id, importJob.requestorMail, 'update');
                    break;
                case 'MIDraft':
                case 'DML':
                    await massImportService.publishDocumentChangeEvent(itemMetadata.ParentList.Id, itemMetadata.Id, importJob.requestorMail, 'create');
                    break;
            }
            //Update the import job status
            await massImportService.updateImportDocJob(docJob, 'Completed');

            //Check if all documents are completed
            const updatedImportJob = await massImportService.getImportJob(docJob.parentJobId);
            if (updatedImportJob.documents.every(d => d.status === 'Completed' || d.status === 'Failed')) {
                await massImportService.completeImportJob(updatedImportJob);
            }
        } catch (error) {
            await massImportService.failImportDocJob(docJob, error, +context.triggerMetadata.deliveryCount >= MaxAttempts, +context.triggerMetadata.deliveryCount);
            if (+context.triggerMetadata.deliveryCount >= MaxAttempts) {
                //Check if all documents are completed
                const updatedImportJob = await massImportService.getImportJob(docJob.parentJobId);
                if (updatedImportJob.documents.every(d => d.status === 'Completed' || d.status === 'Failed')) {
                    await massImportService.completeImportJob(updatedImportJob);
                }
            }
            //Check if all documents are completed
            const updatedImportJob = await massImportService.getImportJob(docJob.parentJobId);
            if (updatedImportJob.documents.every(d => d.status === 'Completed' || d.status === 'Failed')) {
                await massImportService.completeImportJob(updatedImportJob);
            }
            throw new Error(error);
        }
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE }
        });
        throw new Error(error);
    } finally {
        massImportService && massImportService.cleanup();
    }
}

app.serviceBusTopic('newImportDocJobTopicTrigger', {
    connection: 'SERVICE_BUS_CONNECTION_STRING',
    topicName: 'massImportEvents',
    subscriptionName: 'newImportDocJob',
    handler: newImportDocJobTopicTrigger
});
