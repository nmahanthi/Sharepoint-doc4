import { app, InvocationContext, Timer } from "@azure/functions";
import { logger, MassImportService, pnpjs } from "az-common/dist/index.mjs";

const LOG_SOURCE = "jobSanityTimerTrigger";
export async function jobSanityTimerTrigger(myTimer: Timer, context: InvocationContext): Promise<void> {
    context.log('Timer function processed request.');
    logger.init(context);
    const massImportService = new MassImportService();

    try {
        logger.trackEvent({
            name: 'jobSanityTimerTrigger',
            properties: { source: LOG_SOURCE }
        });

        logger.trackTrace({
            message: `Initializing mass import service`,
            severity: 'Verbose',
            properties: { source: LOG_SOURCE }
        });
        await massImportService.ensureTable();
        await pnpjs.Init(true);
        const filter = `status ne 'Completed' and status ne 'Failed'`;
        logger.trackTrace({
            message: `Getting incomplete jobs. Filter: ${filter}`,
            severity: 'Verbose',
            properties: { source: LOG_SOURCE }
        });
        const jobs = await massImportService.getImportJobs(filter, 10);
        //Sanity check for jobs
        for (const job of jobs) {
            try {
                const sp = pnpjs.sp(job.siteUrl);
                massImportService.setSP(sp);
                const _graph = pnpjs.graph();
                massImportService.setGraph(_graph);
                switch (job.status) {
                    case 'Processing':
                        if (!job.documents?.some(doc => doc.status === 'Processing')) {
                            if (job.documents?.every(doc => doc.status === 'Completed' || doc.status === 'Failed')) {
                                logger.trackTrace({
                                    message: `Completing job ${job.id}. Reason: All documents are completed or failed`,
                                    severity: 'Information',
                                    properties: { source: LOG_SOURCE }
                                });
                                await massImportService.completeImportJob(job);
                            }
                        }
                        break;
                    case 'Queued':
                        //If last updated time is more than a day, then mark it as failed
                        if (new Date().getTime() - new Date(job.timestamp).getTime() > 24 * 60 * 60 * 1000) {
                            logger.trackTrace({
                                message: `Failing job ${job.id}. Reason: Job is in Queued state for more than 24 hours`,
                                severity: 'Information',
                                properties: { source: LOG_SOURCE }
                            });
                            await massImportService.failImportJob(job, 'Job is in Queued state for more than 24 hours', true, job.attempts);
                        }
                        break;
                }
            } catch (error) {
                logger.trackException({
                    exception: error,
                    properties: { source: LOG_SOURCE }
                });
            }
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

app.timer('jobSanityTimerTrigger', {
    schedule: '0 */5 * * * *',
    handler: jobSanityTimerTrigger
});
