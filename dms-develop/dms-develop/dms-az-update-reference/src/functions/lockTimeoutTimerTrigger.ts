import { app, InvocationContext, Timer } from "@azure/functions";
import { logger } from "az-common/dist/index.mjs";
import { Service } from "../service/Service.js";

const LOG_SOURCE = "lockTimeoutTimerTrigger";
const lockTimeout = +(process.env['LOCK_TIMEOUT'] || 0) || 5;
export async function lockTimeoutTimerTrigger(myTimer: Timer, context: InvocationContext): Promise<void> {
    context.log('Timer function processed request.');
    logger.init(context);
    logger.trackTrace({ message: `Retrieving lock.`, properties: { source: LOG_SOURCE }, severity: 'Information' });
    const service = new Service();
    await service.ensureTable();

    const lockDate = await service.getLockModifiedDate();
    if(!lockDate){
        logger.trackTrace({ message: `No lock.`, properties: { source: LOG_SOURCE }, severity: 'Information' });
        return;
    }
    const diffTime = Math.abs((new Date()).getTime() - lockDate.getTime());
    const diffMins = Math.floor(diffTime / (1000 * 60)); 
    if(diffMins>lockTimeout){
        logger.trackTrace({ message: `Force releasing lock. Reason: ${diffMins} > ${lockTimeout}`, properties: { source: LOG_SOURCE }, severity: 'Information' });
        await service.releaseLock();
    }else{
        logger.trackTrace({ message: `Lock duration minutes ${diffMins}`, properties: { source: LOG_SOURCE }, severity: 'Information' });
    }
}

app.timer('lockTimeoutTimerTriggerV2', {
    schedule: '0 */5 * * * *',
    handler: lockTimeoutTimerTrigger
});
