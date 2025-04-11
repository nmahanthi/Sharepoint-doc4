import { app, InvocationContext } from "@azure/functions";
import {
    IDmsEvent, logger, pnpjs, ProjectPropertiesService
} from "az-common/dist/index.mjs";
import { Service } from "../service/Service.js";
import { AppSettings } from "../AppSettings.js";

const LOG_SOURCE = 'updateAlstomReferenceTopicTrigger';

export async function updateAlstomReferenceTopicTrigger(messages: IDmsEvent[], context: InvocationContext): Promise<void> {
    logger.init(context);

    try {
        for (const message of messages) { // Iterate through the batch of messages
            logger.trackEvent({
                name: 'updateAlstomReferenceTopicTrigger',
                properties: { source: LOG_SOURCE, requestBody: JSON.stringify(message) }
            });

            if (message.action !== "create") {
                logger.trackTrace({
                    message: `Document change action is not Create, so skipping from further proceedings.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                continue;
            }
            await pnpjs.Init(true);
            const projectPropertiesService = new ProjectPropertiesService(message.siteUrl);

            // Ensure user
            logger.trackTrace({
                message: `Ensuring user ${message.user}`,
                properties: { source: LOG_SOURCE },
                severity: 'Information'
            });
            const user = await pnpjs.sp(message.siteUrl).web.ensureUser(message.user);

            const service = new Service();
            await service.ensureTable();

            const sp = pnpjs.sp(message.siteUrl);
            service.setSP(sp);

            // Get Project code from Site Property bag
            const projCode = await projectPropertiesService.getProperty('dms_proj_code');

            // Get Document set data
            const doctSetData = await service.getEventData(message);

            if (doctSetData.AlstomReference && doctSetData.AlstomReference.length > 0) {
                logger.trackTrace({
                    message: `Document set already has the Alstom Reference, so no further actions needed.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                continue;
            }

            // Check for Data table on Project reference and Project Code
            await service.setLock(context.invocationId);
            try {
                logger.trackTrace({
                    message: `Checking if reference already assigned ${projCode}-${doctSetData.ProjectReference}.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                let ret = await service.getReferenceID(doctSetData.ProjectReference, projCode);
                if (!ret) {
                    logger.trackTrace({
                        message: `No existing ref, retrieving latest.`,
                        properties: { source: LOG_SOURCE },
                        severity: 'Information'
                    });
                    ret = await service.getReferenceID("Default", "Default");
                }
                logger.trackTrace({
                    message: `Previous reference ${JSON.stringify(ret || {})}.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                const newId = service.getNewReferenceID(ret ? ret.referenceId : AppSettings.InitialReference);
                logger.trackTrace({
                    message: `New reference ${JSON.stringify(newId || {})}.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                await service.updateReferenceID(newId, doctSetData.ProjectReference, projCode);
                logger.trackTrace({
                    message: `Updated latest ref.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                await service.ensureLatestRefUpdated(newId);
                logger.trackTrace({
                    message: `Updating SP Item.`,
                    properties: { source: LOG_SOURCE },
                    severity: 'Information'
                });
                await service.updateSPReferenceID(message, newId, user.LoginName);
            } catch (error) {
                throw new Error(error);
            } finally {
                await service.releaseLock();
            }
        }
    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE }
        });
        throw new Error(error);
    }
}

// Register the function with batch processing enabled
app.serviceBusTopic('updateAlstomReferenceTopicTrigger', {
    connection: 'SERVICE_BUS_CONNECTION_STRING',
    topicName: 'documentchanges',
    subscriptionName: 'DocsetCreated2',
    cardinality: 'many', // Enables batch processing
    handler: updateAlstomReferenceTopicTrigger
});