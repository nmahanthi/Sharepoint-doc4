import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger } from "az-common/dist/index.mjs";
import { pnpjs } from "az-common/dist/index.mjs";
import { changeTokenService } from "../service/ChangeTokenService.js";
import { ISpNotification } from "../model/index.js";
import { DmsEventService } from "../service/DmsEventService.js";
import { sbSenderService } from "../service/SBSenderService.js";
import { AppSettings } from "../AppSettings.js";

const LOG_SOURCE = "notifyHttpTrigger";
export async function notifyHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
  logger.init(context);
  try {
    await pnpjs.Init();
    await changeTokenService.init();
    //if webhhok validation
    if (request.method == "POST" && request.query.get("validationtoken")) {
      logger.trackTrace({
        message: 'Validation token',
        properties: {
          source: LOG_SOURCE,
          request_type: "POST"
        },
        severity: "Verbose"
      });
      return { status: 200, body: request.query.get("validationtoken") };
    } else {
      const requestBody: ISpNotification = await request.json() as ISpNotification;
      logger.trackEvent({
        name: 'notifyHttpTrigger',
        properties: { source: LOG_SOURCE, requestBody: JSON.stringify(requestBody) }
      });
      const eventService = new DmsEventService(changeTokenService);
      const events = await Promise.all(requestBody.value.map(e => eventService.getDmsEventsFromSpEvent(e)));
      const allEvents = events.flatMap(e => e.events);
      logger.trackEvent({
        name: 'notifyHttpTrigger',
        properties: { source: LOG_SOURCE, requestBody: JSON.stringify(allEvents) }
      });
      if (allEvents.length > 0) {
        const partitionKeys = allEvents.map(e => e.libraryId).filter((value, index, self) => self.indexOf(value) === index);
        for(const key of partitionKeys){
          logger.trackEvent({
            name: 'notifyHttpTrigger',
            properties: { source: LOG_SOURCE, requestBody: `Sending ${allEvents.filter(e => e.libraryId === key).length} events for ${key}` }
          });
          await sbSenderService.Send(allEvents.filter(e => e.libraryId === key), AppSettings.ServiceBusTopicName);
        }
        logger.trackEvent({
          name: 'notifyHttpTrigger',
          properties: { source: LOG_SOURCE, requestBody: `${allEvents.length} events sent` }
        });
        await Promise.all(events.map(e => changeTokenService.SaveChangeToken(e.listId, e.lastToken)));
        logger.trackEvent({
          name: 'notifyHttpTrigger',
          properties: { source: LOG_SOURCE, requestBody: 'Change token saved' }
        });
      }
    }
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

app.http('notifyHttpTrigger', {
  methods: ['POST'],
  authLevel: 'function',
  route: 'notify',
  handler: notifyHttpTrigger
});
