import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger } from  "az-common/dist/index.mjs";
import { ServiceBusClient, ServiceBusMessage } from "@azure/service-bus";
import { AppSettings } from "../AppSettings.js";
const LOG_SOURCE = "baselineHttpTrigger";

export async function baselineHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    logger.init(context);
    try {
        const requestBody: any = await request.json();
       
        context.log(requestBody);
        logger.trackEvent({
            name: 'baselineHttpTrigger',
            properties: { source: LOG_SOURCE, requestBody: JSON.stringify(requestBody) }
        });
        const loginName = request.headers.get('X-MS-CLIENT-PRINCIPAL-NAME');
       
        const sbClient = new ServiceBusClient(AppSettings.ServiceBusConnectionString);
 
        const sbSender = sbClient.createSender(AppSettings.ServiceBusTopicName);
 
        const result = await sbSender.sendMessages({
            subject: "BaselineEvent",
            body: requestBody,
            messageId: `${`${requestBody.SiteUrl}`.replace(/\//g,'')}${requestBody.RootBaseLine.ItemId}${`${requestBody.UserEmail}`.replace(/[^a-z]/g,'')}`
        } as ServiceBusMessage);
        context.log(result);
    } catch (err) {
        context.log(err);
        throw new Error(err);
    }
    return { status: 200, body: 'Success' };
}

app.http('baselineHttpTrigger', {
    methods: ['POST'],
    authLevel: 'function',
    route: 'baseline',
    handler: baselineHttpTrigger
});
