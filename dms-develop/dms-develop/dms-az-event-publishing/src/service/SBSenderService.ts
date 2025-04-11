import { IImportJob, logger } from "az-common/dist/index.mjs";
import { AppSettings } from "../AppSettings.js";
import { ServiceBusClient, ServiceBusSender } from "@azure/service-bus";
import { IDmsEvent } from "../model";

export class SBSenderService {
    private LOG_SOURCE = "SBSenderService";
    // private _sbSender: ServiceBusSender;
    private _sender: { [key: string]: ServiceBusSender } = {};

    constructor() {
    }

    private _getSender(topicName: string): ServiceBusSender {
        if (!this._sender[topicName]) {
            this._sender[topicName] = new ServiceBusClient(AppSettings.ServiceBusConnectionString).createSender(topicName);
        }
        return this._sender[topicName];
    }
    
    public async Send(messages: IDmsEvent[] | IImportJob[], topicName?: string): Promise<void> {
        try {
            const sender = topicName ? this._getSender(topicName) : this._getSender(AppSettings.ServiceBusTopicName);
            const batch = await sender.createMessageBatch();
            for (const msg of messages) {
                if ((msg as IDmsEvent).action) {
                    const message = msg as IDmsEvent;
                    const applicationProperties = {
                        eventType: message.action,
                        user: message.user,
                        spContentType: message.spContentType
                    };
                    if (AppSettings.MessagePromotedTaxonomyFields && message.fields) {
                        for (const field of AppSettings.MessagePromotedTaxonomyFields) {
                            const taxValue = message.fields[field] as { Label: string; TermId: string };
                            if (taxValue) {
                                applicationProperties[field] = taxValue.Label;
                            }
                        }
                    }
                    if (AppSettings.MessagePromotedFields && message.fields) {
                        for (const field of AppSettings.MessagePromotedFields) {
                            if (message.fields[field])
                                applicationProperties[field] = message.fields[field];
                        }
                    }
                    batch.tryAddMessage({
                        contentType: "application/json",
                        subject: 'SPEvent',
                        body: message,
                        sessionId: `${message.libraryId}`,
                        applicationProperties,
                    });
                } else {
                    const message = msg as IImportJob;
                    const applicationProperties = {
                        eventType: 'import',
                        user: message.requestorMail,
                        spContentType: 'import',
                        status:message.status //added this line
                    };
                    batch.tryAddMessage({
                        contentType: "application/json",
                        subject: 'newDml',
                        body: message,
                        sessionId: message.id,
                        applicationProperties //put status in properties
                    });
                }

            }
            await sender.sendMessages(batch);
        } catch (err) {
            logger.trackException({
                exception: err,
                properties: { source: this.LOG_SOURCE, method: "Send" }
            });
            throw new Error(err);
        }
    }
}
export const sbSenderService: SBSenderService = new SBSenderService();