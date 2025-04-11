import { app, InvocationContext } from "@azure/functions";
import { IDmsEvent, logger, MassImportService, pnpjs, SPUtilityService } from "az-common/dist/index.mjs";
import { Service } from "../service/Service.js";
import { IDocumentChangeRecord } from "../model/index.js";

const docIdFieldNameRest = "_dlc_DocId";
const draftLibName = "Draft";
const applicalbleLibName = "Applicable Documents";
const prevVersionLibName = "Previous Versions";
const LOG_SOURCE = 'documentChangesServiceBusTopicTrigger';

export async function documentChangesTopicTrigger(message: IDmsEvent, context: InvocationContext): Promise<void> {
    logger.init(context);
    try {
        logger.trackEvent({
            name: LOG_SOURCE,
            properties: { source: LOG_SOURCE, requestBody: JSON.stringify(message) }
        });

        const docSetId = message.fields[docIdFieldNameRest];
        if (!docSetId)
            throw new Error("Document set Id is missing as part of incoming message.");

        //Initialize service
        await pnpjs.Init(true);
        let sp = pnpjs.sp(message.siteUrl);
        const service = new Service();
        service.setSP(sp);
        let docSetInfo = null;
        docSetInfo = await service.getListDatabyDocId(draftLibName, docSetId);
        if (!docSetInfo)
            docSetInfo = await service.getListDatabyDocId(applicalbleLibName, docSetId);
        if (!docSetInfo)
            docSetInfo = await service.getListDatabyDocId(prevVersionLibName, docSetId);
        if (!docSetInfo)
            throw new Error("Document set is not found.");

        //Create if Azure table not exists
        await service.ensureTable();

        //Remove all unwanted field from the list object
        ["^[_][A-Z]", "\\.", "TaxCatchAll", "File", "MetaInfo", "SM", "ID", "ContentTypeId", "WelcomeViewId", "UniqueId",
            "ParentUniqueId", "Order", "GUID", "SelectTitle", "ScopeId", "LinkFile", "ChangeRequestID", "ServerUrl", "EncodedAbsUrl",
            "BaseName", "DocConcurrencyNumber", "_UIVersion", "owshiddenversion", "Last_x0020_Modified", "ServerRedirectedEmbedUrl", "ComplianceAssetId",
            "Edit", "WorkflowVersion", "AppEditor", "Interactivity", "RepairDocument", "PolicyDisabledUICapabilities", "PreviewThumbnailsQualitySets", "Edits",
            "StreamHash", "Combine", "ParentVersionString", "AccessPolicy", "_dlc_DocIdUrl", "TriggerFlowInfo", "xd_Signature"].forEach(k => {
                const regExp = new RegExp(`${k}`);
                const matchingEntries = Object.entries(docSetInfo).filter(([key]) => regExp.test(key));
                matchingEntries.forEach(obj => {
                    delete docSetInfo[obj[0]];
                })
            })


        const record: IDocumentChangeRecord = {
            documentID: docSetInfo[docIdFieldNameRest],
            action: message.action,
            metadata: JSON.stringify(docSetInfo),
            modifiedBy: message.user,
            projectReference: docSetInfo["ProjectReference"],
            projectRevision: docSetInfo["ProjectRevision"],
            userName: await service.getUserName(message.user)
        };
        // if (message.action === 'create') {
        //     record.changes = JSON.stringify(docSetInfo);
        //     await service.recordChanges(record);
        // } 
        if (message.action === 'create') {
            record.changes =JSON.stringify(service.getPropertyDifferencesforNew(docSetInfo));
            await service.recordChanges(record);
        }
        if (message.action === 'update') {
            const lastEditedData = await service.getLatestChanges(record.documentID);
            record.changes = JSON.stringify(service.getPropertyDifferences(docSetInfo, lastEditedData));
                await service.recordChanges(record);
        }

    } catch (error) {
        logger.trackException({
            exception: error,
            properties: { source: LOG_SOURCE }
        });
        throw new Error(error);
    }
}


app.serviceBusTopic('documentChangesTopicTrigger', {
    connection: 'dmsdev1_SERVICEBUS',
    topicName: 'documentchanges',
    subscriptionName: 'docChanges',
    handler: documentChangesTopicTrigger
});


