
import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { logger } from "az-common/dist/index.mjs";
import { AppSettings } from "../AppSettings.js";
import { TableClient, TableServiceClient } from "@azure/data-tables";
import { IFieldChange, ILogHistoryItem } from "../model/ILogHistory.js";
const LOG_SOURCE = "documentChangeHttpTrigger";

const REFERENCE_TABLE_PARTITION_KEY = "DocumentChangesRecordKey";

export async function documentChangeHttpTrigger(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    let entityObj: ILogHistoryItem[] = [];
    logger.init(context);
    try {
        const documentId = request.query.get('documentId');
        //Initialize Azure storage tables
        const tableSerClient = TableServiceClient.fromConnectionString(AppSettings.DocChangesTableConnectionString, {
            allowInsecureConnection: AppSettings.IsLocalEnvironment
        });

        //Making sure desired table exists, if not creat one
        await tableSerClient.createTable(AppSettings.DocumentChangesTableName);
        const jobTable = TableClient.fromConnectionString(AppSettings.DocChangesTableConnectionString, AppSettings.DocumentChangesTableName, {
            allowInsecureConnection: AppSettings.IsLocalEnvironment
        });
        const results = jobTable.listEntities({ queryOptions: { filter: `PartitionKey eq '${REFERENCE_TABLE_PARTITION_KEY}' and documentId eq '${documentId}'` } });

        for await (const result of results) {
            if (result.changes.toString().length > 2) {
                console.log(result);
                const changes = JSON.parse(result.changes.toString());
                //Process field change values
                let fieldchanges: IFieldChange[] = [];
                Object.keys(changes).forEach((key) => {
                    console.log(key, changes[key]);
                    fieldchanges.push({
                        fieldName: key,
                        newValue: changes[key].newValue,
                        oldValue: changes[key].oldValue
                    });
                });
                entityObj.push({
                    documentId: result.documentId.toString(),
                    projectReference: result.projectReference.toString(),
                    projectRevision: result.projectRevision.toString(),
                    userEmail: result.modifiedBy.toString(),
                    userDisplayName: result.userName.toString(),
                    Date: result.modified.toString(),
                    actionType: result.action.toString(),
                    changedFields: fieldchanges
                });
            }
        }

    } catch (err) {
        context.log(err);
        throw new Error(err);
    }
    return { status: 200, body: JSON.stringify(entityObj) };
};

app.http('documentChangeHttpTrigger', {
    methods: ['GET'],
    authLevel: 'anonymous',
    handler: documentChangeHttpTrigger
});
