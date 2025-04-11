const LOG_SOURCE = "Service";
const REFERENCE_TABLE_PARTITION_KEY = "DocumentChangesRecordKey";
import { TableClient, TableEntityResult, TableServiceClient, TransactionAction } from "@azure/data-tables";
import { logger, IDmsEvent } from "az-common/dist/index.mjs";
import { AppSettings } from "../AppSettings.js"
import { spfi, SPFI } from "@pnp/sp/index.js";
import "@pnp/sp/presets/all.js";
import { IRenderListData, RenderListDataOptions } from "@pnp/sp/presets/all.js";
import { IDocumentChangeRecord } from "../model/index.js";
export class Service {
    private _tableServiceClient: TableServiceClient;
    private _jobTable: TableClient;
    private _sp: SPFI;

    constructor(
    ) {
        this._tableServiceClient = TableServiceClient.fromConnectionString(AppSettings.DocChangesTableConnectionString, {
            allowInsecureConnection: AppSettings.IsLocalEnvironment
        });
    }
    public setSP(sp: SPFI): void {
        this._sp = sp;
    }

    public async getListDatabyDocId(listName: string, docId: string): Promise<any> {
        const itemsData = (await this._sp.web.lists.getByTitle(listName).renderListDataAsStream({
            ViewXml: `<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='_dlc_DocId' /><Value Type='Text'>${docId}</Value></Eq></Where></Query><RowLimit>1</RowLimit></View>`,
            RenderOptions: RenderListDataOptions.ListData,
        })).Row[0];

        return itemsData;
    }

    public async ensureTable(): Promise<void> {
        try {
            await this._tableServiceClient.createTable(AppSettings.DocumentChangesTableName);
            this._jobTable = TableClient.fromConnectionString(AppSettings.DocChangesTableConnectionString, AppSettings.DocumentChangesTableName, {
                allowInsecureConnection: AppSettings.IsLocalEnvironment
            });
        } catch (error) {
            logger.trackException({
                exception: error,
                properties: { source: LOG_SOURCE }
            });
            throw new Error(error);
        }
    }

    public async getUserName(userName: string): Promise<string> {
        const user = await this._sp.web.ensureUser(userName);
        return user.Title;

    }

    public async recordChanges(record: IDocumentChangeRecord): Promise<void> {

        await this._jobTable.upsertEntity({
            partitionKey: REFERENCE_TABLE_PARTITION_KEY,
            rowKey: Date.now().toString(),
            action: record.action,
            actualData: record.metadata,
            changes: record.changes,
            documentId: record.documentID,
            modified: new Date().toLocaleString(),
            modifiedBy: record.modifiedBy,
            projectReference:record.projectReference,
            projectRevision:record.projectRevision,
            userName:record.userName
        });
    }

    public async getLatestChanges(documentId: string): Promise<any> {
        try {
            const results = this._jobTable.listEntities({ queryOptions: { filter: `PartitionKey eq '${REFERENCE_TABLE_PARTITION_KEY}' and documentId eq '${documentId}'` } });
            let entityObj = [];
            for await (const result of results) {
                entityObj.push(result);
            }
            logger.trackEvent({
                name: LOG_SOURCE,
                properties: { source: LOG_SOURCE, requestBody: `Retrieved ${entityObj.length} records from Azure table for the document id ${documentId}` }
            });
            if (entityObj.length > 0) {
                //sort by last modified and get the actual list item data
                entityObj.sort(function (a, b) {
                    const aDate = new Date(a.modified);
                    const bDate = new Date(b.modified);
                    return bDate.valueOf() - aDate.valueOf();
                });
                return JSON.parse(entityObj[0].actualData);
            }
            return {};
        } catch (error) {
            if (error.statusCode === 404) {
                return {};
            }
            logger.trackException({
                exception: error,
                properties: { source: LOG_SOURCE }
            });
            throw new Error(error);
        }
    }

    public getPropertyDifferences(newObj, oldObj) {
        return Object.entries(oldObj).reduce((diff, [key, value]) => {
            if (typeof value === 'object') {
                if (value.hasOwnProperty("TermID") ) {
                    if (newObj[key]["TermID"] !== oldObj[key]["TermID"])
                        return {
                            ...diff,
                            [key]: {
                                "oldValue": oldObj[key],
                                "newValue": newObj[key]
                            },
                        };
                }
                else if (Array.isArray(value)) {
                    if (JSON.stringify(newObj[key]) !== JSON.stringify(oldObj[key]))
                        return {
                            ...diff,
                            [key]: {
                                "oldValue": oldObj[key],
                                "newValue": newObj[key]
                            },
                        };
                }

            } else {
                // Check if the property exists in oldObj.
                if (oldObj.hasOwnProperty(key)) {
                    const val = newObj[key];

                    // Check if newObj's property's value is different from oldObj's.
                    if (val !== value) {
                        return {
                            ...diff,
                            [key]: {
                                "oldValue": value,
                                "newValue": val
                            },
                        };
                    }
                }
            }

            // Otherwise, just return the previous diff object.
            return diff;
        }, {});
    }

    public getPropertyDifferencesforNew(newObj) {
        return Object.entries(newObj).reduce((diff, [key, value]) => {
            if (typeof value === 'object') {
                if (value.hasOwnProperty("TermID") || value.hasOwnProperty("email")) {
                        return {
                            ...diff,
                            [key]: {
                                "oldValue": {},
                                "newValue": newObj[key]
                            },
                        };
                }
            } else {
                    const val = newObj[key];
                    if(val && val.length > 0)
                        return {
                            ...diff,
                            [key]: {
                                "oldValue": "",
                                "newValue": val
                            },
                        };
            }
            return diff;
        }, {});
    }

}
