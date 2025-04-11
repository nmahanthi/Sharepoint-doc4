const LOG_SOURCE = "Service";
const REFERENCE_TABLE_PARTITION_KEY = "AlstomReferenceID";
const REFERENCE_LOCK_ROW_ID = "Lock";
import { TableClient, TableEntityResult, TableServiceClient, TransactionAction } from "@azure/data-tables";
import { logger, IDmsEvent } from "az-common/dist/index.mjs";
import { AppSettings } from "../AppSettings.js"
import { SPFI } from "@pnp/sp";
import "@pnp/sp/presets/all.js";
import { IDmsReference } from "../model/index.js";
export class Service {
    private _tableServiceClient: TableServiceClient;
    private _jobTable: TableClient;
    private _sp: SPFI;
    /**
     *
     */
    constructor(
    ) {
        this._tableServiceClient = TableServiceClient.fromConnectionString(AppSettings.MassImportTableConnectionString, {
            allowInsecureConnection: AppSettings.IsLocalEnvironment
        });
    }
    public setSP(sp: any): void {
        this._sp = sp;
    }

    public async ensureTable(): Promise<void> {
        try {
            await this._tableServiceClient.createTable(AppSettings.ReferenceIdTableName);
            this._jobTable = TableClient.fromConnectionString(AppSettings.MassImportTableConnectionString, AppSettings.ReferenceIdTableName, {
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

    public async updateReferenceID(refID: string, projRef: string, projCode: string): Promise<void> {
        const transactions: unknown[] = [];
        transactions.push(["upsert", {
            partitionKey: REFERENCE_TABLE_PARTITION_KEY,
            rowKey: `${projCode}-${projRef}`,
            referenceId: refID,
            ProjectReference: projRef,
            ProjectCode: projCode
        }]);
        transactions.push(["upsert", {
            partitionKey: REFERENCE_TABLE_PARTITION_KEY,
            rowKey: "DefaultKey",
            referenceId: refID,
            ProjectReference: "Default",
            ProjectCode: "Default"
        }]);
        await this._jobTable.submitTransaction(transactions as TransactionAction[]);
    }

    public getNewReferenceID(refId: string): string {
        //refId: DOC00099
        const getPart = refId.replace(/[^\d.]/g, ''); // returns 00099
        const num = parseInt(getPart); // returns 99
        const newVal = (num + 1).toString().padStart(getPart.length, '0'); // returns 00100
        const newstring = refId.replace(getPart, newVal); // returns DOC00100
        return newstring;
    }

    public async updateSPReferenceID(request: IDmsEvent, refId: string, userLogin: string): Promise<void> {
        const result = await this._sp.web.lists.getById(request.libraryId).items.getById(request.itemId)
            .validateUpdateListItem([
                { "FieldName": "AlstomReference", "FieldValue": refId },
                { "FieldName": "Editor", "FieldValue": JSON.stringify({ Key: userLogin }) }
            ]);
        if (!result && result.every(x => !x.HasException)) {
            {
                throw new Error(`Failed to update the reference ID in the document library. Details: ${result ?? JSON.stringify(result)}`);
            }
        }
    }

    public async getEventData(request: IDmsEvent) {
        return await this._sp.web.lists.getById(request.libraryId).items.getById(request.itemId).select("ProjectReference", "AlstomReference")();
    }

    public async getReferenceID(projRef: string, projCode: string): Promise<any> {
        try {
            const results = this._jobTable.listEntities({ queryOptions: { filter: `PartitionKey eq '${REFERENCE_TABLE_PARTITION_KEY}' and ProjectReference eq '${projRef}' and  ProjectCode eq '${projCode}'` } });
            let entityObj = null;
            console.log(results);
            for await (const result of results) {
                entityObj = result;
            }

            return entityObj;
        } catch (error) {
            if (error.statusCode === 404) {
                return null;
            }
            logger.trackException({
                exception: error,
                properties: { source: LOG_SOURCE }
            });
            throw new Error(error);
        }
    }
    public async setLock(instanceId: string): Promise<void> {
        let currentLock: {
            instanceId: string;
        } = null;
        logger.trackTrace({ message:`Trying to acquire lock for instance ${instanceId}`,properties: { source: LOG_SOURCE },severity:'Information' });
        try {
        currentLock = await this._jobTable.getEntity<{ instanceId: string }>(REFERENCE_TABLE_PARTITION_KEY, REFERENCE_LOCK_ROW_ID);
        } catch (error) {
            if (error.statusCode === 404) {
                currentLock = null;
            } else {
                throw new Error(error);
            }
        }
        let waitCount = 0;
        while (currentLock && currentLock.instanceId) {
            if (waitCount > 3*60) {
                throw new Error("Unable to acquire lock");
            }
            await (new Promise((resolve) => setTimeout(resolve, 1000)));
            waitCount++;
            try {
                currentLock = await this._jobTable.getEntity<{ instanceId: string }>(REFERENCE_TABLE_PARTITION_KEY, REFERENCE_LOCK_ROW_ID);
                } catch (error) {
                    if (error.statusCode === 404) {
                        currentLock = null;
                    } else {
                        throw new Error(error);
                    }
                }
        }
        logger.trackTrace({ message:`Acquired lock for instance ${instanceId}`,properties: { source: LOG_SOURCE },severity:'Information' });

        await this._jobTable.upsertEntity({
            partitionKey: REFERENCE_TABLE_PARTITION_KEY,
            rowKey: REFERENCE_LOCK_ROW_ID,
            instanceId,
            modified:JSON.stringify(new Date())
        });
    }
    public async getLockModifiedDate():Promise<Date | undefined>{
        let currentLock : {instanceId?:string; modified?:string;} =null;
        try {
            currentLock = await this._jobTable.getEntity<{instanceId?:string; modified?:string;}>(REFERENCE_TABLE_PARTITION_KEY, REFERENCE_LOCK_ROW_ID);
        } catch (error) {
            if (error.statusCode === 404) {
                currentLock = null;
            } else {
                throw new Error(error);
            }
        }
        return currentLock?.modified ? JSON.parse(currentLock?.modified) : undefined;
    }

    public async releaseLock(): Promise<void> {
        logger.trackTrace({ message:`Releasing lock`,properties: { source: LOG_SOURCE },severity:'Information' });
        try {
            await this._jobTable.deleteEntity(REFERENCE_TABLE_PARTITION_KEY,REFERENCE_LOCK_ROW_ID);
        } catch (error) {
            if (error.statusCode === 404) {
                logger.trackTrace({ message:`Lock not found`,properties: { source: LOG_SOURCE },severity:'Warning' });
            } else {
                throw new Error(error);
            }
        }
    }
    public async ensureLatestRefUpdated(expectedValue:string): Promise<void>{
        let latestRef = await this._jobTable.getEntity<IDmsReference>(REFERENCE_TABLE_PARTITION_KEY,"DefaultKey");
        let waitCount = 0;
        while(latestRef.referenceId !== expectedValue){
            if(waitCount > 10){
                throw new Error('Could not ensure latest ref update');
            }
            await (new Promise((resolve) => setTimeout(resolve, 1000)));
            waitCount++;
            latestRef = await this._jobTable.getEntity<IDmsReference>(REFERENCE_TABLE_PARTITION_KEY,"DefaultKey");
        }
    }
}
