import { TableServiceClient, TableClient } from "@azure/data-tables";
import { logger } from "az-common/dist/index.mjs";
import moment from "moment";
import { AppSettings } from "../AppSettings.js";

const PARTITION_KEY = "ChangeToken";

export interface IChangeTokenService {
    GetChangeToken: (listId: string) => Promise<string>;
    SaveChangeToken: (listId: string, changeToken: string) => Promise<string>;
    init: () => Promise<void>;
}

export class ChangeTokenService implements IChangeTokenService {
    private LOG_SOURCE = "ChangeTokenService";
    private _tableClient: TableClient;

    public async init(): Promise<void> {
        const serviceClient = TableServiceClient.fromConnectionString(AppSettings.ChangeTokenStorageConnectionString, { allowInsecureConnection: true });
        const tableName = AppSettings.ChangeTokenTableName;
        // If the table 'newTable' already exists, createTable doesn't throw
        await serviceClient.createTable(tableName);
        this._tableClient = TableClient.fromConnectionString(AppSettings.ChangeTokenStorageConnectionString, tableName, { allowInsecureConnection: true });
        logger.trackTrace({
            message: 'Init success',
            properties: {
                source: this.LOG_SOURCE,
                method: "Init"
            },
            severity: "Verbose"
        });
    }
    public async GetChangeToken(listId: string): Promise<string> {
        let changeToken = `1;3;${listId};${this.toTicks(moment().subtract(1, 'hours'))};-1`
        try {
            if (!this._tableClient) {
                throw new Error("ChangeTokenService not initialized");
            }
            try{
                const entity = await this._tableClient.getEntity<{ changeToken: string }>(PARTITION_KEY, listId);
                if(entity?.changeToken){
                    changeToken = entity.changeToken;
                }
            }catch(error){
                if(error.statusCode !== 404){
                    throw error;
                }
            }
            return changeToken;
        } catch (err) {
            logger.trackException({
                exception: err,
                properties: { source: this.LOG_SOURCE, method: "GetChangeToken" }
            });
            throw new Error(err);
        }
    }
    public async SaveChangeToken(listId: string, changeToken: string): Promise<string> {
        try {
            if (!this._tableClient) {
                throw new Error("ChangeTokenService not initialized");
            }
            await this._tableClient.upsertEntity({ partitionKey:PARTITION_KEY, rowKey:listId, changeToken });
            return changeToken;
        } catch (err) {
            logger.trackException({
                exception: err,
                properties: { source: this.LOG_SOURCE, method: "SaveChangeToken" }
            });
            throw new Error(err);
        }
    }
    private toTicks(date: moment.Moment): number {
        return (date.valueOf() * 10000) + 621355968000000000;
    }
}
export const changeTokenService: IChangeTokenService = new ChangeTokenService();