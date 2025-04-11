
const LOG_SOURCE = "MassImportService";
const IMPORT_JOB_PARTITION_KEY = "ImportJob";
import { TableClient, TableEntityResult, TableServiceClient, TransactionAction } from "@azure/data-tables";
import { AppSettings } from "../AppSettings.js";
import { ServiceBusClient, ServiceBusSender, ServiceBusMessage } from "@azure/service-bus";
import { Workbook } from "exceljs";
import { SPFI } from "@pnp/sp";
import { IListItemFormUpdateValue, RenderListDataOptions } from '@pnp/sp/lists/index.js';
import ExcelJS from 'exceljs';
import "@pnp/sp/presets/all.js";
import { GraphFI } from "@pnp/graph";
import { logger } from "./Logger.js";
import { IImportDocumentJob, IImportJob } from "../model/mass-import/index.js";
import { SPUtilityService } from "./SPUtilityService.js";
import { IDmsEvent } from "../model/index.js";
import { IFieldInfo, IItem } from "@pnp/sp/presets/all.js";
const docIdFieldName = process.env.DOC_ID_FIELD_NAME || '_dlc_DocId';
const MAX_FILTER_LENGTH = 90;

export class MassImportService {
    private _tableServiceClient: TableServiceClient;
    private _jobTable: TableClient;
    private _sbSender: ServiceBusSender;
    private _sp: SPFI;
    private _sbAutoClassifySender: ServiceBusSender;
    private _graph: GraphFI;

    /**
     *
     */
    constructor(
    ) {
        this._tableServiceClient = TableServiceClient.fromConnectionString(AppSettings.MassImportTableConnectionString, {
            allowInsecureConnection: AppSettings.IsLocalEnvironment
        });
        const sbClient = new ServiceBusClient(AppSettings.ServiceBusConnectionString);
        this._sbSender = sbClient.createSender(AppSettings.MassImportTopicName);
        this._sbAutoClassifySender = sbClient.createSender(AppSettings.AutoClassifyTopicName);
    }
    public setGraph(graph: GraphFI): void {
        this._graph = graph;
    }

    public setSP(sp: SPFI): void {
        this._sp = sp;
    }

    public async ensureTable(): Promise<void> {
        try {
            await this._tableServiceClient.createTable(AppSettings.ImportJobTableName);
            this._jobTable = TableClient.fromConnectionString(AppSettings.MassImportTableConnectionString, AppSettings.ImportJobTableName, {
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

    private splitObjectArray<T>(array: T[], chunkSize: number): T[][] {
        const result: T[][] = [];
        for (let i = 0; i < array.length; i += chunkSize) {
          result.push(array.slice(i, i + chunkSize));
        }
        return result;
      }

    // public async createImportJob(importJob: IImportJob): Promise<void> {
    //     try {
    //         const messages = importJob.documents.filter(d => d.status !== 'Failed').map(doc => ({
    //             body: doc,
    //             contentType: "application/json",
    //             subject: 'newDocImport',
    //             sessionId: doc.parentJobId
    //         }) as ServiceBusMessage);
    //         if (messages.length > 0) {
    //             let batch = await this._sbSender.createMessageBatch();
    //             for (let i = 0; i < messages.length; i++) {
    //                 const message = messages[i];
    //                 logger.trackTrace({ message: `Adding message ${JSON.stringify(message)} to the batch`, properties: { source: LOG_SOURCE }, severity: 'Verbose' });
    //                 if (!batch.tryAddMessage(message)) {
    //                     await this._sbSender.sendMessages(batch);
    //                     batch = await this._sbSender.createMessageBatch();
    //                     if (!batch.tryAddMessage(message)) {
    //                         throw new Error("Message too big to fit in a batch");
    //                     }
    //                 }
    //             }
    //             await this._sbSender.sendMessages(batch);
    //         }
    //         const transactions: unknown[] = importJob.documents.map(doc => {
    //             return ["upsert", {
    //                 partitionKey: IMPORT_JOB_PARTITION_KEY,
    //                 docJobId: doc.docJobId,
    //                 rowKey: doc.id,
    //                 status: !doc.status || doc.status === 'None' ? 'Queued' : doc.status,
    //                 attempts: 0,
    //                 metadata: JSON.stringify(doc.metadata),
    //                 files: JSON.stringify(doc.files),
    //                 id: doc.id,
    //                 parentJobId: doc.parentJobId,
    //                 rowNumber: doc.rowNumber,
    //                 spStatus: doc.spStatus,
    //                 error: doc.error,
    //                 friendlyError: doc.friendlyError,
    //                 spId: doc.spId,
    //             } as IImportDocumentJob & { partitionKey: string; rowKey: string; files: string; metadata: string }];
    //         });
    //         let jobStatus = importJob.status || 'Queued';
    //         if (messages.length === 0) {
    //             jobStatus = 'Completed';
    //         }
    //         transactions.push(["upsert", {
    //             partitionKey: IMPORT_JOB_PARTITION_KEY,
    //             rowKey: importJob.id,
    //             attempts: importJob.attempts,
    //             status: jobStatus,
    //             fileServerRelativeUrl: importJob.fileServerRelativeUrl,
    //             id: importJob.id,
    //             jobId: importJob.jobId,
    //             requestorMail: importJob.requestorMail,
    //             siteUrl: importJob.siteUrl,
    //             error: importJob.error,
    //             type: importJob.type,
    //             friendlyError: importJob.friendlyError,
    //             documents: undefined,
    //             spId: importJob.spId,
    //             spLibraryId: importJob.spLibraryId
    //         } as IImportJob & { partitionKey: string; rowKey: string, documents: undefined }]);
    //         await this._jobTable.submitTransaction(transactions as TransactionAction[]);

    //         //upload result xlsx files in case of all rows validation error
    //         if (messages.length <= 0) {
    //             const updatedImportJob = await this.getImportJob(importJob.jobId);
    //             if (updatedImportJob.documents.every(d => d.status === 'Completed' || d.status === 'Failed')) {
    //                 await this.completeImportJob(updatedImportJob);
    //             }
    //         }
    //     } catch (error) {
    //         logger.trackException({
    //             exception: error,
    //             properties: { source: LOG_SOURCE }
    //         });
    //         throw new Error(error);
    //     }
    // }

    public async createImportJob(importJob: IImportJob): Promise<void> {
        try {
            const messages = importJob.documents.filter(d => d.status !== 'Failed').map(doc => ({
                body: doc,
                contentType: "application/json",
                subject: 'newDocImport',
                sessionId: doc.parentJobId
            }) as ServiceBusMessage);
            if (messages.length > 0) {
                //Split largerjobs group into smaller group of messages which holds 99 on each group
                const chunkedMessages = this.splitObjectArray(messages, 50);
                //Process one group at a time
                for(let chunkCnt=0;chunkCnt<chunkedMessages.length;chunkCnt++)
                {
                    let batch = await this._sbSender.createMessageBatch();
                    for (let i = 0; i < chunkedMessages[chunkCnt].length; i++) {
                        const message = chunkedMessages[chunkCnt][i];
                        logger.trackTrace({ message: `Adding message ${JSON.stringify(message)} to the batch`, properties: { source: LOG_SOURCE }, severity: 'Verbose' });
                        if (!batch.tryAddMessage(message)) {
                            await this._sbSender.sendMessages(batch);
                            batch = await this._sbSender.createMessageBatch();
                            if (!batch.tryAddMessage(message)) {
                                throw new Error("Message too big to fit in a batch");
                            }
                        }
                    }
                    await this._sbSender.sendMessages(batch);
                }
 
            }
            const transactions: unknown[] = importJob.documents.map(doc => {
                return ["upsert", {
                    partitionKey: IMPORT_JOB_PARTITION_KEY,
                    docJobId: doc.docJobId,
                    rowKey: doc.id,
                    status: !doc.status || doc.status === 'None' ? 'Queued' : doc.status,
                    attempts: 0,
                    metadata: JSON.stringify(doc.metadata),
                    files: JSON.stringify(doc.files),
                    id: doc.id,
                    parentJobId: doc.parentJobId,
                    rowNumber: doc.rowNumber,
                    spStatus: doc.spStatus,
                    error: doc.error,
                    friendlyError: doc.friendlyError,
                    spId: doc.spId,
                } as IImportDocumentJob & { partitionKey: string; rowKey: string; files: string; metadata: string }];
            });
            let jobStatus = importJob.status || 'Queued';
            if (messages.length === 0) {
                jobStatus = 'Completed';
            }
            transactions.push(["upsert", {
                partitionKey: IMPORT_JOB_PARTITION_KEY,
                rowKey: importJob.id,
                attempts: importJob.attempts,
                status: jobStatus,
                fileServerRelativeUrl: importJob.fileServerRelativeUrl,
                id: importJob.id,
                jobId: importJob.jobId,
                requestorMail: importJob.requestorMail,
                siteUrl: importJob.siteUrl,
                error: importJob.error,
                type: importJob.type,
                friendlyError: importJob.friendlyError,
                documents: undefined,
                spId: importJob.spId,
                spLibraryId: importJob.spLibraryId
            } as IImportJob & { partitionKey: string; rowKey: string, documents: undefined }]);
           
            const chunkedMessages = this.splitObjectArray(transactions, 98);
            for(let chunkCnt=0;chunkCnt<chunkedMessages.length;chunkCnt++)
            await this._jobTable.submitTransaction(chunkedMessages[chunkCnt] as TransactionAction[]);
 
            //upload result xlsx files in case of all rows validation error
            if (messages.length <= 0) {
                const updatedImportJob = await this.getImportJob(importJob.jobId);
                if (updatedImportJob.documents.every(d => d.status === 'Completed' || d.status === 'Failed')) {
                    await this.completeImportJob(updatedImportJob);
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

    public async getImportJob(id: string): Promise<IImportJob> {
        try {
            logger.trackTrace({ message: `Getting job PK:${IMPORT_JOB_PARTITION_KEY} RK:${id}`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            const job = await this._jobTable.getEntity<IImportJob>(IMPORT_JOB_PARTITION_KEY, id);
            logger.trackTrace({ message: `Getting job documents filter: PartitionKey eq '${IMPORT_JOB_PARTITION_KEY}' and parentJobId eq '${id}'`, properties: { source: LOG_SOURCE }, severity: 'Information' });
            const docs = this._jobTable.listEntities<IImportDocumentJob & { metadata: string }>({
                queryOptions: { filter: `PartitionKey eq '${IMPORT_JOB_PARTITION_KEY}' and parentJobId eq '${id}'` }
            });
            job.documents = [];
            for await (const doc of docs) {
                job.documents.push({
                    ...doc,
                    metadata: doc.metadata ? JSON.parse(doc.metadata) : undefined
                });
            }
            return job;
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
    public async failImportJob(job: IImportJob, errorObj: any, lastAttempt: boolean, attempt: number): Promise<void> {
        logger.trackTrace({ message: `Failing job ${JSON.stringify(job || {})}`, severity: 'Information', properties: { source: LOG_SOURCE } });
        let error = '';
        let friendlyError = '';
        if (errorObj instanceof Error) {
            error = `${errorObj.message}\n${errorObj.stack}`;
            friendlyError = errorObj.message;
        } else if (typeof errorObj === 'string') {
            error = errorObj;
            friendlyError = errorObj;
        } else {
            error = JSON.stringify(errorObj);
            friendlyError = 'Unknown error';
        }
        const failedJob: IImportJob = {
            ...job, attempts: attempt,
            status: lastAttempt ? "Failed" : job.status,
            error: job.error ? `${job.error}\nAttempt ${attempt}\n${error}` : `${attempt}\n${error}`,
            friendlyError: job.friendlyError && !job.friendlyError?.includes(friendlyError) ? `${job.friendlyError}; ${friendlyError}` : friendlyError
        };
        await this._jobTable.upsertEntity({
            partitionKey: IMPORT_JOB_PARTITION_KEY,
            jobId: job.jobId,
            rowKey: job.id,
            documents: undefined,
            attempts: failedJob.attempts,
            status: failedJob.status,
            fileServerRelativeUrl: failedJob.fileServerRelativeUrl,
            error: failedJob.error,
            friendlyError: failedJob.friendlyError,
            type: failedJob.type,
            id: failedJob.id,
            requestorMail: failedJob.requestorMail,
            siteUrl: failedJob.siteUrl
        } as IImportJob & { partitionKey: string, rowKey: string });
        if (this._sp) {
            if (failedJob.spId && failedJob.spLibraryId) {
                await this._sp.web.lists.getById(job.spLibraryId).items.getById(job.spId).update({ Status: failedJob.status, Error: failedJob.friendlyError, ErrorDetails: failedJob.error });
            } else {
                logger.trackTrace({ message: `Could not update status of the item ${job.spId} in the library ${job.spLibraryId}`, properties: { source: LOG_SOURCE }, severity: 'Warning' });
            }
        }
        job.status = failedJob.status;
    }
    public async updateJob(importJob: IImportJob, newStatus: 'Queued' | 'Processing' | 'Completed' | 'Failed'): Promise<void> {
        this._jobTable.upsertEntity({
            partitionKey: IMPORT_JOB_PARTITION_KEY,
            rowKey: importJob.id,
            documents: undefined,
            status: newStatus,
            attempts: importJob.attempts,
            fileServerRelativeUrl: importJob.fileServerRelativeUrl,
            id: importJob.id,
            jobId: importJob.jobId,
            requestorMail: importJob.requestorMail,
            siteUrl: importJob.siteUrl,
            error: importJob.error,
            friendlyError: importJob.friendlyError,
            type: importJob.type,
            spId: importJob.spId,
            spLibraryId: importJob.spLibraryId
        } as IImportJob & { partitionKey: string, rowKey: string });
        if (newStatus !== importJob.status) {
            // await this._sp.web.lists.getById(importJob.spLibraryId).items.getById(importJob.spId).update({ Status: newStatus });
            const currStatus = await this._sp.web.lists.getById(importJob.spLibraryId).items.getById(importJob.spId).select("Status")();
            if (currStatus.Status !== newStatus)
                try {
                    await this._sp.web.lists.getById(importJob.spLibraryId).items.getById(importJob.spId).update({ Status: newStatus });
                }
                catch (error) {
                    if (error.status !== 409)
                        throw new Error(error)
                }

        }
        importJob.status = newStatus;
    }
    public async updateImportDocJob(importJob: IImportDocumentJob, newStatus: 'None' | 'Queued' | 'Processing' | 'Completed' | 'Failed'): Promise<void> {
        this._jobTable.upsertEntity({
            partitionKey: IMPORT_JOB_PARTITION_KEY,
            rowKey: importJob.id,
            status: newStatus,
            attempts: importJob.attempts,
            id: importJob.id,
            docJobId: importJob.docJobId,
            error: importJob.error,
            friendlyError: importJob.friendlyError,
            spId: importJob.spId,
            files: JSON.stringify(importJob.files),
            metadata: JSON.stringify(importJob.metadata),
            parentJobId: importJob.parentJobId,
            rowNumber: importJob.rowNumber,
            spStatus: importJob.spStatus
        } as IImportDocumentJob & { partitionKey: string, rowKey: string; files: string; metadata: string });
        importJob.status = newStatus;
    }
    public async completeImportJob(updatedImportJob: IImportJob): Promise<void> {
        await this._jobTable.upsertEntity({
            partitionKey: IMPORT_JOB_PARTITION_KEY,
            rowKey: updatedImportJob.id,
            documents: undefined,
            status: 'Completed',
            attempts: updatedImportJob.attempts,
            fileServerRelativeUrl: updatedImportJob.fileServerRelativeUrl,
            id: updatedImportJob.id,
            jobId: updatedImportJob.jobId,
            requestorMail: updatedImportJob.requestorMail,
            siteUrl: updatedImportJob.siteUrl,
            error: updatedImportJob.error,
            type: updatedImportJob.type,
            friendlyError: updatedImportJob.friendlyError,
        } as IImportJob & { partitionKey: string, rowKey: string });

        const originalFileName = updatedImportJob.fileServerRelativeUrl.split('/').pop();
        //Check if an attachment folder exists and is empty, delete if so
        try {
            const folderName = originalFileName.split('.').slice(0, -1).join('.');
            const folder = await this._sp.web.getFolderByServerRelativePath(updatedImportJob.fileServerRelativeUrl.replace(originalFileName, folderName)).select('Exists', 'ItemCount')();
            if (folder.Exists && folder.ItemCount === 0) {
                await this._sp.web.getFolderByServerRelativePath(updatedImportJob.fileServerRelativeUrl.replace(originalFileName, folderName)).delete();
            }
        } catch (error) {
            logger.trackException({
                exception: error,
                properties: { source: LOG_SOURCE }
            });
        }

        if ('Completed' !== updatedImportJob.status) {
            // await this._sp.web.lists.getById(updatedImportJob.spLibraryId).items.getById(updatedImportJob.spId).update({ Status: 'Completed' });
            const currStatus = await this._sp.web.lists.getById(updatedImportJob.spLibraryId).items.getById(updatedImportJob.spId).select("Status")();
            if (currStatus.Status !== "Completed")
                try {
                    await this._sp.web.lists.getById(updatedImportJob.spLibraryId).items.getById(updatedImportJob.spId).update({ Status: "Completed" });
                }
                catch (error) {
                    if (error.status !== 409)
                        throw new Error(error)
                }
        }
        updatedImportJob.status = 'Completed';
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet(`DML Import ${updatedImportJob.id}`);
        sheet.columns = [
            { header: 'ID', key: 'id' },
            { header: 'Parent Job ID', key: 'parentJobId' },
            { header: 'Status', key: 'status' },
            { header: 'Error', key: 'error' },
            { header: 'Attempts', key: 'attempts' },
            { header: 'Metadata', key: 'metadata' },
            { header: 'File Id', key: 'fileItemId' },
            { header: 'Row Number', key: 'rowNumber' },
            { header: 'Error Details', key: 'errorDetails' },

        ];
        updatedImportJob.documents.forEach(doc => {
            sheet.addRow({
                id: doc.id,
                parentJobId: doc.parentJobId,
                status: doc.status,
                errorDetails: doc.error,
                error: doc.friendlyError,
                attempts: doc.attempts,
                metadata: JSON.stringify(doc.metadata),
                filItemId: doc.spId,
                rowNumber: doc.rowNumber
            });
        });
        const buffer = await workbook.xlsx.writeBuffer();
        const resultFileName = originalFileName.replace('.xlsx', `_${updatedImportJob.id}_result.xlsx`);
        //await this._sp.web.lists.getByTitle("Batch Logs").rootFolder.files.addUsingPath(resultFileName, buffer, { Overwrite: true });
        const fileInfo = await this._sp.web.lists.getByTitle("Batch Logs").rootFolder.files.addUsingPath(resultFileName, buffer, { Overwrite: true });
        const item = await this._sp.web.getFileByServerRelativePath(fileInfo.ServerRelativeUrl).getItem();
        await item.update({ BatchName: originalFileName});
    }

    public async processImportDocJob(importJob: IImportJob, docJob: IImportDocumentJob): Promise<IItem> {
        const spUtil = new SPUtilityService(this._sp);
        const targetLib = await spUtil.getListFromName(AppSettings.DraftDocsLibName, 'Id,RootFolder/ServerRelativeUrl', 'RootFolder', true);

        for (const libName of [AppSettings.DraftDocsLibName, AppSettings.ApplicableDocsLibName]) {
            const lib = await spUtil.getListFromName(libName, 'Id,RootFolder/ServerRelativeUrl', 'RootFolder', true);
            const filter = AppSettings.ImportDocumentKeyFields[importJob.type].map(field => `${AppSettings.ImportFieldsMap[field].name} eq '${docJob.metadata[field]}'`).join(' and ');
            logger.trackTrace({ message: `Checking if document set exists in the library ${libName}. Filter: ${filter}`, properties: { source: LOG_SOURCE }, severity: 'Verbose' });
            const existingItems = await this._sp.web.lists.getById(lib.Id).items.filter(filter).select('Id')();
            if (
                existingItems.length > 0 &&
                (
                    docJob.attempts === 1 ||
                    !docJob.error ||
                    docJob.error?.includes('Document Set already exists in the library') === true
                )
            ) {
                throw new Error(`Document Set already exists in the library. Id: ${existingItems[0].Id}`);
            }
        };
        const item = await this._createDocSet(importJob, docJob, targetLib.Id);
        try {
            await this._uploadFilesToDocSet(importJob, docJob, targetLib.RootFolder.ServerRelativeUrl, 'Draft');
        } catch (error) {
            //try to clean up the created doc set
            try {
                const itemWithId = await item.select('Id')();
                await this._sp.web.lists.getById(targetLib.Id).items.getById(itemWithId.Id).delete();
            } catch (error1) {
                throw new Error(`An error occurred while rolling back the change: ${error}\n${error1}`);
            }
            throw new Error(error);
        }
        return item;
    }

    public async publishDocumentChangeEvent(listId: string, itemId: number, userMail: string, action: 'create' | 'update' | 'delete', overrideSessionId?: string): Promise<void> {
        const viewFieldsXml = [...AppSettings.MessageIncludedFields, docIdFieldName].reduce((acc, cur) => {
            return `${acc}<FieldRef Name='${cur}' />`;
        }, '');
        const additionalData = await this._sp.web.lists.getById(listId).items
            .getById(itemId).select('Editor/EMail', 'Id', 'Modified', 'ContentType/Name').expand("Editor", "ContentType")()
            .then(i => ({ id: i.Id, editor: i.Editor.EMail, modified: i.Modified, contentType: i.ContentType.Name }));
        const itemsData = (await this._sp.web.lists.getById(listId).renderListDataAsStream({
            ViewXml: `<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>${itemId}</Value></Eq></Where></Query><ViewFields>${viewFieldsXml}<FieldRef Name='ID' /></ViewFields><RowLimit>1</RowLimit></View>`,
            RenderOptions: RenderListDataOptions.ListData,
        })).Row[0];
        if (!itemsData || !additionalData) {
            throw new Error(`Error getting data for the item ${itemId} in the list ${listId}.`);
        }
        const siteServerRelariveUrl = await this._sp.web.select("ServerRelativeUrl")();
        const message: IDmsEvent = {
            fromBackend: true,
            siteUrl: siteServerRelariveUrl.ServerRelativeUrl,
            siteServerRelativeUrl: siteServerRelariveUrl.ServerRelativeUrl,
            libraryId: listId,
            itemId: itemId,
            user: userMail,
            modified: additionalData.modified,
            spContentType: additionalData.contentType,
            action,
            fields: [...AppSettings.MessageIncludedFields, docIdFieldName].reduce((acc, cur) => {
                acc[cur] = itemsData[cur];
                return acc;
            }, {})
        };
        const applicationProperties = {
            eventType: message.action,
            user: message.user,
            spContentType: message.spContentType,
            uniqueId: message.fields[docIdFieldName]
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
        await this._sbAutoClassifySender.sendMessages({
            contentType: "application/json",
            subject: 'SPEvent',
            body: message,
            sessionId: overrideSessionId || message.fields[docIdFieldName],
            applicationProperties
        } as ServiceBusMessage);
    }

    private async _uploadFilesToDocSet(importJob: IImportJob, docJob: IImportDocumentJob, libraryUrl: string, libName: string) {
        const folderName = AppSettings.ImportDocumentKeyFields[importJob.type].map(field => {
            if (!docJob.metadata[field]) throw new Error(`Field ${field} is required`);
            return docJob.metadata[field];
        }).join('-');

        const attachFile = docJob.metadata["Attachments"];
        if (attachFile) {
            const attachCount = attachFile.split(/\r?\n/).map(f => f.trim()).length;

            let attachDetails = [];
            AppSettings.ImportDocumentAttachmentField.map(k => {
                if (docJob.metadata[k]) {
                    const splittedData = docJob.metadata[k].split(/\r?\n/).map(f => f.trim());
                    if (attachCount !== splittedData.length)
                        throw new Error(`Error: Column "${k}" is not matching with the attachment files included in the file.`);
                    attachDetails[k] =
                    {
                        "Data": splittedData,
                        "Properties": AppSettings.ImportFieldsMap[k]
                    };
                } else {
                    logger.trackTrace({ message: `No data found for the column ${k} in the document ${docJob.id}`, properties: { source: LOG_SOURCE }, severity: 'Warning' });
                }

            });

            const attachments = [];
            //Get each file and its properties to update
            await Promise.all(attachDetails["Attachments"].Data.map(async (file, index) => {
                let valuesForUpdate: IListItemFormUpdateValue[] = [];
                valuesForUpdate = await Promise.all(AppSettings.ImportDocumentAttachmentField.map(fields => {
                    if (fields !== "Attachments") {
                        if (!this.isNullOrEmpty(attachDetails[fields]?.Data[index])) {
                            const fieldData = attachDetails[fields];
                            return this._getValueForSP(fieldData.Data[index], fieldData.Properties.type, fieldData.Properties.name, libName).then(value => {
                                return { FieldName: fieldData.Properties.name, FieldValue: value };
                            });
                        }
                    }
                }))
                attachments.push({
                    "FileUrl": file,
                    "UpdateProperties": valuesForUpdate.filter(f => f)
                });
            })
            )
            const fileUrl = [];

            // Get folder name where files are exists
            const originalFileName = importJob.fileServerRelativeUrl.split('/').pop();
            const fileFolderName = originalFileName.replace('.xlsx', ``);
            const sourceFolderPath = importJob.fileServerRelativeUrl.replace(originalFileName, fileFolderName);

            //Check source folder exists
            const sourceFolder = await this._sp.web.getFolderByServerRelativePath(sourceFolderPath).select('Exists')();
            if (!sourceFolder.Exists)
                throw new Error(`Could not find the folder ("${sourceFolderPath}"), with the uploaded filename `);

            //Check all files are exists
            await Promise.all(attachments.map(async (file) => {
                if (file.FileUrl.length > 2) {
                    const sourceFile = await this._sp.web.getFileByServerRelativePath(sourceFolderPath + "/" + file.FileUrl).select('Exists')().then((d) => d.Exists).catch(() => false);
                    if (!sourceFile)
                        fileUrl.push(file.FileUrl);
                }
            })
            );
            if (fileUrl.length > 0)
                throw new Error(`Could not find the file(s), [ ${fileUrl.join('; ')} ]. `);

            //Make sure Doc set is created and ready for file uploads
            const docSetPath = libraryUrl + "/" + folderName;
            const destFolder = await this._sp.web.getFolderByServerRelativePath(docSetPath).select('Exists')();
            if (!destFolder.Exists)
                throw new Error(`Could not find the document set to upload the files, [ ${docSetPath} ]. `);

            await Promise.all(attachments.map(async (file) => {
                if (file.FileUrl.length > 2) {
                    const fileName = file.FileUrl.split('/').pop();
                    const fileItem = await this._sp.web.getFileByServerRelativePath(sourceFolderPath + "/" + file.FileUrl).moveByPath(docSetPath + "/" + fileName, false, true);

                    const item = await fileItem.getItem();
                    try {
                        await item.validateUpdateListItem(file.UpdateProperties);
                    } catch (error1) {
                        throw new Error(error1);
                    }
                }
            })
            );
        }
    }

    private isNullOrEmpty(str: string | null | undefined): boolean {
        return str === null || str === undefined || str.trim() === "";
    }

    private async _createDocSet(importJob: IImportJob, docJob: IImportDocumentJob, targetLibId: string): Promise<IItem> {
        const valuesForUpdate: IListItemFormUpdateValue[] = await Promise.all(
            Object.keys(AppSettings.ImportFieldsMap)
                .filter(f => !AppSettings.ImportDocumentAttachmentField.includes(f))
                .filter(f => AppSettings.ImportFieldsMap[f].required?.[importJob.type])
                .map(key => {
                    return this._getValueForSP(docJob.metadata[key], AppSettings.ImportFieldsMap[key].type, AppSettings.ImportFieldsMap[key].name).then(value => {
                        return { FieldName: AppSettings.ImportFieldsMap[key].name, FieldValue: value };
                    });
                })
        );
        const fields: { InternalName: string; TypeAsString: string; }[] = await this._sp.web.contentTypes.getById(AppSettings.DocSetContentTypeId).fields.select('InternalName', 'TypeAsString')();
        //Add default values
        const defaultValues = await Promise.all(
            Object.keys(docJob.metadata)
                .map(excelName => ({ spName: AppSettings.ImportFieldsMap[excelName]?.name || excelName, excelName }))
                .filter(f => !AppSettings.ImportDocumentAttachmentField.includes(f.excelName))
                .filter(f => !valuesForUpdate.some(v => v.FieldName === f.spName) && docJob.metadata[f.excelName])
                .map(f =>
                    this._getValueForSP(
                        docJob.metadata[f.excelName],
                        fields.find(field => field.InternalName === f.spName)?.TypeAsString,
                        f.spName
                    ).then(value => ({ FieldName: f.spName, FieldValue: value }))
                )
        );
        valuesForUpdate.push(...defaultValues);

        let addOrUpdateResult: IListItemFormUpdateValue[] | undefined = undefined;
        const folderName = AppSettings.ImportDocumentKeyFields[importJob.type].map(field => {
            if (!docJob.metadata[field]) throw new Error(`Field ${field} is required`);
            return docJob.metadata[field];
        }).join('-');
        const contentType = await this._sp.web.lists.getById(targetLibId).contentTypes.select('StringId').filter(`startswith(StringId, '${AppSettings.DocSetContentTypeId}')`)();
        if (contentType.length === 0) {
            throw new Error(`Content Type not found for the library ${targetLibId}`);
        }
        await this._sp.web.lists.getById(targetLibId).rootFolder.folders.addUsingPath(folderName);
        let folderItem: IItem;
        try {
            folderItem = await this._sp.web.lists.getById(targetLibId).rootFolder.folders.getByUrl(folderName).getItem();
            logger.trackTrace({ message: `Updating fields for the document set ${folderName}. Values: ${JSON.stringify(valuesForUpdate)}`, properties: { source: LOG_SOURCE }, severity: 'Verbose' });
            addOrUpdateResult = await folderItem.validateUpdateListItem(valuesForUpdate.concat(
                { FieldName: 'ContentTypeId', FieldValue: contentType[0].StringId }
            ));
        } catch (error1) {
            try {
                await this._sp.web.lists.getById(targetLibId).rootFolder.folders.getByUrl(folderName)
                    .deleteWithParams({ BypassSharedLock: true, DeleteIfEmpty: false });
            } catch (error2) {
                throw new Error(`${error1}\n${error2}`);
            }
            throw new Error(error1);
        }
        if (addOrUpdateResult) {
            const erroredFields = addOrUpdateResult.filter(r => r.HasException).map(r => r.FieldName).join(', ');
            if (erroredFields) {
                try {
                    await this._sp.web.lists.getById(targetLibId).rootFolder.folders.getByUrl(folderName)
                        .deleteWithParams({ BypassSharedLock: true, DeleteIfEmpty: false });
                } catch (error) {
                    throw new Error(`${JSON.stringify(erroredFields)}\n${error}`);
                }
                throw new Error(`Error updating fields: ${JSON.stringify(addOrUpdateResult.filter(r => r.HasException))}`);
            }
        } else {
            throw new Error('Unknown Error updating fields');
        }
        return folderItem;
    }


    private async getContentTypeId(fieldValue: string, libName: string) {
        const allContentTypes1 = (await this._sp.web.lists.getByTitle(libName).contentTypes()).filter(ct => ct.Name === fieldValue);
        if (allContentTypes1.length > 0) return allContentTypes1[0].StringId

        throw new Error(`Error: No Content Type with name "${fieldValue}" found in the SP List ${libName}, while trying to update Attachment fields`);
    }


    private async getTermValues(fieldName: string, fieldValue: string, spUtil: SPUtilityService) {
        const termSetId = await this._sp.web.lists.getByTitle("Draft").fields.getByInternalNameOrTitle(fieldName)();
        const infoByTermSetId1 = await this._graph.termStore.sets.getById(termSetId["TermSetId"]).terms();
        const terms = infoByTermSetId1.filter(term => { return term.labels.some(lab => lab.name === fieldValue) });
        if (terms.length > 0) {
            return `${fieldValue}|${terms[0].id}`;
        }
        throw new Error(`Error in parsing value for Taxonomy field ${fieldName} with value ${fieldValue}`);
    }

    private async _getValueForSP(value: string, type: string, name: string, libName?: string): Promise<string> {
        if (!value) return '';
        if ((value as unknown as { hyperlink: string }).hyperlink) {
            value = (value as unknown as { text: string }).text;
        }
        const spUtil = new SPUtilityService(this._sp);
        try {
            switch (type) {
                case 'User':
                    const user = await this._sp.web.ensureUser(value);
                    return JSON.stringify([{ Key: user.LoginName }]);
                case 'UserMulti':
                    const result = await Promise.all(value.split(';').map(v => this._sp.web.ensureUser(v.trim())));
                    return JSON.stringify(result.map(v => ({ Key: v.LoginName })));
                case 'TaxonomyFieldType':
                    return await this.getTermValues(name, value, spUtil);
                case 'TaxonomyFieldTypeMulti':
                    let returnValue = [];
                    await Promise.all(value.split(",").map(async termValue => {
                        returnValue.push(`${await this.getTermValues(name, termValue, spUtil)}`);
                    }));
                    return returnValue.join(';');
                case 'DateTime':
                    if (!/(\d{2})\/(\d{2})\/(\d{4})/.test(value)) {
                        throw new Error(`Invalid date format: ${value}`);
                    }
                    const [day, month, year] = value.split('/');
                    return `${month}/${day}/${year}`;
                case 'MultiChoice':
                    return value.split(';').map(v => v.trim()).join('#;');
                case 'Boolean':
                    if (!libName)
                        return value;
                    if (value.toLowerCase() === "y" || value.toLowerCase() === "yes" || value.toLowerCase() === "true" || value.toLowerCase() === "1")
                        return "1";
                    else return "0";
                case 'ContentType':
                    if (libName)
                        return await this.getContentTypeId(value, libName);
                    else return value
                default:
                    return value;
            }
        } catch (error) {
            throw new Error(`Error parsing value for field ${name}: ${error}`);
        }

    }


    public async failImportDocJob(docJob: IImportDocumentJob, errorObj: any, lastAttempt: boolean, attempt: number): Promise<void> {
        let error = '';
        let friendlyError = '';
        if (errorObj instanceof Error) {
            error = `${errorObj.message}\n${errorObj.stack}`;
            friendlyError = errorObj.message;
        } else if (typeof errorObj === 'string') {
            error = errorObj;
            friendlyError = errorObj;
        } else {
            error = JSON.stringify(errorObj);
            friendlyError = 'Unknown error';
        }
        const failedDocJob: IImportDocumentJob = {
            ...docJob, attempts: attempt,
            status: lastAttempt ? "Failed" : docJob.status,
            error: docJob.error ? `${docJob.error}\nAttempt ${attempt}\n${error}` : `${attempt}\n${error}`,
            friendlyError: docJob.friendlyError && !docJob.friendlyError?.includes(friendlyError) ? `${docJob.friendlyError}; ${friendlyError}` : friendlyError
        };
        await this._jobTable.upsertEntity({
            partitionKey: IMPORT_JOB_PARTITION_KEY,
            rowKey: docJob.id,
            status: failedDocJob.status,
            attempts: failedDocJob.attempts,
            error: failedDocJob.error,
            friendlyError: failedDocJob.friendlyError,
            id: failedDocJob.id,
            docJobId: failedDocJob.docJobId,
            parentJobId: failedDocJob.parentJobId,
            metadata: JSON.stringify(failedDocJob.metadata),
            spStatus: failedDocJob.spStatus,
            rowNumber: failedDocJob.rowNumber
        } as IImportDocumentJob & { partitionKey: string; rowKey: string; metadata: string });
    }

    private _getAvailableFieldsMap(workbook: ExcelJS.Workbook, job: IImportJob): { [key: string]: number } {
        if (workbook.worksheets.length === 0) {
            throw new Error('No worksheets found in the workbook');
        }
        const worksheet = workbook.worksheets[0];
        const importColNames = Object.keys(AppSettings.ImportFieldsMap);
        const colIndexMap: { [key: string]: number } = {};
        worksheet.getRow(1).eachCell((cell, colNumber) => {
            const colName = cell.text;
            if (importColNames.includes(colName)) {
                colIndexMap[colName] = colNumber;
                importColNames.splice(importColNames.indexOf(colName), 1);
            }
        });
        logger.trackTrace({
            message: `Column indexes: ${JSON.stringify(colIndexMap)}`,
            properties: { source: LOG_SOURCE },
            severity: 'Verbose'
        });
        if (importColNames.some(col => AppSettings.ImportFieldsMap[col].required?.[job.type])) {
            throw new Error(`Required columns for import type ${job.type} not found in the file: ${importColNames.filter(col => AppSettings.ImportFieldsMap[col].required?.[job.type]).join(',')}`);
        }
        return Object.keys(AppSettings.ImportFieldsMap).reduce((acc, key) => {
            if (colIndexMap[key])
                acc[key] = colIndexMap[key];
            return acc;
        }, {});
    }

    public async validateDocImportJobs(job: IImportJob): Promise<void> {
        const fieldNames = job.documents.map(doc => Object.keys(doc.metadata)).flat().filter((v, i, a) => a.indexOf(v) === i && !AppSettings.ImportDocumentAttachmentField.includes(v));
        const fieldsFilterArray = fieldNames.map(f => `(InternalName eq '${AppSettings.ImportFieldsMap[f]?.name}')`);
        // const fieldsFilterArray = fieldNames.map(f => `InternalName eq '${f}'`).join(' or ');
        const fields: IFieldInfo[] = [];
        while (fieldsFilterArray.length > 0) {
            let currentFilter: string[] = [];
            while (fieldsFilterArray.length > 0 && currentFilter.join(' or ').length < MAX_FILTER_LENGTH) {
                currentFilter.push(fieldsFilterArray.pop());
            }
            if (currentFilter.length === 0) {
                continue;
            }
            fields.push(...(await this._sp.web.fields.filter(currentFilter.join(' or '))
                .select('Title,InternalName,Required,DefaultValue,TypeAsString,Choices')()));
        }
        job.documents.forEach((doc, index: number) => {
            doc.error = '';
            doc.friendlyError = '';
            //Validate fields from SP parameters
            fieldNames.forEach(f => {
                const field = fields.find(field => field.InternalName === AppSettings.ImportFieldsMap[f]?.name);
                if (field) {
                    if (field.Required && !doc.metadata[f]) {
                        doc.error += `Missing value for required field ${field.Title}. `;
                        doc.friendlyError += `Missing value for required field ${field.Title}. `;
                        doc.status = 'Failed';
                        return;
                    }
                    if (doc.metadata[f] && field.TypeAsString === 'Choice' && !field.Choices?.includes(doc.metadata[f])) {
                        doc.error += `Value ${doc.metadata[f]} is not a valid choice for field ${field.Title}`;
                        doc.friendlyError += `Value ${doc.metadata[f]} is not a valid choice for field ${field.Title}`;
                        doc.status = 'Failed';
                        return;
                    }
                }
            });
            //Check if duplicate
            const key = AppSettings.ImportDocumentKeyFields[job.type].map(f => doc.metadata[f]).join('_');
            if (job.documents.findIndex((d, i) => i !== index && AppSettings.ImportDocumentKeyFields[job.type].map(f => d.metadata[f]).join('_') === key) > -1) {
                doc.error += `Duplicate document found with key ${key}. `;
                doc.friendlyError += `Duplicate document found with key ${key}. `;
                doc.status = 'Failed';
            }
            //If job type is MIApplicable, allow only one revision per reference
            if (job.type === 'MIApplicable') {
                const ref = doc.metadata[AppSettings.ImportDocumentKeyFields[job.type][0]];
                if (job.documents.filter(d => d.metadata[AppSettings.ImportDocumentKeyFields[job.type][0]] === ref).length > 1) {
                    doc.error += `Multiple revisions found for the reference ${ref}. `;
                    doc.friendlyError += `Multiple revisions found for the reference ${ref}. `;
                    doc.status = 'Failed';
                }
            }
        });
    }

    public async applyDefaultsToDocImportJobs(job: IImportJob): Promise<void> {
        const [batch, execute] = this._sp.batched();
        const [fieldLinks, fields] = await Promise.all([
            batch.web.contentTypes.getById(AppSettings.DocSetContentTypeId).fieldLinks(),
            batch.web.contentTypes.getById(AppSettings.DocSetContentTypeId).fields(),
            execute()
        ]);
        //const fieldNames = job.documents.map(doc => Object.keys(doc.metadata)).flat().filter((v, i, a) => a.indexOf(v) === i && !AppSettings.ImportDocumentAttachmentField.includes(v));
        job.documents.forEach((doc) => {
            fieldLinks.forEach(f => {
                const field = fields.find(field => field.InternalName === f.Name);
                const excelName = Object.keys(AppSettings.ImportFieldsMap).find(key => AppSettings.ImportFieldsMap[key].name === f.Name) || f.Name;
                if (field) {
                    if ((field.DefaultValue && !doc.metadata[excelName]) && (
                        field.TypeAsString === 'Text' || field.TypeAsString === 'Note' || field.TypeAsString === 'Choice' || field.TypeAsString === 'Number'
                    )) {
                        doc.metadata[excelName] = field.DefaultValue;
                    }
                }
            });
        });
    }

    public getDocumentImportJobs(workbook: Workbook, job: IImportJob, fileStatus: 'Draft' | 'Applicable'): IImportDocumentJob[] {
        const allAvailableFieldsMap = this._getAvailableFieldsMap(workbook, job);
        const availableFieldsMap = Object.keys(allAvailableFieldsMap).reduce((acc, key) => {
            if (AppSettings.ImportFieldsMap[key].required?.[job.type]) {
                acc[key] = allAvailableFieldsMap[key];
            }
            return acc;
        }, {} as { [key: string]: number });
        if (workbook.worksheets.length === 0) {
            throw new Error('No worksheets found in the workbook');
        }
        const worksheet = workbook.worksheets[0];
        if (Object.keys(availableFieldsMap).length === 0) {
            throw new Error(`No known fields found in the file. Known fields: ${Object.keys(AppSettings.ImportFieldsMap).join(',')}`);
        }
        if (AppSettings.ImportDocumentKeyFields[job.type].some(field => !availableFieldsMap[field])) {
            throw new Error(`Key fields not found in the file: ${AppSettings.ImportDocumentKeyFields[job.type].filter(field => !availableFieldsMap[field]).join(',')}`);
        }
        return worksheet.getRows(2, worksheet.rowCount - 1).map(row => {
            const id = `${job.id}_${AppSettings.ImportDocumentKeyFields[job.type].map(field => {
                return row.getCell(availableFieldsMap[field]).text || `ROW${row.number}`;
            }).join('_')}`;
            const metadataObj = Object.keys(availableFieldsMap).reduce((acc, key) => {
                const value = row.getCell(availableFieldsMap[key]).value;
                acc[key] = AppSettings.ImportFieldsMap[key].transform ?
                    value?.toString().replace(
                        new RegExp(AppSettings.ImportFieldsMap[key].transform?.regex, 'g'),
                        AppSettings.ImportFieldsMap[key].transform?.replace) :
                    value;
                return acc;
            }, {});
            //TODO: Get field name from settings
            metadataObj[AppSettings.DocStatusExcelColumn] = fileStatus;
            return {
                id,
                docJobId: id,
                attempts: 0,
                status: 'None',
                parentJobId: job.id,
                spStatus: fileStatus,
                files: undefined,
                metadata: metadataObj,
                rowNumber: row.number
            }
        });
    }

    public async getImportJobs(azTableFilter?: string, limit?: number): Promise<TableEntityResult<IImportJob>[]> {
        const filters = [
            `PartitionKey eq '${IMPORT_JOB_PARTITION_KEY}'`,
            'spId ge 0'
        ];
        if (azTableFilter) {
            filters.push(azTableFilter);
        }
        const importJobsIterator = this._jobTable.listEntities<IImportJob>({
            queryOptions: { filter: filters.map(f => `(${f})`).join(' and ') }
        });
        let importJobs: TableEntityResult<IImportJob>[] = [];
        for await (const job of importJobsIterator) {
            importJobs.push(job);
        }
        if (limit) {
            importJobs = importJobs.slice(0, limit);
        }
        const docFilter = importJobs.map(job => `(parentJobId eq '${job.id}')`).join(' or ');
        const docJobsIterator = this._jobTable.listEntities<IImportDocumentJob>({
            queryOptions: { filter: `PartitionKey eq '${IMPORT_JOB_PARTITION_KEY}' and (${docFilter})` }
        });
        const docJobs: IImportDocumentJob[] = [];
        for await (const docJob of docJobsIterator) {
            docJobs.push(docJob);
        }
        importJobs.forEach(job => {
            job.documents = docJobs.filter(doc => doc.parentJobId === job.id);
        });
        return importJobs;
    }

    public cleanup(): void {
        this._sbAutoClassifySender && this._sbAutoClassifySender.close();
        this._sbSender && this._sbSender.close();
    }
}