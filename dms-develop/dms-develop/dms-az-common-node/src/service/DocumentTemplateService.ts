
const LOG_SOURCE = "DocumentTemplateService";
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
import { getGUID } from "@pnp/core/util.js";
import { pnpjs } from "./PnpjsService.js";
import { Readable } from "stream";
const docIdFieldName = process.env.DOC_ID_FIELD_NAME || '_dlc_DocId';
const MAX_FILTER_LENGTH = 90;

export class DocumentTemplateService {
    private _graph: GraphFI;
    private _sp: SPFI;

    constructor(
    ) {
    }
    public setGraph(graph: GraphFI): void {
        this._graph = graph;
    }

    public setSP(sp: SPFI): void {
        this._sp = sp;
    }

    async createNewContentType(requestBody) {
        const sharepointHostName = `${AppSettings.TenantName}.sharepoint.com`;
        const relUrl = await this._sp.web.select("ServerRelativeUrl")();
        const site = await this._graph.sites.getByUrl(sharepointHostName, relUrl.ServerRelativeUrl);
        const siteInfo = await site();
        
        const newGuid = getGUID().replace(/-/g, "").toLocaleUpperCase();
        const cId = `${AppSettings.NativeContentType}00${newGuid}`;
        const sampleContentType = {
          name: requestBody.displayName,
          description: requestBody.displayName,
          base: {
            name: AppSettings.NativeContentTypeName,
            id: AppSettings.NativeContentType
          },
          group: "DMS",
          id: cId
        };
        const ct = await this._graph.sites.getById(siteInfo.id).contentTypes.add(sampleContentType);

        const tempFile = await this._sp.web.getFolderByServerRelativePath(AppSettings.ContentTypeTemplateDocumentLibrary).files.addChunked(`${requestBody.displayName}.docx`, this.base64ToStream(requestBody.template), { progress: data => { console.log(`progress`); }, Overwrite: true });

        await this._sp.web.contentTypes.getById(ct.data.id).update({ DocumentTemplate: tempFile.ServerRelativeUrl});

        return {
          displayName: ct.data.name,
          id: ct.data.id
        };
      }

    public async updateContentType(requestBody: any, id: string) {
        /*
         addFile(file, folderPath)
            .then((result) => {
              console.log("File uploaded successfully:", result);
            })
            .catch((error) => {
              console.error("Error uploading file:", error);
            });
        */
        const ct = await this._sp.web.contentTypes.getById(id).update({ Name: requestBody.displayName });

        const tempFile = await this._sp.web.getFolderByServerRelativePath(AppSettings.ContentTypeTemplateDocumentLibrary).files.addChunked(`${requestBody.displayName}.docx`, this.base64ToStream(requestBody.template), { progress: data => { console.log(`progress`); }, Overwrite: true });

        await this._sp.web.contentTypes.getById(id).update({ DocumentTemplate: tempFile.ServerRelativeUrl});

        return {
            displayName: requestBody.displayName,
            id: id
        }
    }

    private base64ToStream(base64String: string) {
        const buffer = Buffer.from(base64String, 'base64');
        const stream = new Readable();
        stream.push(buffer);
        stream.push(null);
        return stream;
      }
      
    public async deleteContentType(name: string, id: string) {

        const ct = await this._sp.web.contentTypes.getById(id).delete();

    }
    public async getAllNativeContentTypes() {
      const contentTypes: any = await (await this._sp.web.contentTypes.select("Id","Name")()).filter(ct => ct.Id.StringValue.startsWith(AppSettings.NativeContentType));
      const result: any[] = [];
      contentTypes.map(ct => {
          result.push({ displayName: ct.Name, id: ct.Id.StringValue });
      })
      return result;
  }
}