import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { IWfTask } from "../interfaces/IWfTask";
import WfTaskMapper from "./WfTaskMapper";
//import { IWfTaskMapper } from "./IWfTaskMapper";
//import CONSTANTS from "../_Constants";
import { IDraftRequest, ISPDraftRequest } from "../interfaces/IDraftRequest";
import { IDraftRequestMapper } from "./IDraftRequestMapper";
import { IWfStatus } from "../interfaces/IWfStatus";



export class DraftRequest implements IDraftRequest {
    constructor(
        public tasks :IWfTask[],
        public wfStatus:IWfStatus[],
        public id: number,
        public author: { id: number; displayName: string; eMail: string; },
        public created: Date,
        public modified: Date,
        public workflowstatus: string,
        public documentstatus: string,           
        public documentId?:string,
        public document?: { id: number, reference: string }

    ) {

    }
}

export class DraftRequestMapper implements IDraftRequestMapper {

    public static readonly serviceKey: ServiceKey<IDraftRequestMapper> =
        ServiceKey.create<IDraftRequestMapper>('dms:IWfRequestMapper', DraftRequestMapper);
   // private _taskMapper: WfTaskMapper;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
           // this._taskMapper = 
            serviceScope.consume(WfTaskMapper.serviceKey);
        });
    }
    fromSPItem(item: ISPDraftRequest, tasks:IWfTask[], wfStatus: IWfStatus[]): IDraftRequest {
        return new DraftRequest(
            tasks,
            wfStatus,
            item.ID,//[CONSTANTS.DraftListFieldNames.ID],
            item.Author ? { id: item.Author.Id, displayName: item.Author.Title, eMail: item.Author.EMail } : {id:0,displayName:'',eMail:''},
            item.Created,//[CONSTANTS.DraftListFieldNames.Created],
            item.Modified,//[CONSTANTS.DraftListFieldNames.Modified],
            item.WorkflowStatus,//[CONSTANTS.DraftListFieldNames.WorkflowStatus],
            item.DocumentStatus,//[CONSTANTS.DraftListFieldNames.DocumentStatus],            
            item.OData__dlc_DocId//[CONSTANTS.DraftListFieldNames.DocumentId],
        );
    }
}