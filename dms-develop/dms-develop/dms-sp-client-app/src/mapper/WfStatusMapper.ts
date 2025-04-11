import { ServiceKey } from "@microsoft/sp-core-library";
import { ISPWfStatus, IWfStatus } from "../interfaces/IWfStatus";
import { IWfStatusMapper } from "./IWfStatusMapper";
//import CONSTANTS from "../_Constants";

export default class WfStatusMapper implements IWfStatusMapper {

    public static readonly serviceKey: ServiceKey<IWfStatusMapper> =
        ServiceKey.create<IWfStatusMapper>('ams:IWfStatusMapper', WfStatusMapper);


    fromSPItem(task: ISPWfStatus): IWfStatus {
        return {
            // requestor:task.Requestor?{
            //     displayName: task.Requestor.Title,
            //     eMail: task.Requestor.EMail,
            //     id: task.Requestor.Id
            // }:{displayName:'', eMail:'',id:0,},
            id: task.Id,//[CONSTANTS.WFStatusListFieldNames.Id],
            created: task.Created,//[CONSTANTS.WFStatusListFieldNames.Created],
            modified: task.Modified,//[CONSTANTS.WFStatusListFieldNames.Modified],
            workflowType: task.WorkflowType as 'AVVA' | 'Proofreading',//[CONSTANTS.WFStatusListFieldNames.WorkflowType] as 'AVVA' | 'Proofreading',            
            instanceURL:task.InstanceURL,//[CONSTANTS.WFStatusListFieldNames.InstanceURL],
            revisionId:task.RevisionId,//[CONSTANTS.WFStatusListFieldNames.RevisionId],
            instanceId:task.InstanceId0,//[CONSTANTS.WFStatusListFieldNames.InstanceId],
            workflowStatus:task.WorkflowStatus1,//[CONSTANTS.WFStatusListFieldNames.WorkflowStatus],
            closedDate:task.CloseDate,//[CONSTANTS.WFStatusListFieldNames.ClosedDate],
            startDate:task.StartDate1,//[CONSTANTS.WFStatusListFieldNames.StardDate]
            requestor:task.Requestor?.Title
        };
    }
}