import { ServiceKey } from "@microsoft/sp-core-library";
import { ISPWfTask, IWfTask } from "../interfaces/IWfTask";
import { IWfTaskMapper } from "./IWfTaskMapper";
//import CONSTANTS from "../_Constants";

export default class WfTaskMapper implements IWfTaskMapper {

    public static readonly serviceKey: ServiceKey<IWfTaskMapper> =
        ServiceKey.create<IWfTaskMapper>('ams:IWfTaskMapper', WfTaskMapper);


    fromSPItem(task: ISPWfTask): IWfTask {
        return {
            id: task.Id,//[CONSTANTS.TaskListFieldNames.ID],
            comments: task.TaskComments,//[CONSTANTS.TaskListFieldNames.TaskComments],
            created: task.Created,//[CONSTANTS.TaskListFieldNames.Created],
            modified: task.Modified,//[CONSTANTS.TaskListFieldNames.Modified],
            outcome: task.TaskRunningStatus as 'In Progress' | 'Completed' | 'Cancelled' | 'Delegated',//[CONSTANTS.TaskListFieldNames.TaskRunningStatus] as 'In Progress' | 'Completed' | 'Cancelled' | 'Delegated',
            startDate: task.AssignedDate0,//[CONSTANTS.TaskListFieldNames.AssignedDate],
            assignee:task.Assignee0?{
                displayName: task.Assignee0.Title,
                eMail: task.Assignee0.EMail,
                id: task.Assignee0.Id
            }:{displayName:'', eMail:'',id:0,},
            endDate:task.ClosedDate0,//[CONSTANTS.TaskListFieldNames.ClosedDate],
            taskAssigneeType:task.TaskAssigneeType,//[CONSTANTS.TaskListFieldNames.TaskAssigneeType],
            workflowResult:task.Result0,//[CONSTANTS.TaskListFieldNames.Result],
            nextTask:task.NextTaskId0,//[CONSTANTS.TaskListFieldNames.NextTaskId],
            isDelegatedTask:task.IsDelegatedTask,//[CONSTANTS.TaskListFieldNames.IsDelegatedTask],
            prevTaskId:task.PrevTaskId,//[CONSTANTS.TaskListFieldNames.PrevTaskId],
            instanceId:task.InstanceId1//[CONSTANTS.TaskListFieldNames.InstanceId]
        };
    }
}