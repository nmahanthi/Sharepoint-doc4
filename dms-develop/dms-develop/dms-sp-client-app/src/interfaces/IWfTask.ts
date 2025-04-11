export interface ISPWfTask {
    Assignee0: { Id: number; EMail: string; Title: string; };
    TaskAssigneeType:string;
    AssignedDate0: Date;
    ClosedDate0:Date
    TaskComments: string;
    TaskRunningStatus: string;
    Id: number;
    Created: Date;
    Modified: Date;
    Result0:string;
    NextTaskId0:string;
    IsDelegatedTask:boolean;
    PrevTaskId:string;
    InstanceId1:string;
}
export interface IWfTask {
    assignee: {
        id: number;
        displayName: string;
        eMail: string;
    };
    taskAssigneeType:string;
    startDate: Date;
    endDate: Date;
    comments: string;
    outcome: 'In Progress' | 'Completed' | 'Cancelled' | 'Delegated' | 'Rejected';
    id: number;
    created: Date;
    modified: Date;
    workflowResult:string;
    delegator?:string;
    nextTask:string;
    isDelegatedTask:boolean;
    prevTaskId:string;
    instanceId:string;
}