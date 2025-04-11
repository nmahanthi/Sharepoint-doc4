export interface ISPWfStatus {
    WorkflowType: string;
    Id: number;
    Created: Date;
    Modified: Date;
    RevisionId: string;
    InstanceURL: string;
    InstanceId0: string;
    WorkflowStatus1:string
    StartDate1:string
    CloseDate:string;
    Requestor:any;
   // Requestor:{Id: number; EMail: string; Title: string; };
}
export interface IWfStatus {
    workflowType: 'AVVA' | 'Proofreading';
    id: number;
    created: Date;
    modified: Date;
    revisionId: string;
    instanceURL: string;
    instanceId: string;
    workflowStatus:string;
    startDate:string
    closedDate:string;
    requestor:string;
    // requestor:{
    //     id: number;
    //     displayName: string;
    //     eMail: string;
    // };
}