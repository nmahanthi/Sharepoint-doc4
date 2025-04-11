import { IWfStatus } from "./IWfStatus";
import { IWfTask } from "./IWfTask";

export interface IDraftRequest{   
    tasks: IWfTask[],
    wfStatus:IWfStatus[],
    document?: { id: number, reference: string },
    id:number;
    author: {
        id: number;
        displayName: string;
        eMail: string;
    };
    created:Date;
    modified:Date;
    workflowstatus: string;
    documentstatus: string;
    documentId?:string;
}

export interface ISPDraftRequest{
    ContentTypeId:string;   
    document?: { id: number, reference: string },
    ID:number;
    Author: { Id: number; EMail: string; Title: string; };
    Created: Date;
    Modified: Date;
    OData__dlc_DocId:string;
    DocumentStatus:string;
    WorkflowStatus: string;

}