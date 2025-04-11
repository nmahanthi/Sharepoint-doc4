import { SPFI } from "@pnp/sp";
import { IWfTask } from "../interfaces/IWfTask";
import { IDraftRequest } from "../interfaces/IDraftRequest";
//import {  HttpClientResponse } from '@microsoft/sp-http';
import { ICancelWFDetails } from "../interfaces/IGlobalInterfaces";


export interface IRequestService {
   
    getDraftDocumentRequest(requestId: number): Promise<IDraftRequest>;   
    getRequestTasks(RequestId: string): Promise<IWfTask[]>;
    getRequestTasks(RequestId: string): Promise<IWfTask[]>;
    triggerFlow(cancelDetail:ICancelWFDetails[], reason:string,workflowUrl:string): Promise<void> ;
    configure(requestListId:string,  sp:SPFI, context:any): void;
    cancelWorkflow(requestIds: number[], reason: string):Promise<void>;
    canCurrentUserCancel(requestId: number, currentUserLogin: string): Promise<boolean>;
}