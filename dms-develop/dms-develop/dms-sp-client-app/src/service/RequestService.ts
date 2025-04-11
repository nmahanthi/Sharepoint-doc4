import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { IRequestService } from './IRequestService';
import { SPFI } from '@pnp/sp';
import CONSTANTS from '../_Constants';
import { IDraftRequestMapper } from '../mapper/IDraftRequestMapper';
import { ISPWfTask, IWfTask } from '../interfaces/IWfTask';
import WfTaskMapper from '../mapper/WfTaskMapper';
import { IWfTaskMapper } from '../mapper/IWfTaskMapper';
import { PermissionKind } from '@pnp/sp/security';
import { IDraftRequest, ISPDraftRequest } from '../interfaces/IDraftRequest';
import { DraftRequestMapper } from '../mapper/DraftRequestMapper';
import { IWfStatusMapper } from '../mapper/IWfStatusMapper';
import WfStatusMapper from '../mapper/WfStatusMapper';
import { ISPWfStatus } from '../interfaces/IWfStatus';
import { ICancelWFDetails } from '../interfaces/IGlobalInterfaces';
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { HttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { DmsRole, IPermissionsService, PermissionsService } from './PermissionsService';
import { ListItemAccessor, ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';


export class RequestService implements IRequestService {

    //Create a ServiceKey which will be used to consume the service.
    public static readonly serviceKey: ServiceKey<IRequestService> =
        ServiceKey.create<IRequestService>('dms:IRequestService', RequestService);
    private _requestMapper: IDraftRequestMapper;//IWfRequestMapper;
    private _taskMapper: IWfTaskMapper;
    private _wfStatusMapper: IWfStatusMapper;
    private _permissionsService: IPermissionsService;
    private _globalRequestorDetails: { email: string | null } = { email: null };
    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            this._requestMapper = serviceScope.consume(DraftRequestMapper.serviceKey);
            this._taskMapper = serviceScope.consume(WfTaskMapper.serviceKey);
            this._wfStatusMapper = serviceScope.consume(WfStatusMapper.serviceKey);
            this._permissionsService = serviceScope.consume(PermissionsService.serviceKey);
        });
    }

    private _requestListId: string;
    private _getRequestListId(): string {
        if (!this._requestListId) {
            throw new Error('requestListId is null. Did you call configure?');
        }
        return this._requestListId;
    }

    private _sp: SPFI;
    private _getSP(): SPFI {
        if (!this._sp) {
            throw new Error('sp is null. Did you call configure?');
        }
        return this._sp;
    }

    private _context: ListViewCommandSetContext | { httpClient: HttpClient, currentItem: ListItemAccessor, siteAbsoluteUrl: string, listId: string };
    private _getConext(): ListViewCommandSetContext | { httpClient: HttpClient, currentItem: ListItemAccessor, siteAbsoluteUrl: string, listId: string } {
        if (!this._context) {
            throw new Error('sp is null. Did you call configure?');
        }
        return this._context;
    }

    public configure(requestListId: string, sp: SPFI, context: ListViewCommandSetContext | { httpClient: HttpClient, currentItem: ListItemAccessor, siteAbsoluteUrl: string, listId: string }): void {
        this._requestListId = requestListId;
        this._sp = sp;
        this._context = context;
    }


    public async getDraftDocumentRequest(requestId: number): Promise<IDraftRequest> {
        const { DocumentId, Created, Modified, ID, WorkflowStatus, DocumentStatus } = CONSTANTS.DraftListFieldNames;
        const sp = this._getSP();
        const listId = this._getRequestListId();
        const result: ISPDraftRequest = await sp.web.lists.getById(listId).items.getById(requestId)//();
            .select(`${ID}`, `${DocumentId}`, `${Created}`, `${Modified}`, 'Author/Id', 'Author/Title', `${WorkflowStatus}`, `${DocumentStatus}`)
            .expand('Author')();

        const taskResult = await this._getSPTasks(result.OData__dlc_DocId);
        const taskOutput = taskResult.map(this._taskMapper.fromSPItem);

        const wfStatusResult = await this._getSPWFStatus(result.OData__dlc_DocId);
        const wfStatusOutput = wfStatusResult.map(this._wfStatusMapper.fromSPItem);
        return this._requestMapper.fromSPItem(result, taskOutput, wfStatusOutput);
    }

    public async getRequestTasks(requestId: string): Promise<IWfTask[]> {
        return (await this._getSPTasks(requestId)).map(this._taskMapper.fromSPItem);
    }

    private async _getSPTasks(requestId: string): Promise<ISPWfTask[]> {
        const { ClosedDate, TaskAssigneeType, ID, TaskRunningStatus, AssignedDate, Result, TaskComments, IsDelegatedTask, PrevTaskId, RevisionId, NextTaskId, Created, Modified, InstanceId } = CONSTANTS.TaskListFieldNames;
        const sp = this._getSP();
        const result: ISPWfTask[] = await sp.web.lists.getByTitle(CONSTANTS.TaskListName).items
            .select(`${ID}`, `${RevisionId}`, `${Created}`, `${Modified}`, `${AssignedDate}`, `${ClosedDate}`, `${IsDelegatedTask}`, `${PrevTaskId}`, `${TaskAssigneeType}`, `${TaskRunningStatus}`, `${Result}`, `${TaskComments}`, `${InstanceId}`, `Assignee0/Id`, `Assignee0/EMail`, `Assignee0/Title`, `${NextTaskId}`)
            .expand('Assignee0')
            .filter(`${RevisionId} eq '${requestId}'`).orderBy(`${ID}`)();
        return result;
    }

    private async _getListId(listTitle: string): Promise<string> {
        const sp = this._getSP();
        const result = await sp.web.lists.getByTitle(listTitle)();
        return result.Id;
    }

    // private async _getSPWFStatus(requestId: string): Promise<ISPWfStatus[]> {
    //     const { WorkflowType, RevisionId, InstanceId, InstanceURL, Created, Modified, Id, WorkflowStatus, StardDate, ClosedDate } = CONSTANTS.WFStatusListFieldNames;
    //     const sp = this._getSP();
    //     const result: ISPWfStatus[] = await sp.web.lists.getByTitle(CONSTANTS.WFStatusListName).items
    //         .select(`${Id}`, `${RevisionId}`, `${Created}`, `${Modified}`, `${WorkflowType}`, `${InstanceId}`, `${InstanceURL}`, `${WorkflowStatus}`, `${ClosedDate}`, `${StardDate}`,`Requestor/Id`, `Requestor/EMail`, `Requestor/Title`)
    //         .expand('Requestor')
    //         .filter(`${RevisionId} eq '${requestId}'`).orderBy(`${Id}`, false)();
    //         this._globalRequestorDetails.email = result.length > 0 ? result[0]?.Requestor?.EMail : null;
    //     return result;
    // }

    private async _getSPWFStatus(requestId: string): Promise<ISPWfStatus[]> {
        const { WorkflowType, RevisionId, InstanceId, InstanceURL, Created, Modified, Id, WorkflowStatus, StardDate, ClosedDate } = CONSTANTS.WFStatusListFieldNames;
        const sp = this._getSP();
        const result: ISPWfStatus[] = await sp.web.lists.getByTitle(CONSTANTS.WFStatusListName).items
            .select(`${Id}`, `${RevisionId}`, `${Created}`, `${Modified}`, `${WorkflowType}`, `${InstanceId}`, `${InstanceURL}`, `${WorkflowStatus}`, `${ClosedDate}`, `${StardDate}`, `Requestor/Id`, `Requestor/EMail`, `Requestor/Title`)
            .expand('Requestor')
            .filter(`${RevisionId} eq '${requestId}'`).orderBy(`${Id}`, false)();
        this._globalRequestorDetails.email = result.length > 0 ? result[0]?.Requestor?.EMail : null;
        console.log(this._globalRequestorDetails.email);
        return result;
    }


    public async canCurrentUserCancel(requestId: number, currentUserLogin: string): Promise<boolean> {
        const sp = this._getSP();
        const isWebAdmin = await sp.web.currentUserHasPermissions(PermissionKind.ManagePermissions);
        const validUser = await this._permissionsService.currentUserHasRole(DmsRole.ProjectAdmin);
        const requestorEmail = this._globalRequestorDetails.email;
        let isCurruserequalsRequestor = requestorEmail?.toLowerCase() === currentUserLogin.toLowerCase();
        const canCancel = isWebAdmin || requestorEmail?.toLowerCase() === currentUserLogin.toLowerCase() || validUser;
        console.log("Can Cancel Result:", canCancel);
        return isWebAdmin || isCurruserequalsRequestor || validUser;
        //return isWebAdmin || issuerLogin.Author.UserName === currentUserLogin || validUser;
    }

    public async cancelWorkflow(requestIds: number[], reason: string): Promise<void> {
        const { TaskRunningStatus, ClosedDate, TaskComments } = CONSTANTS.TaskListFieldNames;
        const value = CONSTANTS.CancelledStatus;
        const update: { [key: string]: unknown } = {};
        update[TaskRunningStatus] = value;
        update[TaskComments] = reason;
        update[ClosedDate] = new Date();
        const sp = this._getSP();
        requestIds.map(async (_id) => {
            await sp.web.lists.getByTitle(CONSTANTS.TaskListName).items.getById(_id).update(update);
        })
    }

    public async triggerFlow(cancelDetail: ICancelWFDetails[], reason: string, workflowUrl: string): Promise<void> {

        const context = this._getConext();
        const taskListID = await this._getListId(CONSTANTS.TaskListName);
        const wfListID = await this._getListId(CONSTANTS.WFStatusListName);
        const postURL = workflowUrl;


        if (cancelDetail.length > 0) {
            const cancel = cancelDetail[0];

            //Update flow
            if ((context instanceof ListViewCommandSetContext && context.listView.list && context.listView.selectedRows && context.listView.selectedRows.length > 0) || 
                    (context && 'httpClient' in context && 'currentItem' in context && 'siteAbsoluteUrl' in context && 'listId' in context)) {
                let body: string = "";
                if (context instanceof ListViewCommandSetContext && context.listView.list && context.listView.selectedRows && context.listView.selectedRows.length > 0) {
                    body = JSON.stringify({
                        'docitemID': context.listView.selectedRows[0]?.getValueByName('ID').toString(),
                        'doclibraryID': context.listView.list.guid.toString(),
                        'siteURL': context.pageContext.web.absoluteUrl,
                        'docGUId': context.listView.selectedRows[0]?.getValueByName('_dlc_DocIdUrl').toString().substring(context.listView.selectedRows[0]?.getValueByName('_dlc_DocIdUrl').toString().indexOf('=') + 1),
                        'docTitle': context.listView.selectedRows[0]?.getValueByName('FileLeafRef').toString(),
                        'projReference': context.listView.selectedRows[0]?.getValueByName('ProjectReference').toString(),
                        'TaskListID': taskListID,
                        'itemLink': `${context.pageContext.web.absoluteUrl}${context.listView.selectedRows[0]?.getValueByName('FileRef').toString()}`,
                        'workflowListItemID': cancel.wfID,
                        'taskListItemID': cancel.taskIds,
                        'WorkflowListID': wfListID

                    });
                } else if(context && 'httpClient' in context && 'currentItem' in context && 'siteAbsoluteUrl' in context && 'listId' in context) {
                    body = JSON.stringify({
                        'docitemID': context.currentItem.getValueByName('ID').toString(),
                        'doclibraryID': context.listId,
                        'siteURL': context.siteAbsoluteUrl,
                        'docGUId': context.currentItem.getValueByName('_dlc_DocIdUrl').toString().substring(context.currentItem.getValueByName('_dlc_DocIdUrl').toString().indexOf('=') + 1),
                        'docTitle': context.currentItem.getValueByName('FileLeafRef').toString(),
                        'projReference': context.currentItem.getValueByName('ProjectReference').toString(),
                        'TaskListID': taskListID,
                        'itemLink': `${context.siteAbsoluteUrl}${context.currentItem.getValueByName('FileRef').toString()}`,
                        'workflowListItemID': cancel.wfID,
                        'taskListItemID': cancel.taskIds,
                        'WorkflowListID': wfListID

                    });
                }
                const requestHeaders: Headers = new Headers();
                requestHeaders.append('Content-type', 'application/json');
                const httpClientOptions: IHttpClientOptions = {
                    body: body,
                    headers: requestHeaders
                };
                const response = await context.httpClient
                    .post(postURL,
                        HttpClient.configurations.v1,
                        httpClientOptions);
                if (!response.ok) {
                    throw new Error("Unexpected Error");
                }
            } else {
                throw new Error('No items selected');
            }
        }


    }

}