import { IDraftRequest, ISPDraftRequest } from "../interfaces/IDraftRequest";
import { IWfStatus } from "../interfaces/IWfStatus";
import {  IWfTask } from "../interfaces/IWfTask";

export interface IDraftRequestMapper {
    fromSPItem(item: ISPDraftRequest, task : IWfTask[], wfStatus: IWfStatus[]): IDraftRequest;
}