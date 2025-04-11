import { ISPWfStatus, IWfStatus } from "../interfaces/IWfStatus";

export interface IWfStatusMapper {
    fromSPItem(task: ISPWfStatus): IWfStatus
}