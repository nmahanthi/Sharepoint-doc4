import { ISPWfTask, IWfTask } from "../interfaces/IWfTask";

export interface IWfTaskMapper {
    fromSPItem(task: ISPWfTask): IWfTask
}