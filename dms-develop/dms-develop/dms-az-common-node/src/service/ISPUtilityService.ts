import { IListInfo } from "@pnp/sp/lists";

export interface ISPUtilityService {
    getListFromName(listName: string, select?: string, expand?: string): Promise<IListInfo>;
}