import { ILogHistoryItem } from "./ILogHistoryItem";

export interface ILogHistoryService {
    getLogHistoryItems(documentId: string): Promise<ILogHistoryItem[]>;
}