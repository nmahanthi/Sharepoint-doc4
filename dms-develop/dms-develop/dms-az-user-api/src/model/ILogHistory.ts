export interface ILogHistoryItem {

    documentId: string;
    projectReference: string;
    projectRevision: string;
    userDisplayName: string;
    userEmail: string;
    Date: string;
    actionType: string;
    changedFields?: IFieldChange[]
}

export interface IFieldChange {
    fieldName: string;
    oldValue: string;
    newValue: string;
}