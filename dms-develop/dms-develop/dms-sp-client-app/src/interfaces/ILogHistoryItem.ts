export interface ILogHistoryItem {
    documentId: string;
    projectReference: string;
    projectRevision: string;
    userDisplayName: string;
    userEmail: string;
    date: string;
    actionType: "create"|"update"|"delete";
    changedFields?: IFieldChange[];
}
export interface IFieldChange {
    fieldName: string;
    oldValue: string;
    newValue: string;
    fieldTitle?:string;
    fieldType?:string;
    isHidden?:boolean;
    isReadOnly?:boolean;
}