export interface IDocumentChangeRecord {
    documentID:string;
    action:string;
    metadata:string;
    changes?:string;
    modifiedBy:string;
    projectReference:string;
    projectRevision:string;
    userName:string;
}