export interface IImportJob {
    id: string;
    status: 'Queued' | 'Processing' | 'Completed' | 'Failed';
    documents: IImportDocumentJob[];
    error?: string;
    attempts: number;
    requestorMail: string;
    siteUrl: string;
    fileServerRelativeUrl: string;
    spId?:number;
    spLibraryId?:string;
    type: 'MIApplicable' | 'MIDraft' | 'DML';
}
export interface IImportDocumentJob {
    id: string;
    parentJobId: string;
    status: 'None' | 'Queued' | 'Processing' | 'Completed' | 'Failed';
    error?: string;
    attempts: number;
    spId?:number;
    spStatus: 'Applicable' | 'Draft';
    files: IImportFileJob[];
    metadata: { [key: string]: string };
    rowNumber: number;
}
export interface IImportFileJob {
    id: string;
    parentDocumentId: string;
    status: 'None' | 'Queued' | 'Processing' | 'Completed' | 'Failed';
    error?: string;
    spId?:number;
    metadata: { [key: string]: string };
}