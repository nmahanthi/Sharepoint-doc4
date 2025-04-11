export interface IMassImportRequest {
    fileServerRelativeUrl: string;
    id: string;
    siteUrl: string;
    spId: number;
    spListId: string;
    type: 'DML' | 'MIDraft' | 'MIApplicable';
  }