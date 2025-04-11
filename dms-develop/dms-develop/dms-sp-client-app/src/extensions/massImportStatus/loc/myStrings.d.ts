declare interface IMassImportStatusFieldCustomizerStrings {
  UpdateStatusFailed: string;
  OperationFailed: string;
  OperationSuccessful: string;
  TimeoutError: string;
  ResetIdError: string;
  ErrorGettingStatus: string;
  UpdateJobErrorMessage: string;
  UpdateJobError:string;
  JobDetails: string;
  Processing: string;
  Filter: string;
  Requested: string;
  Queued: string;
  ErrorStartingImport:string;
  DisabledTooltip: string;
  MIApplicable: string;
  MIDraft: string;
  DML: string;
  StartImport: string;
  MIApplicableOperation: string;
  MIDraftOperation: string;
  DMLOperation: string;
}

declare module 'MassImportStatusFieldCustomizerStrings' {
  const strings: IMassImportStatusFieldCustomizerStrings;
  export = strings;
}
