declare interface ICreateRevisionCommandSetStrings {
  CreateRevisionDialogTitle: string;
  CreateRevisionValidatingMessage: string;
  CreateRevisionProcessingMessage: string;
  CreateRevisionSuccessMessage: string;
  CreateRevisionErrorDraftExistMessage: string;
  CreateRevisionErrorUnknownMessage:string;
  DoNotRefresh:string;
}

declare module 'CreateRevisionCommandSetStrings' {
  const strings: ICreateRevisionCommandSetStrings;
  export = strings;
}
