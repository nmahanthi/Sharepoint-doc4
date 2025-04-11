declare interface IApplicableMenuCommandSetStrings {
  resultMessage: string;
  errorMessage: string;
  DialogTitle: string;
  ConfirmationMessage: string;
  existedMessage:string;
  Command1: string;
  Command2: string;
}

declare module 'ApplicableMenuCommandSetStrings' {
  const strings: IApplicableMenuCommandSetStrings;
  export = strings;
}
