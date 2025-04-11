declare interface INotificationsContainerApplicationCustomizerStrings {
  ShowDetails: ReactNode;
  OperationInProgress: string;
  CloseNotification: string;
  OperationFailed: string;
  OperationCompletedSuccessfully: string;
}

declare module 'NotificationsContainerApplicationCustomizerStrings' {
  const strings: INotificationsContainerApplicationCustomizerStrings;
  export = strings;
}
