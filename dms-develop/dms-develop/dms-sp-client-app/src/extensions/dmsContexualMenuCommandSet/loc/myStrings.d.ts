declare interface IDmsContexualMenuCommandSetCommandSetStrings { 
  AvvaTitle:string;
  ProofReadingTitle:string;
  WfTimelineErrorMessage: string;
  NoWfMessage: string;  
  TimelineTitle: string;
  ViewHistoryTitle:string;  
  WfPanelTitle: string;
  WfHistoryColumnHeaderAssigneType: string;
  WfHistoryColumnHeaderCreated: string;
  WfHistoryColumnHeaderLastUpdated: string;
  WfHistoryColumnHeaderStatus: string;
  WfHistoryColumnHeaderRequestor: string;
  WfHistoryColumnHeaderComments: string;
  WfHistoryColumnHeaderDelegator:string;
  WfPanelConfirmCancelationMessageIsLoading:string;
  WfPanelConfirmCancelationMessageError:string;
  WfPanelConfirmCancelationMessageSuccess:string;
  WfPanelConfirmCancelationYes:string;
  WfPanelConfirmCancelationNo:string;
  WfPanelConfirmCancelationMessage:string;
  WfPanelCancelationReason: string;
  WfPanelConfirmCancelationTitle:string;
  ThreeSixtyTitle:string
}

declare module 'DmsContexualMenuCommandSetCommandSetStrings' {
  const strings: IDmsContexualMenuCommandSetCommandSetStrings;
  export = strings;
}
