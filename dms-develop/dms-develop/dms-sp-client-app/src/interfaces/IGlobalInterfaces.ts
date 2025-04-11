/* eslint-disable @rushstack/no-new-null */
import { IWorkflowTimelineStep } from "../components/WorkflowTimeline/WorkflowTimeline";
import { IWfTask } from "./IWfTask";

export interface IHistoryTimeLine{   
    tasks: IWfTask[],
    created:string;
    modified:Date;
    workflowstatus: string;
    requestor: string;
    steps:IWorkflowTimelineStep[];
    id:number;
}

export interface ICancelWFDetails{   
    wfID: number;
    taskIds:number[]
}

/* Baseline starts /*

/* eslint-disable @rushstack/no-new-null */

export interface ICustomMessage {
    target?: string;
    type?: ('error' | 'success' | 'info');
    targetFields?: IRequiredFields[];
    message: string;
  }
  
  export interface ICustomError {
    target: string;
    message: string;
  }
  
  export interface IRequiredFields {
    stateName: string;
    fieldName: string;
    fields?: string[];
    index?: number;
  }
  
  export interface ISPColumntateName {
    stateName: string;
  }
  
  export interface IValidationFields {
    stateName: string;
    fieldName: string;
    expression: RegExp;
    message: string;
  }
  
  export interface IReactSelectOption {
    value: string;
    label: string;
  }
  
  export interface ICustomSearchResults {
    [x: string]: any;
    DocumentId?: string; 
    DocId?: string;
    ItemId?: number;
    ProjectReference?: string;
    ProjectRevision?: string;
    CRID?: string;
    ChangeRequestT?: string;
    Code?: string;
    Title?: string;
    SitePath?: string;
    isSelected:boolean;
    ContentType?:string;
    Label?: string;
    BaselineStatus?: string;
    index:number;
    DocumentStatus?:string;
    DocumentLink?:string;
  }
  
  export interface ILinkedItems {
    DocumentId?: string | null; 
    ItemId?: number;
    SiteUrl?: string;
    ChildBaselineId?:number | null;
    isSelected?:boolean;
    ContentType?:string;
    Title?:string;    
    Label?: string;
    index:number;
    Status?:string;
    ProjectReference?:string;
    ProjectRevision?:string;
    BaselineStatus?:string;
    CRID?:string;
  }
  
export interface ISPType {
    title: string;
    id: string;
  }

  export interface ISPColumn {
    group: string;
    hidden: boolean;
    internalName: string;
    title: string;
    typeAsString: string;
    id: string;
  }
  export interface ISearchParam {
    field?:ISPColumn;
    condition?:string;
    operator?:string
    searchText?:string;
    searchFieldName?:string;
  }


  export interface IBaseLineExportEvent {
    ExportDocuments?:IRelatedBaselines;
    UserEmail?: string;
    SiteUrl?: string;
    RootBaseLine?:IBaselineEvent;
  }
  
  
  export interface ITechnicalEvent {
    DocumentId?: string;
    SitePath?: string;
    DocumentLink?: string;
    Status?:string;
    ProjcectReference?:string;
    ProjcectRevision?:string;
    FileName?:string;

  }
  
  
  export interface IBaselineEvent {
    ItemId?: number | null;
    SitePath?: string;
    Title?: string;
    Status?: string;
    Code?:string;
    BaselineVersion?:string;
    relatedBaseline?: IRelatedBaselines
  }
  
  
  export interface IRelatedBaselines {
    baseline: IBaselineEvent[];
    technical: ITechnicalEvent[];
  }
  
  
  export interface IBaselineExport{
    ItemId: number;
    SitePath?: string;
  }
      
  /* Baseline ends */