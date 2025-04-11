export interface IBaseLineExportEvent {
  ExportDocuments?: IRelatedBaselines;
  UserEmail?: string;
  SiteUrl?: string;
  RootBaseLine?: IBaselineEvent;
}


export interface ITechnicalEvent {
  DocumentId?: string;
  SitePath?: string;
  DocumentLink?: string;
  Status?: string;
  ProjcectReference?: string;
  ProjcectRevision?: string;
  FileName?: string;

}


export interface IBaselineEvent {
  ItemId?: number | string;
  SitePath?: string;
  Title?: string;
  Status?: string;
  Code?: string;
  BaselineVersion?: string;
  relatedBaseline?: IRelatedBaselines
}


export interface IRelatedBaselines {
  baseline: IBaselineEvent[];
  technical: ITechnicalEvent[];
}
export interface IcsvFormat {
  ItemId?: number | string;
  SiteUrl?: string;
  Label?: string;
  Status?: string;
  Path?: string;
  ProjectReference?: string;
  ProjectRevision?: string;
  FileName?: string;
}