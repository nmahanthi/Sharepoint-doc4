export const DmsDocumentFormConstants = {
  defaultSectionsToFieldsMap: {
    "Main": [
      "Title",
      "ProjectReference",
      "ProjectRevision",
      "PLOwnership",
      "DocumentType",
      "DocumentLanguage",
      "ABS",
      "PBS",
      "SubmissionCategory",
      "PlannedSubmissionDate",
      "IssuingEntity",
      "ConfidentialityLevel",
      "GRMilestone",
      "Authors",
      "_Comments",
      "AlstomReference"
    ],
    "Additional Properties": [
      "WBS",
      "DBS",
      "FBS",
      "SBS",
      "SupplierReference",
      "SupplierRevision",
      "SupplierName",
      "OtherReference",
      "OtherRevision",
      "RequirementID",
      "ChangeRequestID",
      "Keywords",
      "FolioNumber",
      "AlstomPartNoReference",
      "AlstomPartNoRevision",
      "ExternalURL",

    ],
    "Project": [
      "Projectattr1",
      "Projectattr2",
      "Projectattr3",
      "Projectattr4",
      "Projectattr5",
      "Projectattr6",
      "Projectattr7",
      "Projectattr8",
      "Projectattr9",
      "Projectattr10"
    ],
    "Additional SER": [
      "LineNumber",
      "MoCStatus",
      "AssetID",
      "MaterialID",
      "LocationID",
      "JobPlanID",
      "InspectionID",
      "FailureID",
      "MaintenancePlanID",
      "HazardID",
      "PermitToWorkID",
      "MoCID",
      "IncidentID",
      "HandoverDocumentClassification",
      "HandoverDocumentTitle",
      "HandoverDocumentCode",
      "HandoverDocumentFilename",
      "HandoverPBSLevelCode",
      "HandoverDocumentStatus",
      "HandoverDocumentRevision"
    ],
    "Follow up": [
      "CustomerResponseExpectedDate",
      "CustomerResponseReference",
      "CustomerResponseDate",
      "CustomerResponseStatus",
      "ReforecastSubmissionDate",
      "ActualSubmissionDate",
      "CustomerMilestone",
      "TransmittalReference",
      "TransmittalDate",
      "ApprovalCycleLaunchDate",
      "Verifiedby",
      "VerificationDate",
      "Validatedby",
      "ValidationDate",
      "Approvedby",
      "ApprovalDate",
      "CustomerComments",
    ],
    "Applicability": [
      "EffectiveDate",
      "ExpiryDate",
      "ReplacementFor",
      "ReplacedBy",
      "ChangeFromNumber",
      "ChangeToNumber",
      "SafetyCriticity",
    ],
    "Technical Properties": [
      "Author",
      "Created",
      "Editor",
      "Modified"
    ]
  },
  hideInEditForm: [
    "_dlc_DocId",
    "_dlc_DocIdUrl",
    "_ExtendedDescription",
    "DocumentStatus",
    "WorkflowStatus",
    "DocumentSetDescription"
  ],
  hideInNewForm: [
    "_ExtendedDescription",
    "WorkflowStatus",
    "DocumentSetDescription",
    "DocumentStatus", "ActualSubmissionDate", "TransmittalReference", "TransmittalDate", "ApprovalCycleLaunchDate", "Verifiedby",
    "VerificationDate", "Validatedby", "ValidationDate", "Approvedby", "ApprovalDate", "AlstomReference", "DocumentStatus",
    "WorkflowStatus", "ScopeCountry", "CustomernameDirect", "CustomernameFinal", "ProductName", "ProjectAlias1",
    "_dlc_DocId", "_dlc_DocIdUrl", "_ExtendedDescription", "Author", "Created", "Editor", "Modified"
  ],
  hideInViewForm: [],
  disableInEditForm: [
    "ProjectReference", "ProjectRevision", "PLOwnership", "CustomerResponseExpectedDate", "CustomerResponseDate",
    "ActualSubmissionDate", "TransmittalReference", "TransmittalDate", "ApprovalCycleLaunchDate", "Verifiedby",
    "VerificationDate", "Validatedby", "ValidationDate", "Approvedby", "ApprovalDate", "AlstomReference", "DocumentStatus",
    "WorkflowStatus", "ScopeCountry", "CustomernameDirect", "CustomernameFinal", "ProductName", "ProjectAlias1",
    "_dlc_DocId", "_dlc_DocIdUrl", "_ExtendedDescription", "Author", "Created", "Editor", "Modified"
  ],
  disableInNewForm: [],

  enableInApplicableViewForm: ["Contractual", "CustomerRevision", "CustomerMilestone", "PlannedSubmissionDate",
    "ReforecastSubmissionDate", "ActualSubmissionDate",
    "TransmittalReference", "CustomerResponseExpectedDate", "CustomerResponseReference",
    "CustomerResponseDate", "CustomerResponseStatus",
    "CustomerComments"
  ],
  enableInApplicableViewFormDateFields: ["PlannedSubmissionDate",
    "ReforecastSubmissionDate", "ActualSubmissionDate",
    "CustomerResponseExpectedDate",
    "CustomerResponseDate"
  ],
  enableInApplicableViewFormBooleanFields: ['Contractual'],

  dmsUserApiAppId: "api://5558567f-a32f-417c-a15c-3e57b67f887e",
  confidentialityEndpoint: "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/confidentiality?code=7UdNGawE1jtHAJ_nH506T94FPHtCk9jWjpZIaEsBVnl5AzFuuMJOBw%3D%3D",
  updateApplicableDocumentEndpoint: "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/applicableDocument?code=g9KiiJXZFE9gcdGui_5Y6reGDlLxqiJ-nkvy0yjFFwMJAzFujkcXmA%3D%3D"

};

export const APPLICABLESTATUS = "Applicable";
export const DocumentStatusFieldKey = "DocumentStatus";

export const projectEndpoint = "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/project?code=ZHrOVblyNJJpSktnXLggB0zMpvsyQlFyexoL3ex4pg2yAzFug7zkjQ%3D%3D";
export const projectTemplateEndPoint="https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/documentContentType?code=HERi6sKQLmX-T8us1KYDLyfJ6tN1pfHKkWjfgidPO3UCAzFu7Q2nGw=="; 

export const CustomPermissionsListTitle = "PermissionsConfigurations";
export const CustomPermissionsLevels = {
  DocumentController: "Document Controller",
  DocumentManager: "Document Manager",
  Contributor: "Contributor",
  ProjectAdmin: 'Project Administrator'
};

export const TechnicalDocumentContentType = "Technical Document";
export const BaselineContentType = "Baseline";
export const CommonFields = {
  ContentTypeFieldName: "ContentType",
  ServerRelativeUrl: "ServerRelativeUrl",
  FileRef: "FileRef",
  FileLeafRef: "FileLeafRef"
};
export const PermissionListValues = {
  DocumentController: "Document Controller",
  Contributor: "Contributor"
};
export const TechnicalDocumentsFields = {
  Title: "Title",
  ProjectReference: "ProjectReference"
};
export const FieldNames = {
  ProjectReference: "ProjectReference",
  Revision: "ProjectRevision",
  PLOwnership: "PLOwnership",
  Product: "AlstomProduct",
  ProjectAlias: "ProjectAlias1",
  DirectCustomer: "DirectCustomer",
  EndCustomer: "EndCustomer",
  ConfidentialityLevel: "ConfidentialityLevel",
  AlstomPartNumbers:"AlstomPartNumbers",
  DocumentStatus: "DocumentStatus"
};
export const ListUrls = {
  Draft: "Draft",
  ApplicableDocuments: "ApplicableDocuments",
  PreviousVersions: "PreviousVersions",
  EditForm: "/_layouts/15/listform.aspx?PageType=6&ListId=[ListId]&ID=[ItemId]&Source=[Src]",
  ViewForm: "/_layouts/15/listform.aspx?PageType=4&ListId=[ListId]&ID=[ItemId]&Source=[Src]",
  NewForm: "/_layouts/15/listform.aspx?PageType=8&ListId=[ListId]&ID=[ItemId]&Source=[Src]"
};
export const SensitivityValueConfidential = "Confidential";
export const ContentTypeSiteIds = {
  TechnicalDocuments: "0x0120D5200010BA2408052EEF418E580CFAEB67F318"
};
export const CustomPermissionListName = "PermissionsConfigurations";

export const BatchImportConstants = {
  MassImportStatusFieldCustomizerProperties: "MassImportStatusFieldCustomizerProperties",
  BatchImportListTitle: "BatchImport",
  BatchImportListFields: {
    Status: "Status",
    Error: "Error",
    JobId: "JobId",
    ErrorDetails: "ErrorDetails",
  },
  BatchImportStatus: {
    NotStarted: "Not Sarted",
    Requested: "Rquested",
    Queued: "Queued",
    Processing: "Processing",
    Completed: "Completed",
    Failed: "Failed"
  },
  DmlFilePrefix: "DML_",
  MidFilePrefix: "MID_",
  MiaFilePrefix: "MIA_",
  MassImportEndpointUrl: "https://lab-dms-int-fnapp-user-api.azurewebsites.net/api/massimport?code=ZHrOVblyNJJpSktnXLggB0zMpvsyQlFyexoL3ex4pg2yAzFug7zkjQ%3D%3D",
};

export const SearchConstants = {
  DocumentIdManagedPropName:"DlcDocId",
  ItemIdManagedPropName: "ListItemID",
  SitePathManagedPropName: "Path",
  ContentTypeManagedPropName: "ContentType"
};
export const ProjectPropertiesList = {
  RelativeUrl : "/Lists/ProjectProperties",
  NameField: "Title",
  ValueField: "Value"
};