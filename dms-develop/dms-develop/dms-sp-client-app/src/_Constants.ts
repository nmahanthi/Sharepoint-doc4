const _CONSTANTS = {

    AbortedStatus: 'Aborted',
    CancelledStatus:'Cancelled',
    CancellingStatus:'Cancelling',
    TaskListName: 'TaskStatus',
    WFStatusListName: 'WorkflowStatus',
    ProofReadingStatusName: "ProofReading Reviewer",
    TechnicalDocumentContentType: "Technical Documents",
    CommonFields: {
        ContentTypeFieldName:"ContentType",
        ServerRelativeUrl: "ServerRelativeUrl",
        FileRef:"FileRef"
    },
    PermissionListValues:{
        DocumentController: "Document Controller",
        Contributor: "Contributor"
    },
    TechnicalDocumentsFields:{
        Title:"Title",
        ProjectReference:"ProjectReference"

    },

    DraftListFieldNames: {
        DocumentId: "OData__dlc_DocId",
        Created: 'Created',
        Modified: 'Modified',
        ID: 'ID',
        WorkflowStatus: 'WorkflowStatus',
        DocumentStatus: 'DocumentStatus'
    },

    TaskListFieldNames: {
        TaskAssigneeType: 'TaskAssigneeType',
        ClosedDate: 'ClosedDate0',
        ID: 'ID',
        Result: 'Result0',
        TaskRunningStatus: 'TaskRunningStatus',
        AssignedDate: 'AssignedDate0',
        TaskComments: 'TaskComments',
        NextTaskId: 'NextTaskId0',
        Created: 'Created',
        Modified: 'Modified',
        RevisionId: 'RevisionId0',
        PrevTaskId: 'PrevTaskId',
        IsDelegatedTask: 'IsDelegatedTask',
        InstanceId: 'InstanceId1',
    },

    WFStatusListFieldNames: {
        WorkflowType: 'WorkflowType',
        Id: 'ID',
        Created: 'Created',
        Modified: 'Modified',
        RevisionId: 'RevisionId',
        InstanceURL: 'InstanceURL',
        InstanceId: 'InstanceId0',
        WorkflowStatus:"WorkflowStatus1",        
        StardDate:'StartDate1',
        ClosedDate:'CloseDate',
        CancellationReason:'CancellationReason',
        Requestor:'Requestor'
    },

    /* Baseline starts */
    BaselineListFieldNames: {
        Title:"Title",
        BaselineStatus:"BaselineStatus",
        BaselineVersion:"BaselineVersion",	
        ChangeDate:"ChangeDate",	
        ChangeUser:"ChangeUser1Id",	
        CreationDate:"CreationDate",	
        CreationUser:"CreationUserId",
        CRID:"CRID",
        Code:"Code",
        Label:"Label",	
        BaselineComments:"BaselineComments",
    },

    FieldTypes: {
        TaxonomyFieldTypeMulti:'TaxonomyFieldTypeMulti',
        Text:'Text',
        Choice:'Choice',
        MultiChoice:'MultiChoice',
        DateTime:'DateTime',
        UserMulti:'UserMulti',
        User:'User',
        Note:'Note',
        URL:'URL',
        Boolean:'Boolean',
        TaxonomyFieldType:'TaxonomyFieldType'
    },

    ContentTypeNames: {
        Technical: 'Technical Document',
        Baseline: 'Baseline'
    },
    ListNames:
    {
        TechnicalReferenceList:"Baseline documents",
        BaselineRefernceList:"Baseline Baseline"
    },
    TechnicalRefernceListFieldNames:{        
        DocumentID:"DocumentID",
        SiteURL:"SiteURL",
        BaselineParentID:"BaselineParentID",
        Title:"Title",
        Status:"DocumentStatus"
    },
    BaselineRefernceListFieldNames:{        
        BaselineChildID:"BaselineChildID",
        SiteURL:"SiteURL",
        BaselineParentID:"BaselineParentID",
        Title:"Title",
        Label:"Label",
        Status:"BaselineStatus",
    },
    TestCICD:"RemoveLater"

    /* Baseline ends */
}
export default _CONSTANTS;