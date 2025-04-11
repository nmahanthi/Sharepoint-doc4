let
     // Connect to the SharePoint site and get the items
    // _SiteUrl = "https://alstomgrouppp.sharepoint.com/sites/DOC4A-INT/",
    // _ApplicableDocLibraryID = "30af08c6-1860-4565-af8b-7b0a0de1dbd7",
    // _ApplicableDocSetContentTypeId = "0x0120D5200010BA2408052EEF418E580CFAEB67F3180069E62A9999038B4EBAAEF45F5F658567",
    
    _TaxonomyTableName = #"Taxonomy Data", // Reference to Table 2 query
    UserFields = {"Change User", "Author","Verifier","Validator","Approver"},  // Add any other user-related fields here
    DateFields = {"CustomerResponseExpectedDate","PlannedSubmissionDate","ReforecastSubmissionDate","CustomerResponseDate","TransmittalDate","Created","Modified","ApprovalCycleLaunchDate","VerificationDate","ValidationDate","ApprovalDate","ActualSubmissionDate"},
    SetBoolean = {"Contractual","Latest Revision"}, // Add any other user-related fields here

    Source = SharePoint.Tables(_SiteUrl, [Implementation=null, ApiVersion=15]),
    DocumentLibrary = Source{[Id=_ApplicableDocLibraryID]}[Items],

    // Filter to include only document sets using the specific ContentTypeId
    FilterContentType = Table.SelectRows(DocumentLibrary, each [ContentTypeId] = _ApplicableDocSetContentTypeId),
    FilteredItems = Table.SelectRows(FilterContentType, each [ServerRedirectedEmbedUri] = null),
    LatestRevision = Table.AddColumn(FilteredItems, "Latest Revision", each "Yes"),

    // Expand relevant internal columns
    ExpandedEditors = Table.ExpandRecordColumn(LatestRevision, "Editor", {"Title"}, {"Change User"}),
    ExpandedAuthors = Table.ExpandRecordColumn(ExpandedEditors, "Author", {"Title"}, {"Author"}),
    ExpandedVerifier = Table.ExpandRecordColumn(ExpandedAuthors, "Verifiedby", {"Title"}, {"Verifier"}),
    ExpandedValidator = Table.ExpandRecordColumn(ExpandedVerifier, "Validatedby", {"Title"}, {"Validator"}),
    ExpandedApprover = Table.ExpandRecordColumn(ExpandedValidator, "Approvedby", {"Title"}, {"Approver"}),

    // Expand Folder to get ServerRelativeUrl
    ExpandedOData__dlc_DocIdUrl = Table.ExpandRecordColumn(ExpandedApprover, "OData__dlc_DocIdUrl", {"Url"}, {"OData__dlc_DocIdUrl"}),

    // Expand Folder to get ServerRelativeUrl
    ExpandedFolder = Table.ExpandRecordColumn(ExpandedOData__dlc_DocIdUrl, "Folder", {"ServerRelativeUrl"}, {"ServerRelativeUrl"}),

    // // Count child items based on ServerRelativeUrl
    // RelatedFilesCount = Table.AddColumn(ExpandedFolder, "Related files Count",
    //     each
    //         let
    //             docSetUrl = [ServerRelativeUrl],
    //             childItems = Table.SelectRows(DocumentLibrary, each Text.StartsWith([Folder][ServerRelativeUrl], docSetUrl))
    //         in
    //             Table.RowCount(childItems),
    //     Int64.Type
    // ),

    
    AddHasRelatedFilesColumn = Table.AddColumn(ExpandedFolder, "Has Related Files", each ""),
    
    // // Add the "Has Related Files" column
    // AddHasRelatedFilesColumn = Table.AddColumn(RelatedFilesCount, "Has Related Files",
    //     each if [Related files Count] > 0 then "YES" else "NO",
    //     type text
    // ),

    AddHasRelatedFilesColumn2 = Table.AddColumn(AddHasRelatedFilesColumn, "Document Translated", each ""),
    AddHasRelatedFilesColumn3 = Table.AddColumn(AddHasRelatedFilesColumn2, "OTD Delay Days", each ""),
    AddHasRelatedFilesColumn4 = Table.AddColumn(AddHasRelatedFilesColumn3, "Time Taken by Verifier", each ""),
    AddHasRelatedFilesColumn5 = Table.AddColumn(AddHasRelatedFilesColumn4, "Time Taken by Validator", each ""),
    AddHasRelatedFilesColumn6 = Table.AddColumn(AddHasRelatedFilesColumn5, "Time Taken by Approver", each ""),

    RenamedColumns = Table.RenameColumns(AddHasRelatedFilesColumn6, {{"ID", "ID.1"}}),

    ReplaceNullsRecords = List.Accumulate(
        UserFields,
        RenamedColumns,
        (state, current) => Table.ReplaceValue(state, null, "", Replacer.ReplaceValue, {current})
    ),


    ReplaceBooleanTrue = List.Accumulate(
        SetBoolean, // Ensure you're using SetBoolean here
        ReplaceNullsRecords,
        (state, current) => Table.ReplaceValue(state, true, "Yes", Replacer.ReplaceValue, {current}) // If the values are Booleans
    ),
    
    ReplaceBooleanFalse = List.Accumulate(
        SetBoolean,
        ReplaceBooleanTrue, // Use the state from the previous replacement
        (state, current) => Table.ReplaceValue(state, false, "No", Replacer.ReplaceValue, {current}) // Replacing false with "No"
    ),

    ReplaceBooleanNUll = List.Accumulate(
            SetBoolean,
            ReplaceBooleanFalse, // Use the state from the previous replacement
            (state, current) => Table.ReplaceValue(state, null, "No", Replacer.ReplaceValue, {current}) // Replacing false with "No"
        ),


    // Expand ABS column to get the TermGuid
    ExpandedABS = Table.ExpandRecordColumn(ReplaceBooleanNUll, "ABS", {"TermGuid"}, {"ABS TermGuid"}),
    // // Merge with Table 2 (assuming you keep Table 2 as a separate query)
     // Merge with Table 2
    MergedTableABS = Table.NestedJoin(ExpandedABS, {"ABS TermGuid"}, _TaxonomyTableName, {"TermID"}, "Table2Details", JoinKind.LeftOuter),
    // Expand the merged details to retrieve the LabelName
    ExpandedTable2DetailsABS = Table.ExpandTableColumn(MergedTableABS, "Table2Details", {"LabelName"}, {"ABS LabelName"}),
    // Replace ABS with LabelName if available
    FinalTableABS = Table.AddColumn(ExpandedTable2DetailsABS, "ABS", each if [ABS LabelName] <> null then [ABS LabelName] else [ABS TermGuid], type text),

    

 
    //Expand PBS column to get the TermGuid
    ExpandedPBS = Table.ExpandRecordColumn(FinalTableABS, "PBS", {"TermGuid"}, {"PBS TermGuid"}),
    // // Merge with Table 2 (assuming you keep Table 2 as a separate query)
     // Merge with Table 2
    MergedTablePBS = Table.NestedJoin(ExpandedPBS, {"PBS TermGuid"}, _TaxonomyTableName, {"TermID"}, "Table2Details", JoinKind.LeftOuter),
    // Expand the merged details to retrieve the LabelName
    ExpandedTable2DetailsPBS = Table.ExpandTableColumn(MergedTablePBS, "Table2Details", {"LabelName"}, {"PBS LabelName"}),
    // Replace PBS with LabelName if available
    FinalTablePBS = Table.AddColumn(ExpandedTable2DetailsPBS, "PBS", each if [PBS LabelName] <> null then [PBS LabelName] else [PBS TermGuid], type text),


  //Expand PLOwnership column to get the TermGuid
    ExpandedPLOwnership = Table.ExpandRecordColumn(FinalTablePBS, "PLOwnership", {"TermGuid"}, {"PLOwnership TermGuid"}),
    // // Merge with Table 2 (assuming you keep Table 2 as a separate query)
     // Merge with Table 2
    MergedTablePLOwnership = Table.NestedJoin(ExpandedPLOwnership, {"PLOwnership TermGuid"}, _TaxonomyTableName, {"TermID"}, "Table2Details", JoinKind.LeftOuter),
    // Expand the merged details to retrieve the LabelName
    ExpandedTable2DetailsPLOwnership = Table.ExpandTableColumn(MergedTablePLOwnership, "Table2Details", {"LabelName"}, {"PLOwnership LabelName"}),
    // Replace PLOwnership with LabelName if available
    FinalTablePLOwnership = Table.AddColumn(ExpandedTable2DetailsPLOwnership, "PLOwnership", each if [PLOwnership LabelName] <> null then [PLOwnership LabelName] else [PLOwnership TermGuid], type text),
    ExtractedDates = Table.TransformColumns(FinalTablePLOwnership,
    {{"CustomerResponseExpectedDate", DateTime.Date, type date}, 
    {"PlannedSubmissionDate", DateTime.Date, type date}, 
    {"ReforecastSubmissionDate", DateTime.Date, type date}, 
    {"CustomerResponseDate", DateTime.Date, type date}, 
    {"TransmittalDate", DateTime.Date, type date}, 
    {"Created", DateTime.Date, type date}, 
    {"Modified", DateTime.Date, type date}, 
    {"VerificationDate", DateTime.Date, type date},
    {"ValidationDate", DateTime.Date, type date},
    {"ApprovalDate", DateTime.Date, type date},
    {"ActualSubmissionDate", DateTime.Date, type date},
    {"ApprovalCycleLaunchDate", DateTime.Date, type date}}),

   // Expand ABS column to get the TermGuid
    // ExpandedABS = Table.ExpandRecordColumn(RenamedColumns, "ABS", {"TermGuid"}, {"ABS"}),
    // ExpandedPBS = Table.ExpandRecordColumn(ExpandedABS, "PBS", {"TermGuid"}, {"PBS"}),
    // ExpandedPLOwnership = Table.ExpandRecordColumn(ExpandedPBS, "PLOwnership", {"TermGuid"}, {"PLOwnership"}),
     // Convert the extracted Date fields to the "MM-DD-YYYY" format
    FinalFormattedDates = Table.TransformColumns(ExtractedDates,
        {
            {"CustomerResponseExpectedDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"PlannedSubmissionDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"ReforecastSubmissionDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"CustomerResponseDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"TransmittalDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"Created", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"Modified", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"VerificationDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"ValidationDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"ApprovalDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"ActualSubmissionDate", each Date.ToText(_, "MM-dd-yyyy"), type text},
            {"ApprovalCycleLaunchDate", each Date.ToText(_, "MM-dd-yyyy"), type text}
        }
    )
in
    FinalFormattedDates