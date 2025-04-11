let
    // _SiteUrl = "https://alstomgrouppp.sharepoint.com/sites/DOC4A-INT",
    // 1. Define the base URL
    baseURL = _SiteUrl & "/_api/v2.1/termstore/groups",

    // 2. Fetch all term groups (Groups)
    GetGroups = Json.Document(Web.Contents(baseURL)),
    GroupsTable = Table.FromRecords(GetGroups[value]),

    // 3. Define a function to fetch all term sets (Sets) within each term group
    GetSetsByGroup = (groupID as text) =>
        let
            setsURL = _SiteUrl & "/_api/v2.1/termstore/groups('" & groupID & "')/sets",
            setsResponse = Json.Document(Web.Contents(setsURL)),
            setsTable = Table.FromRecords(setsResponse[value])
        in
            setsTable,

    // 4. Add a column for term sets to each term group
    GroupsWithSets = Table.AddColumn(GroupsTable, "Sets", each GetSetsByGroup([id])),

    // 5. Define a function to fetch all terms within each term set
    GetTermsBySet = (groupID as text, setID as text) =>
        let
            termsURL = _SiteUrl & "/_api/v2.1/termstore/groups('" & groupID & "')/sets('" & setID & "')/terms",
            termsResponse = Json.Document(Web.Contents(termsURL)),
            termsTable = Table.FromRecords(termsResponse[value])
        in
            termsTable,

    // 6. Add a column for terms to each term set
    AddTermsToSets = Table.AddColumn(GroupsWithSets, "SetsWithTerms", 
        each Table.AddColumn([Sets], "Terms", (currentSet) => GetTermsBySet([id], currentSet[id]))),

    // 7. Expand the final results and add 'localizedNames.name'
    // First, expand 'SetsWithTerms' to get 'id', 'name', 'Terms', and 'localizedNames'
    ExpandedSets = Table.ExpandTableColumn(AddTermsToSets, "SetsWithTerms", {"id", "name", "Terms", "localizedNames"}, {"SetID", "SetName", "Terms", "LocalizedNames"}),

    // 8. Expand 'localizedNames' to get the 'name' field
    ExpandedLocalizedNames = Table.ExpandListColumn(ExpandedSets, "LocalizedNames"),
    ExpandedLocalizedNamesWithName = Table.ExpandRecordColumn(ExpandedLocalizedNames, "LocalizedNames", {"name"}, {"LocalizedSetName"}),

    // 9. Expand the 'Terms' column
    ExpandedTerms = Table.ExpandTableColumn(ExpandedLocalizedNamesWithName, "Terms", {"id", "labels"}, {"TermID", "Labels"}),
    #"Expanded Labels" = Table.ExpandListColumn(ExpandedTerms, "Labels"),
    #"Expanded Labels1" = Table.ExpandRecordColumn(#"Expanded Labels", "Labels", {"name", "isDefault", "languageTag"}, {"Labels.name", "Labels.isDefault", "Labels.languageTag"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded Labels1", each [name] = "DOC4A"),
    #"Renamed Columns" = Table.RenameColumns(#"Filtered Rows",{{"name", "Term Store"}, {"Labels.name", "LabelName"}}),
    #"Filtered Rows1" = Table.SelectRows(#"Renamed Columns", each ([LocalizedSetName] = "ABS" or [LocalizedSetName] = "PBS" or [LocalizedSetName] = "PLOwnership"))

in
    #"Filtered Rows1"