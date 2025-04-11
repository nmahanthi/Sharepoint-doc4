[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $SiteUrl = "https://alstomgrouppp.sharepoint.com/sites/DMS-tempsrc"
)
BeforeDiscovery {
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin
    $DiscoveryExpectedDocFields = @(Import-Csv -Path "$PSScriptRoot\data\document-fields.csv" -Encoding UTF8)
    $DiscoveryDocFieldLinks = Get-PnPContentType -Identity "Technical Document" -Includes FieldLinks | Select-Object -ExpandProperty FieldLinks
    $AttachmenFieldsMap = @{}
    @(Import-Csv -Path "$PSScriptRoot\data\attachment-fields.csv" -Encoding UTF8) | ForEach-Object {
        if($null -eq $AttachmenFieldsMap[$_.ContentType]){
            $AttachmenFieldsMap[$_.ContentType] = @{
                Fields = @();
                FieldLinks = Get-PnPContentType -Identity $_.ContentType -Includes FieldLinks | Select-Object -ExpandProperty FieldLinks;
                ContentType = $_.ContentType
            }
        }
        $AttachmenFieldsMap[$_.ContentType].Fields += $_
    }
    $DiscoveryExpectedAttachmentFields = $AttachmenFieldsMap.Values
    $Libraries = @("Draft","Applicable Documents","Previous Versions")
}
BeforeAll {
    $ExpectedDocFields = @(Import-Csv -Path "$PSScriptRoot\data\document-fields.csv" -Encoding UTF8)
    $AttachmenFieldsMap = @{}
    @(Import-Csv -Path "$PSScriptRoot\data\attachment-fields.csv" -Encoding UTF8) | ForEach-Object {
        if($null -eq $AttachmenFieldsMap[$_.ContentType]){
            $AttachmenFieldsMap[$_.ContentType] = @{
                Fields = @();
                FieldLinks = Get-PnPContentType -Identity $_.ContentType -Includes FieldLinks | Select-Object -ExpandProperty FieldLinks;
                ContentType = $_.ContentType
            }
        }
        $AttachmenFieldsMap[$_.ContentType].Fields += $_
    }
    $ExpectedAttachmentFields = $AttachmenFieldsMap.Values
    $documentSetContentAllowedTypes = @("Native Document","Translation Document","Comments Document","Rendition Document","Appendix")
}
Describe 'Site Content Types'{
    BeforeAll {
        $DocFieldLinks = Get-PnPContentType -Identity "Technical Document" -Includes FieldLinks | Select-Object -ExpandProperty FieldLinks
    }
    Describe 'Technical Document'{
        Describe 'Field Links'{
            Describe 'Field <_.Field> link' -ForEach $DiscoveryExpectedDocFields {
                BeforeAll {
                    $expectedField = $_
                    $fieldLink = $DocFieldLinks | Where-Object { $_.Name -eq $expectedField.Field} | Select-Object -First 1
                }
                It "Should exist"{
                    $fieldLink | Should -Not -BeNullOrEmpty
                }
                It "Should have the correct required setting"{
                    $fieldLink.Required | Should -Be ($_.Required -eq "Yes")
                }
                It "Should have the correct hidden setting"{
                    $fieldLink.Hidden | Should -Be ($_.FieldLinkHidden -eq "Yes")
                }
                It "Should be in the correct order"{
                    if($expectedField.Order -ne "") {
                        @($DocFieldLinks | Select-Object -ExpandProperty Name).IndexOf($expectedField.Field) + 1 | Should -Be $_.Order
                    }
                }
            }
        }
        Describe 'Other fields' {
            It "Field <_.Name> should be in the expected fields" -ForEach $DiscoveryDocFieldLinks {
                if(-not (@("ContentType",
                "ItemChildCount","FolderChildCount","FileLeafRef","_dlc_DocId","_dlc_DocIdUrl",
                "_dlc_DocIdPersistId","b7428419b41846fb99ec3ba983dae58c","TaxCatchAll","TaxCatchAllLabel",
                "n5a1eb9573fb4634a4f840f34bb0574e","l522119e42874dc78b23656c0391ab96","Description") -contains $_.Name)) {
                    $ExpectedDocFields | Foreach-Object { $_.Field } | Should -Contain $_.Name
                }
            }
        }
        Describe 'Document Set settings'{
            It "Should have the correct Document Set settings"{
                $template = Get-PnPDocumentSetTemplate -Identity "Technical Document"
                $documentSetContentAllowedTypes | ForEach-Object { 
                    Get-PnPContentType -Identity $_ | Select-object -ExpandProperty Id | Select-Object -ExpandProperty StringValue
                } | ForEach-Object {
                    $ctId = $_
                    $template.AllowedContentTypes | Foreach-Object { $_.StringValue } | Where-Object { $_.StartsWith($ctId)} | Should -HaveCount 1
                }
                $template.AllowedContentTypes | Should -HaveCount $documentSetContentAllowedTypes.Count
            }
        }
    }
    Describe '<ContentType>' -ForEach $DiscoveryExpectedAttachmentFields {
        BeforeAll {
            $fieldLinks = $_.FieldLinks
            $contentType = $_.ContentType
            $fields = $_.Fields
        }
        Describe 'Field Links'{
            Describe 'Field <_.Field> link' -ForEach $_.Fields {
                BeforeAll {
                    $expectedField = $_
                    $fieldLink = $fieldLinks | Where-Object { $_.Name -eq $expectedField.Field} | Select-Object -First 1
                }
                It "Should exist"{
                    $fieldLink | Should -Not -BeNullOrEmpty
                }
                It "Should have the correct required setting"{
                    $fieldLink.Required | Should -Be ($_.Required -eq "Yes")
                }
                It "Should have the correct hidden setting"{
                    $fieldLink.Hidden | Should -Be ($_.FieldLinkHidden -eq "Yes")
                }
                It "Should be in the correct order"{
                    if($expectedField.Order -ne "") {
                        @($fieldLinks | Select-Object -ExpandProperty Name).IndexOf($expectedField.Field) + 1 | Should -Be $_.Order
                    }
                }
            }
        }
        Describe 'Other fields' {
            It "Field <_.Name> should be in the expected fields" -ForEach $_.FieldLinks {
                if(-not (@("ContentType",
                "ItemChildCount","FolderChildCount","FileLeafRef","_dlc_DocId","_dlc_DocIdUrl",
                "_dlc_DocIdPersistId","b7428419b41846fb99ec3ba983dae58c","TaxCatchAll","TaxCatchAllLabel",
                "n5a1eb9573fb4634a4f840f34bb0574e","l522119e42874dc78b23656c0391ab96","Description","SelectFilename",
                "Created","Title","Modified","Modified_x0020_By","Created_x0020_By") -contains $_.Name)) {
                    $ExpectedAttachmentFields | Where-Object {$_.ContentType -eq $contentType} | Select-Object -ExpandProperty Fields | Select-Object -ExpandProperty Field | Should -Contain $_.Name
                }
            }
        }
    }
}
Describe 'Library "<_>" Content Types' -ForEach $Libraries {
    BeforeAll {
        $library = $_
        $DocFieldLinks = Get-PnPContentType -Identity "Technical Document" -Includes FieldLinks -List $library | Select-Object -ExpandProperty FieldLinks
    }
    Describe 'Technical Document'{
        Describe 'Field Links'{
            Describe 'Field <_.Field> link' -ForEach $DiscoveryExpectedDocFields {
                BeforeAll {
                    $expectedField = $_
                    $fieldLink = $DocFieldLinks | Where-Object { $_.Name -eq $expectedField.Field} | Select-Object -First 1
                }
                It "Should exist"{
                    $fieldLink | Should -Not -BeNullOrEmpty
                }
                It "Should have the correct required setting"{
                    $fieldLink.Required | Should -Be ($_.Required -eq "Yes")
                }
                It "Should have the correct hidden setting"{
                    $fieldLink.Hidden | Should -Be ($_.FieldLinkHidden -eq "Yes")
                }
                It "Should be in the correct order"{
                    if($expectedField.Order -ne "") {
                        @($DocFieldLinks | Select-Object -ExpandProperty Name).IndexOf($expectedField.Field) + 1 | Should -Be $_.Order
                    }
                }
            }
        }
    }
    Describe '<ContentType>' -ForEach $DiscoveryExpectedAttachmentFields {
        BeforeAll {
            $fieldLinks = Get-PnPContentType -Identity $_.ContentType -List $library -Includes FieldLinks | Select-Object -ExpandProperty FieldLinks;
            $contentType = $_.ContentType
            $fields = $_.Fields
        }
        Describe 'Field Links'{
            Describe 'Field <_.Field> link' -ForEach $_.Fields {
                BeforeAll {
                    $expectedField = $_
                    $fieldLink = $fieldLinks | Where-Object { $_.Name -eq $expectedField.Field} | Select-Object -First 1
                }
                It "Should exist"{
                    $fieldLink | Should -Not -BeNullOrEmpty
                }
                It "Should have the correct required setting"{
                    $fieldLink.Required | Should -Be ($_.Required -eq "Yes")
                }
                It "Should have the correct hidden setting"{
                    $fieldLink.Hidden | Should -Be ($_.FieldLinkHidden -eq "Yes")
                }
                It "Should be in the correct order"{
                    if($expectedField.Order -ne "") {
                        @($fieldLinks | Select-Object -ExpandProperty Name).IndexOf($expectedField.Field) + 1 | Should -Be $_.Order
                    }
                }
            }
        }
        Describe 'Other fields' {
            It "Field <_.Name> should be in the expected fields" -ForEach $_.FieldLinks {
                if(-not (@("ContentType",
                "ItemChildCount","FolderChildCount","FileLeafRef","_dlc_DocId","_dlc_DocIdUrl",
                "_dlc_DocIdPersistId","b7428419b41846fb99ec3ba983dae58c","TaxCatchAll","TaxCatchAllLabel",
                "n5a1eb9573fb4634a4f840f34bb0574e","l522119e42874dc78b23656c0391ab96","Description","SelectFilename",
                "Created","Title","Modified","Modified_x0020_By","Created_x0020_By") -contains $_.Name)) {
                    $ExpectedAttachmentFields | Where-Object {$_.ContentType -eq $contentType} | Select-Object -ExpandProperty Fields | Select-Object -ExpandProperty Field | Should -Contain $_.Name
                }
            }
        }
    }
}