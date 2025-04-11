[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $SiteUrl = "https://alstomgrouppp.sharepoint.com/sites/DMS-tempsrc"
)
BeforeDiscovery {
    Connect-PnPOnline -Url $SiteUrl -DeviceLogin
    $DiscoveryExpectedDocFields = @(Import-Csv -Path "$PSScriptRoot\data\document-fields.csv" -Encoding UTF8)
    $DiscoveryExpectedSiteFields = @(Import-Csv -Path "$PSScriptRoot\data\site-fields.csv" -Encoding UTF8)
    $DiscoveryExpectedAttachmentFields = @(Import-Csv -Path "$PSScriptRoot\data\attachment-fields.csv" -Encoding UTF8)
    $DmsGroupFields = Get-PnPField -Group "DMS"
    $Libraries = @("Draft","Applicable Documents","Previous Versions")
}
BeforeAll {
    $ExpectedDocFields = @(Import-Csv -Path "$PSScriptRoot\data\document-fields.csv" -Encoding UTF8)
    $ExpectedSiteFields = @(Import-Csv -Path "$PSScriptRoot\data\site-fields.csv" -Encoding UTF8)
    $ExpectedAttachmentFields = @(Import-Csv -Path "$PSScriptRoot\data\attachment-fields.csv" -Encoding UTF8)
}
Describe 'Site-level Fields'{
    Describe 'Document Fields'{
        Describe 'Field <_.Field> configuration' -ForEach $DiscoveryExpectedDocFields {
            BeforeAll {
                $spField = Get-PnPField -Identity $_.Field
            }
            It "Should have the correct title"{
                $spField.Title | Should -Be $_.Title
            }
            It "Should have the correct type"{
                $spField.TypeAsString | Should -Be $_.FieldType
            }
            It "Should have the correct required setting"{
                $spField.Required | Should -Be ($_.Required -eq "Yes")
            }
            It "Should have the correct hidden setting"{
                $spField.Hidden | Should -Be ($_.SiteColHidden -eq "Yes")
            }
            It "Should be in the DMS group"{
                if(-not (@("Title","_Comments","Keywords") -contains $_.Field)) {
                    $spField.Group | Should -Be "DMS"
                }
            }
        }
    }
    Describe 'Project Fields'{
        Describe 'Field <_.Field> configuration' -ForEach $DiscoveryExpectedSiteFields {
            BeforeAll {
                $spField = Get-PnPField -Identity $_.Field
            }
            It "Should have the correct title"{
                $spField.Title | Should -Be $_.Title
            }
            It "Should have the correct type"{
                $spField.TypeAsString | Should -Be $_.FieldType
            }
            It "Should be in the DMS group"{
                if(-not (@("Title","_Comments","Keywords") -contains $_.Field)) {
                    $spField.Group | Should -Be "DMS"
                }
            }
        }
    }
    Describe 'Attachment Fields'{
        Describe 'Field <_.Field> configuration' -ForEach $ExpectedAttachmentFields {
            BeforeAll {
                $spField = Get-PnPField -Identity $_.Field
            }
            It "Should have the correct title"{
                $spField.Title | Should -Be $_.Title
            }
            It "Should have the correct type"{
                $spField.TypeAsString | Should -Be $_.FieldType
            }
            It "Should be in the DMS group"{
                if(-not (@("Title","_Comments","Keywords") -contains $_.Field)) {
                    $spField.Group | Should -Be "DMS"
                }
            }
        }
    }
    Describe 'Other fields' {
        It "Field <_.InternalName> should be in the expected fields" -ForEach $DmsGroupFields {
            @($ExpectedDocFields; $ExpectedSiteFields; $ExpectedAttachmentFields) | Foreach-Object { $_.Field } | Should -Contain $_.InternalName
        }
    }
}
Describe 'Library "<_>" Fields' -ForEach $Libraries {
    BeforeAll {
        $library = $_
    }
    Describe 'Document Fields'{
        Describe 'Field <_.Field> configuration' -ForEach $DiscoveryExpectedDocFields {
            BeforeAll {
                $spField = Get-PnPField -Identity $_.Field -List $library
            }
            It "Should have the correct title"{
                $spField.Title | Should -Be $_.Title
            }
            It "Should have the correct type"{
                $spField.TypeAsString | Should -Be $_.FieldType
            }
            It "Should have the correct hidden setting"{
                $spField.Hidden | Should -Be ($_.SiteColHidden -eq "Yes")
            }
            It "Should be in the DMS group"{
                if(-not (@("Title","_Comments","Keywords") -contains $_.Field)) {
                    $spField.Group | Should -Be "DMS"
                }
            }
        }
    }
    Describe 'Attachment Fields'{
        Describe 'Field <_.Field> configuration' -ForEach $ExpectedAttachmentFields {
            BeforeAll {
                $spField = Get-PnPField -Identity $_.Field -List $library
            }
            It "Should have the correct title"{
                $spField.Title | Should -Be $_.Title
            }
            It "Should have the correct type"{
                $spField.TypeAsString | Should -Be $_.FieldType
            }
            It "Should be in the DMS group"{
                if(-not (@("Title","_Comments","Keywords") -contains $_.Field)) {
                    $spField.Group | Should -Be "DMS"
                }
            }
        }
    }
}