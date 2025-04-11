#Example:
#.\Invoke-ProvisionDmsSites.ps1 -CsvPath .\sample-sites.csv -EnableWebhook -Interactive 
[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]
    $CsvPath,
    [Parameter(ParameterSetName = 'Interactive')]
    [Parameter(ParameterSetName = 'AppRegistration')]
    [string]
    $ConfigPath = "$PSScriptRoot\config.json",
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]
    $ClientId,
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]
    $Tenant,
    [Parameter(Mandatory = $true, ParameterSetName = 'AppRegistration')]
    [string]
    $Thumbprint,
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [Parameter(Mandatory = $false, ParameterSetName = 'AppRegistration')]
    [switch]
    $EnableWebhook,
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
    [switch]
    $Interactive,
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [Parameter(Mandatory = $false, ParameterSetName = 'AppRegistration')]
    [string[]]
    $Handlers,
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [Parameter(Mandatory = $false, ParameterSetName = 'AppRegistration')]
    [string[]]
    $ExcludeHandlers,
    [Parameter(Mandatory = $false, ParameterSetName = 'AppRegistration')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [switch]
    $FixGroups,
    [Parameter(Mandatory = $false, ParameterSetName = 'AppRegistration')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [switch]
    $SkipCtHub,
    [Parameter(Mandatory = $false, ParameterSetName = 'AppRegistration')]
    [Parameter(Mandatory = $false, ParameterSetName = 'Interactive')]
    [Hashtable]
    $ChangeFieldTypes,
    [Parameter(ParameterSetName = 'AppRegistration')]
    [Parameter(ParameterSetName = 'Interactive')]
    [string]
    $StatusFilePath = "template-application-status.csv"
)

try {
    $ErrorActionPreference = 'Stop'
    $env:ENTRAID_APP_ID = $EntraAppId
    Start-Transcript -Path "$PSScriptRoot\Logs\Invoke-ProvisionDmsSites-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
    if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
        throw "This script requires PowerShell 7 or later."
    }
    Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Invoke-ProvisionDmsSites-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"
    $sites = @(Import-Csv -Path $CsvPath -Encoding UTF8)
    if ($sites.Count -eq 0) {
        throw "The CSV file is empty."
    }
    $expectedProperties = @("SiteUrl", "Create", "Type", "Alias", "Language", "Name", "Code", "IsActive", "EndCustomers", "DirectCustomers", "IsMigrated", "Product", "ProductLine", "Geography", "PLOwnerships")
    $expectedProperties | ForEach-Object {
        if ($sites[0].PSObject.Properties.Name -notcontains $_) {
            throw "The CSV file does not contain the $_ column."
        }
    }
    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    $expectedConfigProperties = @("lcid", "owner", "hubId", "customerTermSetId", "productTermSetId", "productLineTermSetId", "geographyTermSetId", "publishEventEndpoint", "termGroupId", "appIds")
    $expectedConfigProperties | ForEach-Object {
        if ($config.PSObject.Properties.Name -notcontains $_) {
            throw "The config file does not contain the $_ property."
        }
    }
    $rootUrl = $sites[0].SiteUrl -replace "(https:\/\/[^\/]+).*", '$1'
    $connected = $false
    
    $results = @()
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tChecking if the status file exists..."
    if (-not (Test-Path $StatusFilePath)) {
        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tStatus file not found. New file will be created."
    }
    else {
        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tStatus file found. Reading file..."
        $results = @(Import-Csv $StatusFilePath -Encoding UTF8)
    }

    $sites | Where-Object { $site = $_; (@($results | Where-Object{ $_."Site Url" -eq $site })).Count -eq 0 } | ForEach-Object{
        $results += [PSCustomObject]@{
            "Site Name" = $_.SiteUrl -split "/" | Select-Object -Last 1;
            "Status" = "Pending";
            "Last Run" = "";
            "Error" = "";
            "Site Url" = $_.SiteUrl;
            "Parameters" = $_ | ConvertTo-Json -Compress;
        }
    }
    
    $sites | ForEach-Object {
        $site = $_
        $siteName = $site.SiteUrl -split "/" | Select-Object -Last 1
        $resultRow = $results | Where-Object { $_."Site Name" -eq $siteName } | Select-Object -First 1
        if($null -eq $resultRow){
            $resultRow = [PSCustomObject]@{
                "Site Name" = $siteName;
                "Status" = "Pending";
                "Last Run" = "";
                "Error" = "";
                "Site Url" = $site.SiteUrl;
            }
        }else{
            if($resultRow.Status -eq "Success"){
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tSite already processed. Skipping..."
                return
            }
        }

        try {
            write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`t$($site | ConvertTo-Json -Compress)"
            if ($site.Create -eq "Yes") {
                if ($connected -eq $false) {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnecting to root"
                    if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
                        Connect-PnPOnline $rootUrl -Interactive -ClientId $ClientId
                    }
                    else {
                        Connect-PnPOnline $rootUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
                    }
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnected " -ForegroundColor Green
                    $connected = $true
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tCreating site"
                New-PnPSite -Type TeamSiteWithoutMicrosoft365Group -Title "DOC4A - $($site.Name)" -Url $site.SiteUrl -Lcid $config.lcid -Owner $config.owner -HubSiteId $config.hubId -Wait
                Set-PnPTenantSite -Identity $site.SiteUrl -DenyAddAndCustomizePages:$false
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSite created" -ForegroundColor Green
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnecting to site"
                if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
                    Connect-PnPOnline $site.SiteUrl -Interactive -ClientId $ClientId
                }
                else {
                    Connect-PnPOnline $site.SiteUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnected " -ForegroundColor Green
            }
            else {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnecting to site"
                if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
                    Connect-PnPOnline $site.SiteUrl -Interactive -ClientId $ClientId
                }
                else {
                    Connect-PnPOnline $site.SiteUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnected " -ForegroundColor Green
            }
            $webServerRelativeUrl = Get-PnPWeb | Select-Object -ExpandProperty ServerRelativeUrl
            $webUrl = Get-PnPWeb | Select-Object -ExpandProperty Url

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tJoining Hub Site"
            try {
                #Must be admin of the hub site and current site
                Invoke-PnPSPRestMethod -Method Post -Url "$webUrl/_api/site/JoinHubSite('$($config.hubId)')" -Content ""
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tJoined Hub Site" -ForegroundColor Green
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                $resultRow.Status = "Error"
                if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                    $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Join Hub Site]: $($_.Exception.Message)"
                }
                else{
                    $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Join Hub Site]: $($_.Exception.Message)"
                }
            }

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tGranting Application Permissions"
            $config.appIds | ForEach-Object {
                Grant-PnPAzureADAppSitePermission -AppId $_ -Permissions "fullcontrol" -DisplayName $_ | Out-Null
            }
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tApplication Permissions Granted" -ForegroundColor Green

            if (($null -ne $config.additionalSiteColAdmins) -and (@($config.additionalSiteColAdmins).Count -gt 0)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tAdding additional site collection admins"
                Add-PnPSiteCollectionAdmin -Owners @($config.additionalSiteColAdmins)
            }

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRenaming Admin group"
            Set-PnPGroup -Identity (Get-PnPGroup -AssociatedOwnerGroup) -Title "Project Administrators"
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tAdmin group renamed" -ForegroundColor Green

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRenaming Members group"
            Set-PnPGroup -Identity (Get-PnPGroup -AssociatedMemberGroup) -Title "Project Contributors"
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tMembers group renamed" -ForegroundColor Green

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRenaming Visitors group"
            Set-PnPGroup -Identity (Get-PnPGroup -AssociatedVisitorGroup) -Title "Project Visitors"
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tVisitors group renamed" -ForegroundColor Green

            if ($null -ne $config.SPFxAppId) {
                # Install the app
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tInstalling app"
                try {
                    try {
                        Install-PnPApp -Identity $config.SPFxAppId -Wait -ErrorAction Stop   
                    }
                    catch {
                        if ($_.Exception.Message -like "*already*") {
                            Update-PnPApp -Identity $config.SPFxAppId -ErrorAction Stop   
                        }
                        else {
                            throw $_
                        }
                    } 
                }
                catch {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                    $resultRow.Status = "Error"
                    if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                        $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Install App]: $($_.Exception.Message)"
                    }
                    else{
                        $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Install App]: $($_.Exception.Message)"
                    }
                }
            }else{
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tNo SPFx app to install"
            }

            if (-not [string]::IsNullOrWhiteSpace($site.EndCustomers)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRetrieving end customer term ids"
                $endCustomerLabels = $site.EndCustomers -split ';'
                $endCustomerTerms = $endCustomerLabels | ForEach-Object {
                    $term = Get-PnPTerm -Recursive -TermSet $config.customerTermSetId -TermGroup $config.termGroupId -Identity $_.Trim() -ErrorAction SilentlyContinue
                    if ($null -eq $term) {
                        throw "The term $_ does not exist in the End Customers term set."
                    }
                    $term
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tEnd customer term ids retrieved" -ForegroundColor Green
            }
            
            if (-not [string]::IsNullOrWhiteSpace($site.DirectCustomers)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRetrieving direct customer term ids"
                $directCustomerLabels = $site.DirectCustomers -split ';'
                $directCustomerTerms = $directCustomerLabels | ForEach-Object {
                    $term = Get-PnPTerm -Recursive -TermSet $config.customerTermSetId -TermGroup $config.termGroupId -Identity $_.Trim() -ErrorAction SilentlyContinue
                    if ($null -eq $term) {
                        throw "The term $_ does not exist in the Direct Customers term set."
                    }
                    $term
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tDirect customer term ids retrieved" -ForegroundColor Green
            }
            if (-not [string]::IsNullOrWhiteSpace($site.PLOwnerships)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRetrieving product line ownership term ids"
                $productLineOwnershipLabels = $site.PLOwnerships -split ';'
                $productLineOwnershipTerms = $productLineOwnershipLabels | ForEach-Object {
                    $term = Get-PnPTerm -Recursive -TermSet $config.productLineTermSetId -TermGroup $config.termGroupId -Identity $_.Trim() -ErrorAction SilentlyContinue
                    if ($null -eq $term) {
                        throw "The term $_ does not exist in the Product Line Ownership term set."
                    }
                    $term
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProduct line ownership term ids retrieved" -ForegroundColor Green
            }
            if (-not [string]::IsNullOrWhiteSpace($site.Product)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRetrieving product term id"
                $productLabel = $site.Product
                $productTerm = Get-PnPTerm -Recursive -TermSet $config.productTermSetId -TermGroup $config.termGroupId -Identity $site.Product.Trim() -ErrorAction SilentlyContinue
                if ($null -eq $productTerm) {
                    throw "The term $productLabel does not exist in the Products term set."
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProduct term id retrieved" -ForegroundColor Green
            }

            if (-not [string]::IsNullOrWhiteSpace($site.ProductLine)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRetrieving product line term id"
                $productLineLabel = $site.ProductLine
                $productLineTerm = Get-PnPTerm -Recursive -TermSet $config.productLineTermSetId -TermGroup $config.termGroupId -Identity $site.ProductLine.Trim() -ErrorAction SilentlyContinue
                if ($null -eq $productLineTerm) {
                    throw "The term $productLineLabel does not exist in the Products term set."
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProduct line term id retrieved" -ForegroundColor Green
            }

            if (-not [string]::IsNullOrWhiteSpace($site.Geography)) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tRetrieving geography term id"
                $geographyLabel = $site.Geography
                $geographyTerm = Get-PnPTerm -TermSet $config.geographyTermSetId -TermGroup $config.termGroupId -Identity $site.Geography.Trim() -Recursive -ErrorAction SilentlyContinue
                if ($null -eq $geographyTerm) {
                    throw "The term $geographyLabel does not exist in the Geography term set."
                }
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tGeography term id retrieved" -ForegroundColor Green
            }
            
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tEnsuring user"
            $web = Get-PnPWeb
            $user = $web.EnsureUser($site.User)
            $context = Get-PnPContext
            $context.Load($user)
            Invoke-PnPQuery
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tUser ensured" -ForegroundColor Green

            $ContentTypes = @("0x0120D5200010BA2408052EEF418E580CFAEB67F318", "0x010100E0A0DF7BA0770847845D57B6354C874F", "0x0101001E991B1623BDFD4BB52821D67A2C29B1", "0x0101002A7F1562542CE54B98FE1AA6F49342E1", "0x010100A81E9300A359D34E8F69B3F892DB64A1", "0x010100DB0B41B9684542439774BE5B1D130C1D", "0x0100B5FD10A125A25E4B943387A70E5BE5D0")
            if (!$SkipCtHub) {
                $ContentTypes | ForEach-Object {
                    try {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tAdding content type from hub: $_"
                        Add-PnPContentTypesFromContentTypeHub -ContentType $_ | Out-Null
                        Set-PnPContentType -Identity $_ -ReadOnly $false | Out-Null
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tContent type added: $_" -ForegroundColor Green   
                    }
                    catch {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }

            if($null -ne $ChangeFieldTypes){
                $ChangeFieldTypes.Keys | ForEach-Object {
                    try {
                        $fieldName = $_
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tChanging '$_' field type to $($ChangeFieldTypes[$_])"
                        $field = Get-PnPField -Identity $_
                        $field.FieldTypeKind = [Microsoft.SharePoint.Client.FieldType]::Parse([Microsoft.SharePoint.Client.FieldType], $ChangeFieldTypes[$_])
                        $field.Update()
                        Invoke-PnPQuery
                        Set-PnPField -Identity $_ -Values @{ "Title" = $field.Title } -UpdateExistingLists
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tField type changed for $($_)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                        $resultRow.Status = "Error"
                        if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                            $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Change Field Type $fieldName]: $($_.Exception.Message)"
                        }
                        else{
                            $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Change Field Type $fieldName]: $($_.Exception.Message)"
                        }
                    }

                }
            }

            $isActive = "False"
            if (($site.IsActive -like "Yes") -or ($site.IsActive -like "True") -or ($site.IsActive -like "1")) {
                $isActive = "True"
            }
            $isMigrated = "False"
            if (($site.IsMigrated -like "Yes") -or ($site.IsMigrated -like "True") -or ($site.IsMigrated -like "1")) {
                $isMigrated = "True"
            }
            $paramsWithNulls = @{
                ProjectType                                      = [System.Web.HttpUtility]::HtmlEncode($site.Type);
                ProjectAlias                                     = [System.Web.HttpUtility]::HtmlEncode($site.Alias);
                ProjectLanguage                                  = $site.Language;
                ProjectName                                      = [System.Web.HttpUtility]::HtmlEncode($site.Name);
                ProjectCode                                      = $site.Code;
                IsActive                                         = $isActive;
                ModifiedByDisplayName                            = $user.Title;
                ModifiedByEmail                                  = $user.Email;
                ModifiedDate                                     = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ");
                EndCustomerLabels                                = $endCustomerLabels -join "|";
                EndCustomerGuids                                 = $endCustomerTerms.Id -join "|";
                DirectCustomerLabels                             = $directCustomerLabels -join "|";
                DirectCustomerGuids                              = $directCustomerTerms.Id -join "|";
                CreatedByDisplayName                             = $user.Title;
                CreatedByEmail                                   = $user.Email;
                CreatedDate                                      = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ");
                IsMigrated                                       = $isMigrated;
                PLOwnershipLabels                                = $productLineOwnershipLabels -join "|";
                PLOwnershipGuids                                 = $productLineOwnershipTerms.Id -join "|";
                ProductLabel                                     = $productLabel;
                ProductGuid                                      = $productTerm.Id;
                ProductLineLabel                                 = $productLineLabel;
                ProductLineGuid                                  = $productLineTerm.Id;
                GeographyLabel                                   = $geographyLabel;
                GeographyGuid                                    = $geographyTerm.Id;
                DocumentFormProperties                           = $config.documentFormProperties;
                ProofreadingCustomActionProps                    = $config.proofreadingCustomActionProps;
                AvvaCustomActionProps                            = $config.avvaCustomActionProps;
                CancelCustomActionProps                          = $config.cancelCustomActionProps;
                RevisionMenuCustomActionProps                    = $config.revisionMenuCustomActionProps;
                CreateRevisionCustomActionProps                  = $config.createRevisionCustomActionProps;
                DraftCustomActionProps                           = $config.draftCustomActionProps;
                DmlCustomUiChangesAppCustomizerCustomActionProps = $config.dmlCustomUiChangesAppCustomizerCustomActionProps;
                MassImportStatusProperties                       = $config.massImportStatusProperties;
                SPFxAppId                                        = $config.SPFxAppId;
                Doc4aClientId                                    = $config.doc4aClientId;
                ProjectConfEndpoint                              = $config.projectConfEndpoint;
                PPEnvId                                          = $config.ppEnvId;
                HubId                                            = $config.hubId;
            }
            $params = @{}
            $paramsWithNulls.Keys | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($paramsWithNulls[$_])) {
                    $params.Add($_, $paramsWithNulls[$_])
                }
            }

            if ($FixGroups) {
                try {
                    $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup
                    $projectAdmins = Get-PnPGroup -Identity "Project Administrators"
                    if ($ownerGroup.Id -ne $projectAdmins.Id) {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tFixing Project Administrators group"
                        if ($projectAdmins.Users.Count -gt 0) {
                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tWARNING`t$($site.Name)`tProject Administrators group is not empty" -ForegroundColor Yellow
                        }
                        else {
                            Remove-PnPGroup -Identity "Project Administrators" -Force
                            $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup
                            Set-PnPGroup -Identity $ownerGroup -Title "Project Administrators"
                        }
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProject Administrators group fixed" -ForegroundColor Green
                    }
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tFixing Project Contributors group"
                    $membersGroup = Get-PnPGroup -AssociatedMemberGroup
                    $projectMembers = Get-PnPGroup -Identity "Project Contributors"
                    if ($membersGroup.Id -ne $projectMembers.Id) {
                        if ($membersGroup.Users.Count -gt 0) {
                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tWARNING`t$($site.Name)`tProject Members group is not empty" -ForegroundColor Yellow
                        }
                        Set-PnPGroup -Identity "Project Contributors" -SetAssociatedGroup Members
                        Remove-PnPGroup -Identity $membersGroup -Force
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProject Contributors group fixed" -ForegroundColor Green
                    }

                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tFixing Project Visitors group"
                    $visitorsGroup = Get-PnPGroup -AssociatedVisitorGroup
                    $projectVisitors = Get-PnPGroup -Identity "Project Visitors"
                    if ($visitorsGroup.Id -ne $projectVisitors.Id) {
                        if ($projectVisitors.Users.Count -gt 0) {
                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tWARNING`t$($site.Name)`tProject Visitors group is not empty" -ForegroundColor Yellow
                        }
                        else {
                            Remove-PnPGroup -Identity "Project Visitors" -Force
                            Set-PnPGroup -Identity $visitorsGroup -Title "Project Visitors"
                            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProject Visitors group fixed" -ForegroundColor Green   
                        }
                    }
                }
                catch {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                }
            }

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$($site.Name)`tTemplate parameters: $($params | ConvertTo-Json -Compress)"
            if ($null -ne $Handlers) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tApplying site template. Handlers: $($Handlers -join ', ')"
                Invoke-PnPSiteTemplate -Path "$PSScriptRoot\dms-project\dms-project.xml" -Parameters $params -Handlers $Handlers
            }
            else {
                if ($null -ne $ExcludeHandlers) {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tApplying site template. Excluding handlers: $($ExcludeHandlers -join ', ')"
                    Invoke-PnPSiteTemplate -Path "$PSScriptRoot\dms-project\dms-project.xml" -Parameters $params -ExcludeHandlers $ExcludeHandlers
                }
                else {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tApplying site template"
                    Invoke-PnPSiteTemplate -Path "$PSScriptRoot\dms-project\dms-project.xml" -Parameters $params
                }
            }
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSite template applied" -ForegroundColor Green

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSetting site banner"
            Set-PnPPage -Identity "-.aspx" -HeaderType Custom -ServerRelativeImageUrl "$webServerRelativeUrl/SiteAssets/28157-bannerhome.png" | Out-Null
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSite banner set" -ForegroundColor Green

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tReordering fields"
            $orderedFieldNames = ($config.documentFormProperties | ConvertFrom-Json).sectionsToFieldsMap.PSObject.Properties | ForEach-Object { $_.Value }
            $ContentType = Get-PnPContentType -Identity "Technical Document"
            $FieldLinks = Get-PnPProperty -ClientObject $ContentType -Property "FieldLinks"
            $FieldLinks.Reorder($orderedFieldNames)
            $ContentType.Update($true)
            Invoke-PnPQuery
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tFields reordered" -ForegroundColor Green

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tUpdating form properties"
            Set-PnPContentType -Identity "Technical Document" -FormClientSideComponentProperties $config.documentFormProperties -UpdateChildren | Out-Null
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tForm properties updated" -ForegroundColor Green

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tUpdating baseline properties"
            Set-PnPContentType -Identity "Baseline" -FormClientSideComponentId "d71b7328-47d9-4cf3-ab95-3c3f4ae7641d" -FormClientSideComponentProperties $config.baselineProperties -UpdateChildren | Out-Null
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tbaseline properties updated" -ForegroundColor Green

            #Disable grid edit in baseline and draft
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tDisabling grid edit in baseline"
            try {
                Set-PnPList -Identity "Baselines" -DisableGridEditing $true -DisableCommenting $true | Out-Null
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tGrid edit disabled" -ForegroundColor Green
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                $resultRow.Status = "Error"
                if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                    $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Disable Grid Editing Baseline]: $($_.Exception.Message)"
                }
                else{
                    $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Disable Grid Editing Baseline]: $($_.Exception.Message)"
                }
            }
            try {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tDisabling grid edit in draft"
                Set-PnPList -Identity "Draft" -DisableGridEditing $true -DisableCommenting $true | Out-Null
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tGrid edit disabled" -ForegroundColor Green
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                $resultRow.Status = "Error"
                if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                    $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Disable Grid Editing Draft]: $($_.Exception.Message)"
                }
                else{
                    $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Disable Grid Editing Draft]: $($_.Exception.Message)"
                }
            }

            try {
                #Set group owners
                write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSetting group owners"
                Set-PnPGroup -Identity "Project Contributors" -Owner "Project Administrators"
                Set-PnPGroup -Identity "Project Visitors" -Owner "Project Administrators"
                set-PnPGroup -Identity "Project Administrators" -Owner "Project Administrators"
                write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tGroup owners set" -ForegroundColor Green
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                $resultRow.Status = "Error"
                if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                    $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Set Group Owners]: $($_.Exception.Message)"
                }
                else{
                    $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Set Group Owners]: $($_.Exception.Message)"
                }
            }

            if ($EnableWebhook) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tEnabling webhook"
                Add-PnPWebhookSubscription -List "Draft" -NotificationUrl $config.publishEventEndpoint
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tWebhook enabled" -ForegroundColor Green
            }

            $docTemplate = @(
                [PSCustomObject]@{Template = "Appendix Document Template.docx"; ContentType = "Appendix"; },
                [PSCustomObject]@{Template = "Comments Document Template.docx"; ContentType = "Comments Document"; },
                [PSCustomObject]@{Template = "Native Document Template.docx"; ContentType = "Native Document"; },
                [PSCustomObject]@{Template = "Rendition Document Template.docx"; ContentType = "Rendition Document"; },
                [PSCustomObject]@{Template = "Translation Document Template.docx"; ContentType = "Translation Document"; }
            )
            $docTemplate | ForEach-Object {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tUpdating template URL for CT $($_.ContentType)"
                $contentType = Get-PnPContentType -Identity $_.ContentType
                $contentType.DocumentTemplate = "$webServerRelativeUrl/DocumentTemplates/$($_.Template)"
                $contentType.Update($true)
                Invoke-PnPQuery
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tTemplate URL updated for CT $_" -ForegroundColor Green
            }

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSetting indexed columns"
            $config.fieldIndexes | ForEach-Object {
                $fieldName = $_.field
                $field = Get-PnPField -Identity $_.field -List $_.list
                if (!$field.Indexed) {
                    try {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSetting indexed column $($_.field) in list $($_.list)"
                        $field.Indexed = $true
                        $field.Update()
                        Invoke-PnPQuery
                    }
                    catch {
                        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
                        $resultRow.Status = "Error"
                        if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                            $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Indexing field $fieldName]: $($_.Exception.Message)"
                        }
                        else{
                            $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Indexing field $fieldName]: $($_.Exception.Message)"
                        }
                    }

                }
            }
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tIndexed columns set" -ForegroundColor Green
        }
        catch {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
            $resultRow.Status = "Fatal Error"
            if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Fatal]: $($_.Exception.Message)"
            }
            else{
                $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Fatal]: $($_.Exception.Message)"
            }
        }finally{
            $resultRow."Last Run" = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $results = @($results | Where-Object { $_."Site Name" -ne $siteName })
            $results += $resultRow
            try {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSaving status file..."
                $results | Export-Csv $StatusFilePath -Encoding UTF8 -NoTypeInformation -Force
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`t$($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
catch {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`t$($_.Exception.Message)" -ForegroundColor Red
    Write-Error $_
}
finally {
    Stop-Transcript
    Set-PnPTraceLog -Off
}