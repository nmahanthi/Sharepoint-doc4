# Parameter help description
[CmdletBinding( SupportsShouldProcess = $true )]
param (
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [string]
    $StatusFilePath = "reset-version-status.csv",
    [Parameter(ParameterSetName = 'Hub')]
    [string]
    $HubUrl = "https://alstomgroup.sharepoint.com/sites/DOC4A",
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [string]
    $ClientId = "ef114265-5835-48b5-9652-f16b6423c91b",
    [Parameter(ParameterSetName = 'Site', Mandatory = $true)]
    [string]
    $SiteUrl,
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [string[]]
    $FilesSiteRelativeUrls = @("/SiteAssets/DML_dashboard.pbix"),
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [string[]]
    $ExcludeUsers = @(),
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [datetime]
    $BeforeDate,
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [datetime]
    $AfterDate
)
$ErrorActionPreference = 'Stop'
Start-Transcript -Path "$PSScriptRoot\Logs\Reset-FilesVersion-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
    throw "This script requires PowerShell 7 or later."
}
Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Reset-FilesVersion-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"

try {
    $results = @()
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tChecking if the status file exists..."
    if (-not (Test-Path $StatusFilePath)) {
        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tStatus file not found. New file will be created."
    }
    else {
        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tStatus file found. Reading file..."
        $results = @(Import-Csv $StatusFilePath -Encoding UTF8)
    }
    
    if ([string]::IsNullOrWhiteSpace($SiteUrl)) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tConnecting to Hub..."
        Connect-PnPOnline $HubUrl -Interactive -ClientId $ClientId
    
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tGetting Hub Id..."
        $hubId = Get-PnPSite -includes Id | Select-Object -ExpandProperty Id | ForEach-Object { $_.ToString() }
    
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tGetting all sites..."
        $sites = invoke-pnpsearchQuery -Query "DepartmentId:$hubId AND ContentClass:STS_Site" -All | Select-Object -ExpandProperty ResultRows | % { $_.SPWebUrl }
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tN/A`tFound $($sites.Count) sites."    
    }
    else {
        $sites = @($SiteUrl)
    }
    
    $sites | ForEach-Object { $site = $_; $FilesSiteRelativeUrls | ForEach-Object {
            [PSCustomObject]@{
                Site                = $site;
                FileSiteRelativeUrl = $_;
            }
        } } | Where-Object { $site = $_.Site; $file = $_.FileSiteRelativeUrl; (@($results | Where-Object { ($_."Site Url" -eq $site) -and ($_."File Url" -eq $file ) })).Count -eq 0 } | ForEach-Object {
        $results += [PSCustomObject]@{
            "Site Name" = $_.Site -split "/" | Select-Object -Last 1;
            "File Name" = $_.FileSiteRelativeUrl -split "/" | Select-Object -Last 1;
            "Status"    = "Pending";
            "Last Run"  = "";
            "Error"     = "";
            "Site Url"  = $_.Site;
            "File Url"  = $_.FileSiteRelativeUrl;
        }
    }
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSaving status file..."
    $results | Export-Csv $StatusFilePath -Encoding UTF8 -NoTypeInformation -Force
    pause

    foreach ($site in $sites) {
        try {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tNA`tConnecting to site..."
            Connect-PnPOnline $site -Interactive -ClientId $ClientId
            $webServerRelativeUrl = Get-PnPWeb | Select-Object -ExpandProperty ServerRelativeUrl
            foreach ($fileSiteRelativeUrl in $FilesSiteRelativeUrls) {
                $siteName = $site -split "/" | Select-Object -Last 1
                $fileName = $fileSiteRelativeUrl -split "/" | Select-Object -Last 1
                $resultRow = $results | Where-Object { ($_."Site Name" -eq $siteName) -and ($_.'File Name' -eq $fileName) } | Select-Object -First 1
                if ($null -eq $resultRow) {
                    $resultRow = [PSCustomObject]@{
                        "Site Name" = $siteName;
                        "File Name" = $fileName;
                        "Status"    = "Pending";
                        "Last Run"  = "";
                        "Error"     = "";
                        "Site Url"  = $site;
                        "File Url"  = $fileSiteRelativeUrl;
                    }
                }
                else {
                    if ($resultRow.Status -eq "Success") {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tWARNING`t$siteName`t$fileName`tFile already processed. Skipping..." -ForegroundColor Yellow
                        continue
                    }
                }
                try {

                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`t$fileName`tProcessing file..."
                    
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tGetting file versions..." -ForegroundColor DarkGray
                    $versions = Get-PnPProperty -ClientObject (Get-PnPFile -ServerRelativeUrl "$webServerRelativeUrl$fileSiteRelativeUrl") -Property Versions
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tFound $($versions.Count) versions." -ForegroundColor DarkGray
                    
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tLoading versions information..." -ForegroundColor DarkGray
                    $context = Get-PnPContext
                    $versions | ForEach-Object { $context.Load($_.CreatedBy) }
                    Invoke-PnPQuery
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tLoaded." -ForegroundColor DarkGray

                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tSelecting version..." -ForegroundColor DarkGray
                    if ($null -ne $BeforeDate) {
                        $versions = $versions | Where-Object { $_.Created -lt $BeforeDate }
                    }
                    if ($null -ne $AfterDate) {
                        $versions = $versions | Where-Object { $_.Created -gt $AfterDate }
                    }
                    if ($null -ne $ExcludeUsers) {
                        $versions = $versions | Where-Object { -not ($ExcludeUsers -contains $_.CreatedBy.Email) }
                    }
                    $version = $versions | Select-Object -Last 1
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tSelected version $($version.VersionLabel) - $($version.CreatedBy.Email) - $($version.Created)" -ForegroundColor DarkGray

                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tRestoring version..."
                    if ($WhatIfPreference) {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`t$siteName`t$fileName`tWhatIf: Restore-PnPFileVersion -Url '$webServerRelativeUrl$fileSiteRelativeUrl' -Identity $($version.Id)"
                        $resultRow.Status = "WhatIf"
                    }
                    else {
                        Restore-PnPFileVersion -Url "$webServerRelativeUrl$fileSiteRelativeUrl" -Identity $version.Id -Force
                    }
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`t$fileName`tFile processed successfully." -ForegroundColor Green
                    $resultRow.Status = "Success"
                }
                catch {
                    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$site`t$fileName`t$($_.Exception.Message)" -ForegroundColor Red
                    $resultRow.Status = "Error"
                    if ([string]::IsNullOrWhiteSpace($resultRow.Error)) {
                        $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $($_.Exception.Message)"
                    }
                    else {
                        $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $($_.Exception.Message)"
                    }
                }
                finally {
                    $resultRow."Last Run" = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
                    $results = @($results | Where-Object { -not (($_."Site Name" -eq $siteName) -and ($_.'File Name' -eq $fileName)) })
                    $results += $resultRow
                    try {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$site`t$fileName`tSaving status file..."
                        $results | Export-Csv $StatusFilePath -Encoding UTF8 -NoTypeInformation -Force
                    }
                    catch {
                        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$site`t$fileName`t$($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
        }
        catch {
            write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$site`tN/A`t$($_.Exception.Message)" -ForegroundColor Red
            $FilesSiteRelativeUrls | ForEach-Object {
                $fileName = $_ -split "/" | Select-Object -Last 1
                $resultRow = $results | Where-Object { ($_."Site Name" -eq $siteName) -and ($_.'File Name' -eq $fileName) } | Select-Object -First 1
                if ($null -ne $resultRow) {
                    $resultRow.Status = "Error"
                }
                else {
                    $resultRow = [PSCustomObject]@{
                        "Site Name" = $siteName;
                        "File Name" = $fileName;
                        "Status"    = "Error";
                        "Last Run"  = "";
                        "Error"     = "";
                        "Site Url"  = $site;
                        "File Url"  = $_;
                    }
                }
                if ([string]::IsNullOrWhiteSpace($resultRow.Error)) {
                    $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $($_.Exception.Message)"
                }
                else {
                    $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $($_.Exception.Message)"
                }
                $results = @($results | Where-Object { -not (($_."Site Name" -eq $siteName) -and ($_.'File Name' -eq $fileName)) })
                $results += $resultRow
            }
            try {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$site`t$fileName`tSaving status file..."
                $results | Export-Csv $StatusFilePath -Encoding UTF8 -NoTypeInformation -Force
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$site`t$fileName`t$($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
catch {
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`t$($_.Exception.Message)" -ForegroundColor Red
    Write-Error $_
    throw
}
finally {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSaving status file..."
    if($WhatIfPreference){
        $results | Export-Csv ($StatusFilePath -replace ".csv", "-WhatIf.csv") -Encoding UTF8 -NoTypeInformation -Force -WhatIf:$false
    }else{
        $results | Export-Csv $StatusFilePath -Encoding UTF8 -NoTypeInformation -Force
    }
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tEnd of script."
    Stop-Transcript
    Set-PnPTraceLog -Off
}