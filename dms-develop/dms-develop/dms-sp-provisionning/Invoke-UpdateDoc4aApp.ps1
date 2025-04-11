# Parameter help description
[CmdletBinding()]
param (
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [Parameter(ParameterSetName = 'Csv')]
    [string]
    $StatusFilePath = "migrate-project-properties-status.csv",
    [Parameter(ParameterSetName = 'Hub')]
    [string]
    $HubUrl = "https://alstomgroup.sharepoint.com/sites/DOC4A",
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [Parameter(ParameterSetName = 'Csv')]
    [string]
    $ClientId = "ef114265-5835-48b5-9652-f16b6423c91b",
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [Parameter(ParameterSetName = 'Csv')]
    [string]
    $AppId = "b5c07139-9e25-4a5a-b425-308123f10150",
    [Parameter(ParameterSetName = 'Site', Mandatory = $true)]
    [string]
    $SiteUrl,
    [Parameter(ParameterSetName = 'Csv', Mandatory = $true)]
    [string]
    $CsvPath
)
$ErrorActionPreference = 'Stop'
Start-Transcript -Path "$PSScriptRoot\Logs\Invoke-UpdateDoc4aApp-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
    throw "This script requires PowerShell 7 or later."
}
Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Invoke-UpdateDoc4aApp-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"

try {
    $results = @()
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tChecking if the status file exists..."
    if (-not (Test-Path $StatusFilePath)) {
        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tStatus file not found. New file will be created."
    }
    else {
        write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tStatus file found. Reading file..."
        $results = @(Import-Csv $StatusFilePath -Encoding UTF8)
    }
    
    switch ($PsCmdlet.ParameterSetName) {
        "Hub" {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tConnecting to Hub..."
            Connect-PnPOnline $HubUrl -Interactive -ClientId $ClientId
        
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tGetting Hub Id..."
            $hubId = Get-PnPSite -includes Id | Select-Object -ExpandProperty Id | ForEach-Object { $_.ToString() }
        
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tGetting all sites..."
            $sites = invoke-pnpsearchQuery -Query "DepartmentId:$hubId AND ContentClass:STS_Site" -All | Select-Object -ExpandProperty ResultRows | % { $_.SPWebUrl }
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tFound $($sites.Count) sites."
        }
        "Site" {
            $sites = @($SiteUrl)
        }
        "Csv" {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tReading sites from CSV file..."
            $sites = @(Import-Csv $CsvPath -Encoding UTF8 | Select-Object -ExpandProperty SiteUrl)
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tFound $($sites.Count) sites."
        }
        "__AllParameterSets" { 
            throw "This script requires a parameter set to be specified."
        }
    }
    
    $sites | Where-Object { $site = $_; (@($results | Where-Object { $_."Site Url" -eq $site })).Count -eq 0 } | ForEach-Object {
        $results += [PSCustomObject]@{
            "Site Name" = $_ -split "/" | Select-Object -Last 1;
            "Status"    = "Pending";
            "Last Run"  = "";
            "Error"     = "";
            "Site Url"  = $_;
        }
    }

    foreach ($site in $sites) {
        $siteName = $site -split "/" | Select-Object -Last 1
        $resultRow = $results | Where-Object { $_."Site Name" -eq $siteName } | Select-Object -First 1
        if ($null -eq $resultRow) {
            $resultRow = [PSCustomObject]@{
                "Site Name" = $siteName;
                "Status"    = "Pending";
                "Last Run"  = "";
                "Error"     = "";
                "Site Url"  = $site;
            }
        }
        else {
            if ($resultRow.Status -eq "Success") {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tSite already processed. Skipping..."
                continue
            }
        }
        try {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tConnecting to site..."
            Connect-PnPOnline $site -Interactive -ClientId $ClientId

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tInstalling app"
            try {
                try {
                    Install-PnPApp -Identity $AppId -Wait -ErrorAction Stop   
                }
                catch {
                    if ($_.Exception.Message -like "*already*") {
                        Update-PnPApp -Identity $AppId -ErrorAction Stop   
                    }
                    else {
                        throw $_
                    }
                } 
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$siteName`t$($_.Exception.Message)" -ForegroundColor Red
                $resultRow.Status = "Error"
                if ([string]::IsNullOrWhiteSpace($resultRow.Error)) {
                    $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Install App]: $($_.Exception.Message)"
                }
                else {
                    $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - [Install App]: $($_.Exception.Message)"
                }
            }

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tSite processed successfully."
            $resultRow.Status = "Success"
        }
        catch {
            write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$site`t$($_.Exception.Message)"
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
            $results = @($results | Where-Object { $_."Site Name" -ne $siteName })
            $results += $resultRow
        }
    }
}
catch {
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`t$($_.Exception.Message)"
    Write-Error $_
    throw
}
finally {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSaving status file..."
    $results | Export-Csv $StatusFilePath -Encoding UTF8 -NoTypeInformation -Force
    write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tEnd of script."
    Stop-Transcript
    Set-PnPTraceLog -Off
}