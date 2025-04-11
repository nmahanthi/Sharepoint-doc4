# Parameter help description
[CmdletBinding()]
param (
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [string]
    $StatusFilePath = "migrate-project-properties-status.csv",
    [Parameter(ParameterSetName = 'Hub')]
    [string]
    $HubUrl = "https://alstomgroup.sharepoint.com/sites/DOC4A",
    [Parameter(ParameterSetName = 'Hub')]
    [Parameter(ParameterSetName = 'Site')]
    [string]
    $ClientId = "ef114265-5835-48b5-9652-f16b6423c91b",
    [Parameter(ParameterSetName = 'Site', Mandatory = $true)]
    [string]
    $SiteUrl
)
$ErrorActionPreference = 'Stop'
Start-Transcript -Path "$PSScriptRoot\Logs\Invoke-MigrateProjectProperties-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
    throw "This script requires PowerShell 7 or later."
}
Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Invoke-MigrateProjectProperties-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"

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
    
    if([string]::IsNullOrWhiteSpace($SiteUrl))
    {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tConnecting to Hub..."
        Connect-PnPOnline $HubUrl -Interactive -ClientId $ClientId
    
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tGetting Hub Id..."
        $hubId = Get-PnPSite -includes Id | Select-Object -ExpandProperty Id | ForEach-Object { $_.ToString() }
    
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tGetting all sites..."
        $sites = invoke-pnpsearchQuery -Query "DepartmentId:$hubId AND ContentClass:STS_Site" -All | Select-Object -ExpandProperty ResultRows | % { $_.SPWebUrl }
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tFound $($sites.Count) sites."    
    }else{
        $sites = @($SiteUrl)
    }
    
    $sites | Where-Object { $site = $_; (@($results | Where-Object{ $_."Site Url" -eq $site })).Count -eq 0 } | ForEach-Object{
        $results += [PSCustomObject]@{
            "Site Name" = $_ -split "/" | Select-Object -Last 1;
            "Status" = "Pending";
            "Last Run" = "";
            "Error" = "";
            "Site Url" = $_;
        }
    }

    foreach ($site in $sites) {
        $siteName = $site -split "/" | Select-Object -Last 1
        $resultRow = $results | Where-Object { $_."Site Name" -eq $siteName } | Select-Object -First 1
        if($null -eq $resultRow){
            $resultRow = [PSCustomObject]@{
                "Site Name" = $siteName;
                "Status" = "Pending";
                "Last Run" = "";
                "Error" = "";
                "Site Url" = $site;
            }
        }else{
            if($resultRow.Status -eq "Success"){
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tSite already processed. Skipping..."
                continue
            }
        }
        try {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tConnecting to site..."
            Connect-PnPOnline $site -Interactive -ClientId $ClientId

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tGetting property bags..."
            $propertyBags = Get-PnPPropertyBag | Where-Object{$_.Key -like "dms_*"}
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tFound $($propertyBags.Count) property bags."

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tCleaning up project properties list..."
            $itemIds = Get-PnPListItem -List "Project Properties" -Fields "ID" | Select-Object -ExpandProperty ID
            $batchDelete = New-PnPBatch
            foreach($itemId in $itemIds){
                Remove-PnPListItem -List "Project Properties" -Identity $itemId -Batch $batchDelete
            }
            Invoke-PnPBatch -Batch $batchDelete

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tAdding property bags to project properties list..."
            $batchAdd = New-PnPBatch
            foreach($propertyBag in $propertyBags){
                $key = $propertyBag.Key
                $value = $propertyBag.Value
                Add-PnPListItem -List "Project Properties" -Values @{"Title" = $key; "Value" = $value} -Batch $batchAdd
            }
            Invoke-PnPBatch -Batch $batchAdd

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$siteName`tSite processed successfully."
            $resultRow.Status = "Success"
        }
        catch {
            write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$site`t$($_.Exception.Message)"
            $resultRow.Status = "Error"
            if([string]::IsNullOrWhiteSpace($resultRow.Error)){
                $resultRow.Error = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $($_.Exception.Message)"
            }
            else{
                $resultRow.Error = "$($resultRow.Error)`n$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $($_.Exception.Message)"
            }
        }finally{
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