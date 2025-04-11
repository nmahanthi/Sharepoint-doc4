[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]
    $SiteUrl,
    [Parameter(Mandatory = $true)]
    [string]
    $ClientId,
    [Parameter(Mandatory = $true)]
    [string]
    $ClientSecret,
    [Parameter(Mandatory = $true)]
    [string]
    $PackagePath,
    [Parameter(Mandatory = $false)]
    [string]
    $AppId = "b5c07139-9e25-4a5a-b425-308123f10150",
    [Parameter(Mandatory = $false)]
    [switch]
    $IsSiteScoped
)
$ErrorActionPreference = "Stop"
If(-not(Get-InstalledModule PnP.PowerShell -ErrorAction silentlycontinue)){
    Install-Module PnP.PowerShell -Confirm:$False -Force
}
Write-Host "SiteUrl: $SiteUrl ClientId: $ClientId"
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -ClientSecret $ClientSecret
if ($IsSiteScoped) {
    Add-PnPApp -Path $PackagePath -Overwrite -Publish -Scope Site
    try {
        Install-PnPApp -Identity $AppId -Wait -Scope Site
    }
    catch {
        if ($_.Exception.Message -like "*already*") {
            Update-PnPApp -Identity $AppId -Scope Site   
        }
        else {
            throw $_
        }
    }
}
else {
    Add-PnPApp -Path $PackagePath -Overwrite -Publish
    try {
        Install-PnPApp -Identity $AppId -Wait
    }
    catch {
        if ($_.Exception.Message -like "*already*") {
            Update-PnPApp -Identity $AppId
        }
        else {
            throw $_
        }
    }
}