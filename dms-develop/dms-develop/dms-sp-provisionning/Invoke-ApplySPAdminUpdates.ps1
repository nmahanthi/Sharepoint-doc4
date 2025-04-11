[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]
    $CsvPath="$PSScriptRoot\doc4a-sites.csv",
    [Parameter(Mandatory = $false)]
    [string]
    $ClientId = "2f4c083a-43ef-4310-8aeb-200b3cf40de9",
    [Parameter(Mandatory = $false)]
    [string]
    $CertificateThumbprint = "48BC8915B426DE46F261ADE8726D17BF05A44DDB",
    [Parameter(Mandatory = $false)]
    [string]
    $Tenant = "hakachou.onmicrosoft.com"
)

try {
    $ErrorActionPreference = 'Stop'
    $env:ENTRAID_APP_ID = $ClientId
    Start-Transcript -Path "$PSScriptRoot\Logs\Invoke-ApplySPAdminUpdates-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
    if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
        throw "This script requires PowerShell 7 or later."
    }
    Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Invoke-ApplySPAdminUpdates-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"
    $sites = @(Import-Csv -Path $CsvPath -Encoding UTF8)
    if ($sites.Count -eq 0) {
        throw "The CSV file is empty."
    }
    $expectedProperties = @("SiteUrl")
    $expectedProperties | ForEach-Object {
        if ($sites[0].PSObject.Properties.Name -notcontains $_) {
            throw "The CSV file does not contain the $_ column."
        }
    }
    $sites | ForEach-Object {
        $site = $_
        try {
            write-host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tProcessing site"

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnecting to site. Url $($site.SiteUrl) ClientId $ClientId Thumbprint $CertificateThumbprint Tenant $Tenant"
            Connect-PnPOnline -Url $site.SiteUrl -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant $Tenant
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tConnected " -ForegroundColor Green
            
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tSet no-script flag to false"
            try {
                Set-PnPSite -NoScriptSite $false
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tNo-script flag set to false" -ForegroundColor Green                
            }
            catch {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
            }

        }
        catch {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($site.Name)`t$($_.Exception.Message)" -ForegroundColor Red
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