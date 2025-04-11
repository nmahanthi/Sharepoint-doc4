#Example:
#.\Invoke-UpdateCtHub.ps1 -CsvPath .\sample-sites.csv -EnableWebhook -Interactive 
[CmdletBinding()]
param (
    [Parameter(ParameterSetName = 'Interactive')]
    [Parameter(ParameterSetName = 'AppRegistration')]
    [string]
    $HubUrl = "https://alstomgroup.sharepoint.com/sites/contenttypehub",
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
    [Parameter(Mandatory = $true, ParameterSetName = 'Interactive')]
    [switch]
    $Interactive
)

try {
    $ErrorActionPreference = 'Stop'
    $env:ENTRAID_APP_ID = $ClientId
    Start-Transcript -Path "$PSScriptRoot\Logs\Invoke-UpdateCtHub-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
    if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
        throw "This script requires PowerShell 7 or later."
    }
    Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Invoke-UpdateCtHub-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"

    $config = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
    $expectedConfigProperties = @("lcid", "owner", "hubId", "customerTermSetId", "productTermSetId", "productLineTermSetId", "geographyTermSetId", "publishEventEndpoint", "termGroupId", "appIds")
    $expectedConfigProperties | ForEach-Object {
        if ($config.PSObject.Properties.Name -notcontains $_) {
            throw "The config file does not contain the $_ property."
        }
    }
    try {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tConnecting to site"
        if ($PSCmdlet.ParameterSetName -eq 'Interactive') {
            Connect-PnPOnline $HubUrl -Interactive -ClientId $ClientId
        }
        else {
            Connect-PnPOnline $HubUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $Thumbprint
        }
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tConnected " -ForegroundColor Green

        $paramsWithNulls = @{
            DocumentFormProperties                           = $config.documentFormProperties;
            ProofreadingCustomActionProps                    = $config.proofreadingCustomActionProps;
            AvvaCustomActionProps                            = $config.avvaCustomActionProps;
            CancelCustomActionProps                          = $config.cancelCustomActionProps;
            CreateRevisionCustomActionProps                  = $config.createRevisionCustomActionProps;
            DraftCustomActionProps                           = $config.draftCustomActionProps;
            DmlCustomUiChangesAppCustomizerCustomActionProps = $config.dmlCustomUiChangesAppCustomizerCustomActionProps;
            MassImportStatusProperties                       = $config.massImportStatusProperties;
        }
        $params = @{}
        $paramsWithNulls.Keys | ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($paramsWithNulls[$_])) {
                $params.Add($_, $paramsWithNulls[$_])
            }
        }
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tVERBOSE`tTemplate parameters: $($params | ConvertTo-Json -Compress)"
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tApplying site template"
        Invoke-PnPSiteTemplate -Path "$PSScriptRoot\dms-project\dms-ct-hub.xml" -Parameters $params
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tSite template applied" -ForegroundColor Green

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tUpdating form properties"
        Set-PnPContentType -Identity "Technical Document" -FormClientSideComponentProperties $config.documentFormProperties -UpdateChildren | Out-Null
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tForm properties updated" -ForegroundColor Green

        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tUpdating baseline properties"
        Set-PnPContentType -Identity "Baseline" -FormClientSideComponentId "d71b7328-47d9-4cf3-ab95-3c3f4ae7641d" -FormClientSideComponentProperties $config.baselineProperties -UpdateChildren | Out-Null
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tbaseline properties updated" -ForegroundColor Green

        $ContentTypes = @("Technical Document","Native Document","Rendition Document","Translation Document","Comments Document","Appendix","Baseline")
        $ContentTypes | ForEach-Object{
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tPublishing content type: $_"
            Publish-PnPContentType -ContentType $_
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tContent type published: $_" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`t$($_.Exception.Message)" -ForegroundColor Red
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