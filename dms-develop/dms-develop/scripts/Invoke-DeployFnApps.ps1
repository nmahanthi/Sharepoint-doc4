[CmdletBinding()]
param (
    [Parameter()]
    [switch]
    $UpdateSettings,
    [Parameter()]
    [switch]
    $NoCodeDeployment,
    [Parameter(Mandatory = $true)]
    [ValidateSet('Interactive', 'ServicePrincipal', 'ManagedIdentity', 'None')]
    [string]
    $LoginMode,
    [Parameter()]
    [switch]
    $NoBuild,
    [Parameter(Mandatory = $true)]
    [ValidateSet('int','uat', 'ppr', 'prod')]
    [string]
    $Env,
    [Parameter()]
    [ValidateSet(
        'dms-az-baseline-export', 
        'dms-az-doc-postprocessing', 
        'dms-az-event-publishing',
        'dms-az-mass-import',
        'dms-az-update-reference',
        'dms-az-user-api'
    )]
    [string[]]
    $Projects,
    [Parameter()]
    [ValidateSet(
        'dms-az-baseline-export', 
        'dms-az-doc-postprocessing', 
        'dms-az-event-publishing',
        'dms-az-mass-import',
        'dms-az-update-reference',
        'dms-az-user-api'
    )]
    [string[]]
    $ExcludeProjects
)
$ErrorActionPreference = "Stop"
$location = Get-Location
start-transcript -path "$PSScriptRoot\logs\deploy-fnapps-$(Get-Date -Format 'yyyyMMddHHmmss').log"
$configPath = "$PSScriptRoot\Invoke-DeployFnApps.$Env.json"
if (-not (Test-Path $configPath)) {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`tConfiguration file not found" -ForegroundColor Red
    exit 1
}
$config = Get-Content $configPath | Out-String | ConvertFrom-Json
try {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tLogin to Azure" -ForegroundColor Green
    switch ($LoginMode) {
        'Interactive' {
            az config set core.login_experience_v2=off
            az login
        }
        'ServicePrincipal' {
            $requiredParams = @('clientId', 'tenantId', 'subscriptionId')
            $requiredParams | ForEach-Object {
                if (-not $config.$_) {
                    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`tMissing required parameter: $_" -ForegroundColor Red
                    exit 1
                }
            }
            $secret = $env:DOC4A_PROV_SECRET
            if (-not $secret) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`tSecret not found in environment variable" -ForegroundColor Red
                exit 1
            }
            az login --service-principal -u $config.clientId -p $secret --tenant $config.tenantId
        }
        'ManagedIdentity' {
            az login --identity
        }
        'None' {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSkipping login"
        }
        default {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`tInvalid login mode" -ForegroundColor Red
            exit 1
        }
    }
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tAzure login successful" -ForegroundColor Green

    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSetting subscription" -ForegroundColor Green
    az account set --subscription $config.subscriptionId
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`tN/A`tSubscription set" -ForegroundColor Green

    if (!$NoCodeDeployment -and !$NoBuild) {
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`tBuilding common module"
        Set-Location -Path "$PSScriptRoot\..\dms-az-common-node"
        npm i
        npm run build
        Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($site.Name)`Common module built " -ForegroundColor Green
    }
    
    $i = 0
    $functions = $config.functions
    if($null -ne $Projects) {
        $functions = $functions | Where-Object { @($Projects) -contains $_.folder }
    }
    if($null -ne $ExcludeProjects) {
        $functions = $functions | Where-Object { @($ExcludeProjects) -notcontains $_.folder }
    }
    $functions | ForEach-Object {
        $folderName = $_.folder
        $appName = $_.appName
        try {
            Set-Location -Path "$PSScriptRoot\..\$folderName"
            if (!$NoCodeDeployment -and !$NoBuild) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($folderName)`tBuilding"
                npm i
                npm run build
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($folderName)`t$folderName built " -ForegroundColor Green
            }

            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($folderName)`tPublishing $appName"
            if (!$NoCodeDeployment) {
                func azure functionapp publish $appName
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($folderName)`t$_ published code " -ForegroundColor Green
                return
            }
            if($UpdateSettings) {
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($folderName)`tUpdating settings for $appName"
                func azure functionapp publish $appName --publish-settings-only -o
                Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tINFO`t$($folderName)`tSettings updated for $appName" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`t$($_.Exception.Message)" -ForegroundColor Red
            Write-Error $_
        }
        finally {
            $i++
        }

    }
}
catch {
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`tERROR`tN/A`t$($_.Exception.Message)" -ForegroundColor Red
    Write-Error $_
}
finally {
    Stop-Transcript
    Set-Location -Path $location
}
