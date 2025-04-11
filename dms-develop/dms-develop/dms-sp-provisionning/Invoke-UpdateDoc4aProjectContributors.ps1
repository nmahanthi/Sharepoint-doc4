# Get the script directory
param (
    [Parameter(Mandatory = $true)]
    [string]$SiteUrlsCsvFilePath
)
#update below configuration based on the environment
$tenantId = "456e1d59-fb6a-46dc-b2e6-f1b37b6d29fb"
$clientId = "7016db60-0b9c-4f00-93a4-293a236b490d"
$thumbprint = "A59D783A79950BFFC17BF968A78DFD638F83B2A0"

$ErrorActionPreference = 'Stop'
Start-Transcript -Path "$PSScriptRoot\Logs\Invoke-UpdateDoc4aApp-$(Get-Date -Format 'yyyyMMdd-HHmmss').log" -IncludeInvocationHeader
if (-not ($PSVersionTable.PSVersion.Major -ge 7)) {
    throw "This script requires PowerShell 7 or later."
}
Set-PnPTraceLog -On -Level Debug -LogFile "$PSScriptRoot\Logs\Invoke-UpdateDoc4aApp-$(Get-Date -Format 'yyyyMMdd-HHmmss')-PnP.log"

# Import CSV data
$table = Import-Csv -Path $SiteUrlsCsvFilePath

# Store results
$result = @()

# Loop through each site in the CSV
foreach ($row in $table) {
    try {
        Write-Host "Connecting to: $($row.SiteUrl)" -ForegroundColor Cyan
        $SourceSiteURL = $row.SiteUrl  

        # Connect to the SharePoint site using certificate-based authentication
        Connect-PnPOnline -Url $SourceSiteURL  -ClientId $clientId -Interactive  # -CertificateThumbprint $thumbprint -Tenant $tenantId

        # Set group name explicitly
        $GroupName = "Project Contributors"

        Write-Host "Updating group settings for: $GroupName" -ForegroundColor Cyan
        Set-PnPGroup -Identity $GroupName -AllowMembersEditMembership $false
        Write-Host "Updated membership settings for $GroupName" -ForegroundColor Green

        # Collect success result
        $result += [PSCustomObject]@{
            SiteUrl   = $row.SiteUrl
            GroupName = $GroupName
            Status    = "Successfully updated membership settings"
            Timestamp = (Get-Date).ToString('dd-MMM-yyyy HH:mm:ss')
        }
    }
    catch {
        Write-Host "Error processing $SourceSiteURL : $_" -ForegroundColor Red

        # Log the failure
        $result += [PSCustomObject]@{
            SiteUrl   = $row.SiteUrl
            GroupName = $GroupName
            Status    = "Error: $_"
            Timestamp = (Get-Date).ToString('dd-MMM-yyyy HH:mm:ss')
        }
    }
}

# Export results to CSV
$LogFilePath = "$PSScriptRoot\Logs\Output_$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
$result | Export-Csv -Path $LogFilePath -NoTypeInformation -Encoding UTF8
Write-Host "✅ Output report exported to CSV file successfully: $LogFilePath" -ForegroundColor Green

Stop-Transcript
Set-PnPTraceLog -Off
