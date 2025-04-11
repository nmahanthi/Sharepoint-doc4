$ErrorActionPreference = 'Stop'
Connect-PnPOnline -Url "https://alstomgrouppp.sharepoint.com" -Interactive -ClientId "7016db60-0b9c-4f00-93a4-293a236b490d"
npm run publish
$siteUrls = Invoke-PnPSearchQuery -Query 'Path:"https://alstomgrouppp.sharepoint.com/sites/DOC4A-UAT-*" contentclass:STS_Site' | Select-Object -ExpandProperty ResultRows | %{ $_["SPWebUrl"]}
foreach ($siteUrl in $siteUrls) {
    Write-Host "Pushing to UAT site $siteUrl"
    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId "7016db60-0b9c-4f00-93a4-293a236b490d"
    Add-PnPApp -Path "sharepoint\solution\dms-sp-client-app.sppkg" -Overwrite -Publish
    try {
        Install-PnPApp -Identity "b5c07139-9e25-4a5a-b425-308123f10150" -Wait -ErrorAction Stop   
    }
    catch {
        if($_.Exception.Message -like "*already*"){
            Update-PnPApp -Identity "b5c07139-9e25-4a5a-b425-308123f10150" -ErrorAction Stop   
        }else{
            throw $_
        }
    }
}

Add-Type -AssemblyName System.Speech
$synth = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synth.Speak("Push to UAT completed")