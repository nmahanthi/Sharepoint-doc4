Connect-PnPOnline -Url "https://alstomgrouppp.sharepoint.com/sites/DOC4A-INT" -Interactive -ClientId "7016db60-0b9c-4f00-93a4-293a236b490d"
npm run publish
Add-PnPApp -Path "sharepoint\solution\dms-sp-client-app.sppkg" -Scope Site -Overwrite -Publish
try {
    Install-PnPApp -Identity "b5c07139-9e25-4a5a-b425-308123f10150" -Scope Site -Wait -ErrorAction Stop   
}
catch {
    if ($_.Exception.Message -like "*already*") {
        Update-PnPApp -Identity "b5c07139-9e25-4a5a-b425-308123f10150" -Scope Site -ErrorAction Stop   
    }
    else {
        throw $_
    }
}
Add-Type -AssemblyName System.Speech
$synth = New-Object -TypeName System.Speech.Synthesis.SpeechSynthesizer
$synth.Speak("Push to INT completed")