[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $SUBSCRIPTION,
    [Parameter()]
    [string]
    $GROUP,
    [Parameter()]
    [string]
    $WEBAPP,
    [Parameter()]
    [string]
    $ZipPath
)
$ACCOUNT = "tempdeploy$((Get-Random -Minimum 1000 -Maximum 9999).ToString())"
$CONTAINER = "tempdeploy$((Get-Random -Minimum 1000 -Maximum 9999).ToString())"
#az login
#az account set --subscription $SUBSCRIPTION
az extension add --name webapp
az storage account create -n $ACCOUNT -g $GROUP -l westeurope
az storage container create -n $CONTAINER --account-name $ACCOUNT
az storage blob upload -f $ZipPath --account-name $ACCOUNT -c $CONTAINER -n $ACCOUNT
$expiry = Get-Date ((Get-Date).AddMinutes(30)) -UFormat '+%Y-%m-%dT%H:%M:%SZ'
$ZIP_URL = az storage blob generate-sas --full-uri --expiry $expiry --permissions r --account-name $ACCOUNT -c $CONTAINER -n $ACCOUNT
az webapp deploy --name $WEBAPP --resource-group $GROUP --type zip --src-url $ZIP_URL --async false
#az storage container delete -n $CONTAINER --account-name $ACCOUNT