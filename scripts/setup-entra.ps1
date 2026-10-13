# Creates the Entra app registration m365-mcp needs, using the Azure CLI (Windows PowerShell).
#
#   1. Install the Azure CLI: https://learn.microsoft.com/cli/azure/install-azure-cli
#   2. az login --allow-no-subscriptions     (as someone who can register apps)
#   3. .\scripts\setup-entra.ps1 [-AppName "m365-mcp"]
param([string]$AppName = "m365-mcp")
$ErrorActionPreference = "Stop"

$GraphApi = "00000003-0000-0000-c000-000000000000"
$Scopes = @("User.Read", "Mail.ReadWrite", "Mail.Send", "Files.ReadWrite.All", "Sites.ReadWrite.All", "Calendars.ReadWrite", "MailboxSettings.Read")

if (-not (Get-Command az -ErrorAction SilentlyContinue)) { throw "The Azure CLI (az) isn't installed." }
$TenantId = az account show --query tenantId -o tsv
if (-not $TenantId) { throw "Run 'az login --allow-no-subscriptions' first." }

$Existing = az ad app list --display-name $AppName --query "[0].appId" -o tsv
if ($Existing) { throw "An app called '$AppName' already exists ($Existing). Use -AppName to pick another name." }

Write-Host "Creating app registration '$AppName' in tenant $TenantId..."
$AppId = az ad app create --display-name $AppName --sign-in-audience AzureADMyOrg `
  --is-fallback-public-client true --public-client-redirect-uris "http://localhost" --query appId -o tsv

foreach ($scope in $Scopes) {
  $ScopeId = az ad sp show --id $GraphApi --query "oauth2PermissionScopes[?value=='$scope'].id | [0]" -o tsv
  az ad app permission add --id $AppId --api $GraphApi --api-permissions "$ScopeId=Scope" --only-show-errors | Out-Null
}
az ad sp create --id $AppId --only-show-errors | Out-Null
Start-Sleep -Seconds 20

az ad app permission admin-consent --id $AppId --only-show-errors 2>$null
if ($LASTEXITCODE -eq 0) {
  $Consent = "Admin consent granted for everyone in the organisation."
} else {
  $Consent = "Admin consent was NOT granted. Ask an administrator to approve it here:`n  https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$AppId"
}

Write-Host ""
Write-Host "Done."
Write-Host "  Client ID:  $AppId"
Write-Host "  Tenant ID:  $TenantId"
Write-Host "  $Consent"
Write-Host ""
Write-Host "Each person then runs:"
Write-Host "  m365-mcp configure --client-id $AppId --tenant-id $TenantId"
Write-Host "  m365-mcp login"
