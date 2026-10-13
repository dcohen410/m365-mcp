#!/usr/bin/env bash
# Creates the Entra app registration m365-mcp needs, using the Azure CLI.
#
#   1. Install the Azure CLI: https://learn.microsoft.com/cli/azure/install-azure-cli
#   2. az login --allow-no-subscriptions     (as someone who can register apps)
#   3. ./scripts/setup-entra.sh ["App name"]
#
# Run it once per organisation. It prints the client ID and tenant ID that
# everyone in the organisation then uses with `m365-mcp configure`.
set -euo pipefail

APP_NAME="${1:-m365-mcp}"
GRAPH_API="00000003-0000-0000-c000-000000000000"
SCOPES=(User.Read Mail.ReadWrite Mail.Send Files.ReadWrite.All Sites.ReadWrite.All)

command -v az >/dev/null || { echo "The Azure CLI (az) isn't installed. See the comment at the top of this script."; exit 1; }
TENANT_ID=$(az account show --query tenantId -o tsv) || { echo "Run 'az login --allow-no-subscriptions' first."; exit 1; }

EXISTING=$(az ad app list --display-name "${APP_NAME}" --query "[0].appId" -o tsv)
if [[ -n "${EXISTING}" ]]; then
  echo "An app called '${APP_NAME}' already exists (${EXISTING}). Pass a different name, or delete it first."
  exit 1
fi

echo "Creating app registration '${APP_NAME}' in tenant ${TENANT_ID}..."
APP_ID=$(az ad app create \
  --display-name "${APP_NAME}" \
  --sign-in-audience AzureADMyOrg \
  --is-fallback-public-client true \
  --public-client-redirect-uris "http://localhost" \
  --query appId -o tsv)

echo "Adding delegated Microsoft Graph permissions: ${SCOPES[*]}"
for scope in "${SCOPES[@]}"; do
  SCOPE_ID=$(az ad sp show --id "${GRAPH_API}" --query "oauth2PermissionScopes[?value=='$scope'].id | [0]" -o tsv)
  az ad app permission add --id "${APP_ID}" --api "${GRAPH_API}" --api-permissions "${SCOPE_ID}=Scope" --only-show-errors >/dev/null
done

az ad sp create --id "${APP_ID}" --only-show-errors >/dev/null
# Entra needs a moment before consent can be granted on a brand-new app.
sleep 20

if az ad app permission admin-consent --id "${APP_ID}" --only-show-errors 2>/dev/null; then
  CONSENT="Admin consent granted for everyone in the organisation."
else
  CONSENT="Admin consent was NOT granted (you may not be an admin). Ask a Global or Privileged Role
  Administrator to open this link and approve it:
  https://login.microsoftonline.com/${TENANT_ID}/adminconsent?client_id=${APP_ID}"
fi

cat <<EOF

Done.
  Client ID:  ${APP_ID}
  Tenant ID:  ${TENANT_ID}
  ${CONSENT}

Each person then runs:
  m365-mcp configure --client-id ${APP_ID} --tenant-id ${TENANT_ID}
  m365-mcp login
EOF
