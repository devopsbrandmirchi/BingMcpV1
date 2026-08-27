param(
  [Parameter(Mandatory = $true)]
  [string]$ProjectId,

  [Parameter(Mandatory = $true)]
  [string]$AppBaseUrl,

  [Parameter(Mandatory = $true)]
  [string]$MicrosoftClientId,

  [Parameter(Mandatory = $true)]
  [string]$MicrosoftClientSecret,

  [Parameter(Mandatory = $true)]
  [string]$McpTokenSecret,

  [Parameter(Mandatory = $true)]
  [string]$OauthStateSecret,

  [Parameter(Mandatory = $true)]
  [string]$TokenEncryptionKey,

  [Parameter(Mandatory = $true)]
  [string]$MicrosoftAdsDeveloperToken,

  [string]$FirestoreProjectId = "",
  [string]$FirestoreDatabaseId = "bing-mcp-v1",
  [string]$Region = "us-central1",
  [string]$Service = "bing-mcp-v1"
)

$ErrorActionPreference = "Stop"

if ($FirestoreDatabaseId -ne "bing-mcp-v1") {
  throw "FIRESTORE_DATABASE_ID must be bing-mcp-v1. Do not use (default) or gconnect-mcp-all-v1."
}

$base = $AppBaseUrl.TrimEnd("/")
$redirect = "$base/oauth/microsoft/callback"
$firestoreProject = if ($FirestoreProjectId) { $FirestoreProjectId } else { $ProjectId }

$pairs = @(
  "APP_BASE_URL=$base",
  "MICROSOFT_CLIENT_ID=$MicrosoftClientId",
  "MICROSOFT_CLIENT_SECRET=$MicrosoftClientSecret",
  "MICROSOFT_REDIRECT_URI=$redirect",
  "MCP_TOKEN_SECRET=$McpTokenSecret",
  "OAUTH_STATE_SECRET=$OauthStateSecret",
  "TOKEN_ENCRYPTION_KEY=$TokenEncryptionKey",
  "MICROSOFT_ADS_DEVELOPER_TOKEN=$MicrosoftAdsDeveloperToken",
  "FIRESTORE_PROJECT_ID=$firestoreProject",
  "FIRESTORE_DATABASE_ID=$FirestoreDatabaseId"
)

Write-Host "Updating Cloud Run environment for $Service (values are not printed)"
gcloud run services update $Service `
  --region $Region `
  --project $ProjectId `
  --update-env-vars ($pairs -join ",")

Write-Host "Environment updated. Revision will start automatically."
Write-Host "Add this Microsoft Entra redirect URI if you have not already:"
Write-Host "  $redirect"
