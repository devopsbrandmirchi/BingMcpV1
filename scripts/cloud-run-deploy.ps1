param(
  [Parameter(Mandatory = $true)]
  [string]$ProjectId,

  [string]$Region = "us-central1",
  [string]$Service = "bing-mcp-v1",
  [string]$Repository = "bing-mcp-v1"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

gcloud config set project $ProjectId

$commitSha = (git -C $root rev-parse HEAD).Trim()
if (-not $commitSha) {
  throw "Could not determine git COMMIT_SHA for the image tag."
}

Write-Host "Submitting Cloud Build for bing-mcp-v1"
gcloud builds submit $root `
  --project $ProjectId `
  --config "$root\cloudbuild.yaml" `
  --substitutions "_REGION=$Region,_SERVICE=$Service,_REPOSITORY=$Repository,COMMIT_SHA=$commitSha"
if ($LASTEXITCODE -ne 0) {
  throw "Cloud Build failed."
}

$serviceUrl = gcloud run services describe $Service `
  --region $Region `
  --project $ProjectId `
  --format "value(status.url)"

Write-Host ""
Write-Host "Cloud Run URL: $serviceUrl"
Write-Host "MCP endpoint:  $serviceUrl/mcp"
Write-Host "Health:        $serviceUrl/health"
Write-Host "Microsoft callback: $serviceUrl/oauth/microsoft/callback"
Write-Host ""
Write-Host "Set APP_BASE_URL to $serviceUrl, add the callback URI on this app's Entra registration, then run cloud-run-set-env.ps1"
