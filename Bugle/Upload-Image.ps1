<#
.SYNOPSIS
    Uploads an image to Bunny CDN and copies the markdown syntax to the clipboard.

.DESCRIPTION
    Reads connection settings from bugle-config.json in the same directory.
    Uploads the image via the Bunny CDN Storage API (HTTP PUT), then copies
    a ready-to-paste markdown image tag to your clipboard.

.PARAMETER ImagePath
    Path to the local image file to upload.

.PARAMETER AltText
    Alt text for the markdown image tag. Defaults to the filename without extension.

.PARAMETER RemoteName
    Override the filename used on the CDN. Defaults to the local filename.
    Spaces are replaced with hyphens automatically either way.

.PARAMETER CssClass
    Optional CSS class to append, e.g. "scout-img-40" or "img-600".
    Produces: ![Alt](url){.class}

.EXAMPLE
    .\Upload-Image.ps1 -ImagePath "C:\Photos\WinterCamp.jpg" -AltText "Winter Camp 2026"

.EXAMPLE
    .\Upload-Image.ps1 -ImagePath ".\hero.png" -AltText "BWCA Trip" -CssClass "img-600"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ImagePath,

    [Parameter(Mandatory = $false)]
    [string]$AltText = "",

    [Parameter(Mandatory = $false)]
    [string]$RemoteName = "",

    [Parameter(Mandatory = $false)]
    [string]$CssClass = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Load config
# ---------------------------------------------------------------------------

$configPath = Join-Path $PSScriptRoot "bugle-config.json"
if (-not (Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath`nCopy bugle-config.example.json to bugle-config.json and fill in your values."
    exit 1
}

$config = Get-Content $configPath -Raw | ConvertFrom-Json

$apiKey      = $config.BunnyCDN.StorageApiKey
$storageZone = $config.BunnyCDN.StorageZoneName
$endpoint    = $config.BunnyCDN.StorageEndpoint   # e.g. storage.bunnycdn.com
$cdnBase     = $config.BunnyCDN.CdnBaseUrl         # e.g. https://Troop79.b-cdn.net

if (-not $apiKey -or -not $storageZone -or -not $endpoint -or -not $cdnBase) {
    Write-Error "bugle-config.json is missing required BunnyCDN fields. Check bugle-config.example.json."
    exit 1
}

# ---------------------------------------------------------------------------
# Validate local file
# ---------------------------------------------------------------------------

$ImagePath = Resolve-Path $ImagePath -ErrorAction Stop | Select-Object -ExpandProperty Path

if (-not (Test-Path $ImagePath)) {
    Write-Error "Image file not found: $ImagePath"
    exit 1
}

$localFile = Get-Item $ImagePath

# ---------------------------------------------------------------------------
# Determine remote filename
# ---------------------------------------------------------------------------

if ($RemoteName) {
    $remoteFile = $RemoteName
} else {
    $remoteFile = $localFile.Name
}

# Sanitize: replace spaces with hyphens, remove unsafe characters
$remoteFile = $remoteFile -replace '\s+', '-'
$remoteFile = $remoteFile -replace '[^\w\.\-]', ''

if (-not $RemoteName -and $remoteFile -ne $localFile.Name) {
    Write-Host "Filename sanitized: $($localFile.Name) -> $remoteFile" -ForegroundColor Yellow
}

# Default alt text to filename without extension
if (-not $AltText) {
    $AltText = [System.IO.Path]::GetFileNameWithoutExtension($remoteFile) -replace '[-_]', ' '
}

# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------

$uploadUrl = "https://$endpoint/$storageZone/$remoteFile"
$cdnUrl    = "$($cdnBase.TrimEnd('/'))/$remoteFile"

Write-Host ""
Write-Host "Uploading: $($localFile.Name)" -ForegroundColor Cyan
Write-Host "       To: $uploadUrl" -ForegroundColor Cyan

$bytes = [System.IO.File]::ReadAllBytes($ImagePath)

$headers = @{
    "AccessKey"    = $apiKey
    "Content-Type" = "application/octet-stream"
}

try {
    $response = Invoke-WebRequest `
        -Uri $uploadUrl `
        -Method PUT `
        -Headers $headers `
        -Body $bytes `
        -UseBasicParsing

    if ($response.StatusCode -eq 201) {
        Write-Host "   Upload successful (HTTP 201)" -ForegroundColor Green
    } else {
        Write-Host "   Upload returned HTTP $($response.StatusCode)" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Upload failed: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Build markdown and copy to clipboard
# ---------------------------------------------------------------------------

if ($CssClass) {
    $markdown = "![$AltText]($cdnUrl){.$CssClass}"
} else {
    $markdown = "![$AltText]($cdnUrl)"
}

Set-Clipboard -Value $markdown

Write-Host ""
Write-Host "Copied to clipboard:" -ForegroundColor Green
Write-Host "  $markdown" -ForegroundColor White
Write-Host ""
Write-Host "CDN URL only: $cdnUrl" -ForegroundColor DarkGray
Write-Host ""
