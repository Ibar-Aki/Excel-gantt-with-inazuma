param(
    [string]$PayloadPath,
    [switch]$Force
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = $scriptDir
if (-not (Test-Path (Join-Path $projectDir "excel"))) {
    $parentDir = Split-Path -Parent $projectDir
    if (Test-Path (Join-Path $parentDir "excel")) {
        $projectDir = $parentDir
    }
}

if ([string]::IsNullOrWhiteSpace($PayloadPath)) {
    $PayloadPath = Join-Path $projectDir "excel\WorkbookPayload.json"
}

if (-not (Test-Path -LiteralPath $PayloadPath)) {
    throw "Payload file not found: $PayloadPath"
}

$payload = Get-Content -LiteralPath $PayloadPath -Encoding UTF8 -Raw | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($payload.fileName)) {
    throw "Payload does not include fileName."
}

if ([string]::IsNullOrWhiteSpace($payload.contentBase64)) {
    throw "Payload does not include contentBase64."
}

$targetPath = Join-Path (Split-Path -Parent $PayloadPath) $payload.fileName

if ((Test-Path -LiteralPath $targetPath) -and (-not $Force)) {
    throw "Target workbook already exists: $targetPath"
}

$bytes = [Convert]::FromBase64String([string]$payload.contentBase64)
[System.IO.File]::WriteAllBytes($targetPath, $bytes)

Write-Host "Restored workbook: $targetPath"
