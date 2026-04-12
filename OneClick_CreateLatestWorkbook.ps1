$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = $scriptDir
if (-not (Test-Path (Join-Path $projectDir "vba"))) {
    $parentDir = Split-Path -Parent $projectDir
    if (Test-Path (Join-Path $parentDir "vba")) {
        $projectDir = $parentDir
    }
}

$buildScript = Join-Path $scriptDir "BuildInazumaGantt_UTF8.ps1"
$buildWorkingDir = $projectDir
$watchdogScript = Join-Path $env:USERPROFILE ".codex\tools\invoke_with_desktop_watchdog.ps1"
$outputDir = Join-Path $projectDir "output"

if (-not (Test-Path $buildScript)) {
    throw "Build script not found: $buildScript"
}

if (Test-Path $watchdogScript) {
    & $watchdogScript `
        -ActionName "One-click Inazuma build" `
        -FilePath (Get-Command pwsh).Source `
        -ArgumentList @("-File", $buildScript) `
        -WorkingDirectory $buildWorkingDir `
        -TimeoutSeconds 60 `
        -CaptureIntervalSeconds 15 `
        -MaxCaptures 4
}
else {
    & (Get-Command pwsh).Source -File $buildScript
}

$latestFile = Get-ChildItem -Path $outputDir -Filter "InazumaGantt_v3_*.xlsm" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $latestFile) {
    throw "No workbook was generated in $outputDir"
}

Write-Host "Generated workbook: $($latestFile.FullName)"
