$ErrorActionPreference = "Stop"

function Get-AvailablePowerShellPath {
    # Prefer PowerShell 7 so UTF-8 build scripts keep smoke-test literals intact.
    foreach ($candidate in @("pwsh.exe", "powershell.exe")) {
        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($null -ne $command) {
            return $command.Source
        }
    }

    try {
        $currentProcess = Get-Process -Id $PID -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($currentProcess.Path)) {
            return $currentProcess.Path
        }
    }
    catch {
    }

    throw "PowerShell executable not found."
}

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
$powerShellPath = Get-AvailablePowerShellPath

if (-not (Test-Path $buildScript)) {
    throw "Build script not found: $buildScript"
}

if (Test-Path $watchdogScript) {
    & $watchdogScript `
        -ActionName "One-click Inazuma build" `
        -FilePath $powerShellPath `
        -ArgumentList @("-File", $buildScript) `
        -WorkingDirectory $buildWorkingDir `
        -TimeoutSeconds 60 `
        -CaptureIntervalSeconds 15 `
        -MaxCaptures 4
}
else {
    & $powerShellPath -File $buildScript
}

$latestFile = Get-ChildItem -Path $outputDir -Filter "InazumaGantt_v3_*.xlsm" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $latestFile) {
    throw "No workbook was generated in $outputDir"
}

Write-Host "Generated workbook: $($latestFile.FullName)"
