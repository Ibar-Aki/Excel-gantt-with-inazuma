param(
    [string]$ConverterRoot
)

$ErrorActionPreference = "Stop"

function Resolve-ConverterRoot {
    param(
        [string]$PreferredRoot,
        [string]$BaseDir
    )

    $candidates = @()

    if (-not [string]::IsNullOrWhiteSpace($PreferredRoot)) {
        $candidates += $PreferredRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($env:INAZUMA_BUNDLE_CONVERTER_ROOT)) {
        $candidates += $env:INAZUMA_BUNDLE_CONVERTER_ROOT
    }

    $candidates += (Join-Path $BaseDir "ps1,batの変換器")
    $candidates += (Join-Path (Split-Path -Parent $BaseDir) "ps1,batの変換器")

    foreach ($candidate in ($candidates | Select-Object -Unique)) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
    }

    throw "Converter root not found. Specify -ConverterRoot or set INAZUMA_BUNDLE_CONVERTER_ROOT."
}

function Get-AvailablePowerShellPath {
    try {
        $currentProcess = Get-Process -Id $PID -ErrorAction Stop
        if (-not [string]::IsNullOrWhiteSpace($currentProcess.Path)) {
            return $currentProcess.Path
        }
    }
    catch {
    }

    foreach ($candidate in @("powershell.exe", "pwsh.exe")) {
        $command = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($null -ne $command) {
            return $command.Source
        }
    }

    throw "PowerShell executable not found."
}

function Invoke-BuildScriptWithWatchdog {
    param(
        [Parameter(Mandatory)]
        [string]$BuildScriptPath,
        [Parameter(Mandatory)]
        [string]$WorkingDirectory,
        [Parameter(Mandatory)]
        [string]$ActionName
    )

    $watchdogScript = Join-Path $env:USERPROFILE ".codex\tools\invoke_with_desktop_watchdog.ps1"
    $powerShellPath = Get-AvailablePowerShellPath

    if (Test-Path -LiteralPath $watchdogScript) {
        & $watchdogScript `
            -ActionName $ActionName `
            -FilePath $powerShellPath `
            -ArgumentList @("-File", $BuildScriptPath) `
            -WorkingDirectory $WorkingDirectory `
            -TimeoutSeconds 120 `
            -CaptureIntervalSeconds 15 `
            -MaxCaptures 8
    }
    else {
        & $powerShellPath -File $BuildScriptPath
    }
}

function Move-PathWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,
        [Parameter(Mandatory)]
        [string]$DestinationPath,
        [int]$RetryCount = 5,
        [int]$DelayMilliseconds = 750
    )

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Move-Item -LiteralPath $SourcePath -Destination $DestinationPath -Force
            return
        }
        catch {
            if ($attempt -eq $RetryCount) {
                throw "Failed to move '$SourcePath' to '$DestinationPath'. Close Excel or Explorer windows using this path and retry. $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Remove-PathWithRetry {
    param(
        [Parameter(Mandatory)]
        [string]$TargetPath,
        [int]$RetryCount = 5,
        [int]$DelayMilliseconds = 750
    )

    if (-not (Test-Path -LiteralPath $TargetPath)) {
        return
    }

    for ($attempt = 1; $attempt -le $RetryCount; $attempt++) {
        try {
            Remove-Item -LiteralPath $TargetPath -Recurse -Force
            return
        }
        catch {
            if ($attempt -eq $RetryCount) {
                throw "Failed to remove '$TargetPath'. Close Excel or Explorer windows using this path and retry. $($_.Exception.Message)"
            }
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
}

function Sync-DirectoryWithRobocopy {
    param(
        [Parameter(Mandatory)]
        [string]$SourcePath,
        [Parameter(Mandatory)]
        [string]$DestinationPath
    )

    $robocopyLog = Join-Path ([System.IO.Path]::GetTempPath()) ("robocopy_inazuma_" + [guid]::NewGuid().ToString("N") + ".log")
    try {
        $arguments = @(
            $SourcePath,
            $DestinationPath,
            "/MIR",
            "/R:2",
            "/W:1",
            "/NFL",
            "/NDL",
            "/NJH",
            "/NJS",
            "/NP",
            "/LOG:$robocopyLog"
        )

        & robocopy @arguments | Out-Null
        $exitCode = $LASTEXITCODE
        if ($exitCode -gt 7) {
            $logText = ""
            if (Test-Path -LiteralPath $robocopyLog) {
                $logText = Get-Content -LiteralPath $robocopyLog -Raw
            }
            throw "Robocopy mirror failed with exit code $exitCode. $logText"
        }
    }
    finally {
        if (Test-Path -LiteralPath $robocopyLog) {
            Remove-Item -LiteralPath $robocopyLog -Force -ErrorAction SilentlyContinue
        }
    }
}

function Replace-DirectoryFromStaging {
    param(
        [Parameter(Mandatory)]
        [string]$StagingPath,
        [Parameter(Mandatory)]
        [string]$TargetPath
    )

    $targetParent = Split-Path -Parent $TargetPath
    $targetLeaf = Split-Path -Leaf $TargetPath
    $backupPath = Join-Path $targetParent ($targetLeaf + "_previous_" + [guid]::NewGuid().ToString("N"))

    if (Test-Path -LiteralPath $TargetPath) {
        try {
            Move-PathWithRetry -SourcePath $TargetPath -DestinationPath $backupPath
        }
        catch {
            Write-Warning $_
            Sync-DirectoryWithRobocopy -SourcePath $StagingPath -DestinationPath $TargetPath
            Remove-PathWithRetry -TargetPath $StagingPath
            return
        }
    }

    try {
        Move-PathWithRetry -SourcePath $StagingPath -DestinationPath $TargetPath
    }
    catch {
        if ((Test-Path -LiteralPath $backupPath) -and -not (Test-Path -LiteralPath $TargetPath)) {
            Move-PathWithRetry -SourcePath $backupPath -DestinationPath $TargetPath
        }
        throw
    }

    if (Test-Path -LiteralPath $backupPath) {
        try {
            Remove-PathWithRetry -TargetPath $backupPath
        }
        catch {
            Write-Warning $_
        }
    }
}

function Ensure-DirectoryExists {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path $scriptDir "output"
$distributionDir = Join-Path $scriptDir "distribution_v3"
$distributionStagingDir = Join-Path $scriptDir ("distribution_v3_staging_" + [guid]::NewGuid().ToString("N"))
$scriptsDir = Join-Path $distributionStagingDir "scripts"
$excelDir = Join-Path $distributionStagingDir "excel"
$vbaDir = Join-Path $distributionStagingDir "vba"
$docsDir = Join-Path $distributionStagingDir "docs"
$bundleTextPath = Join-Path $distributionStagingDir "PackageContents.md"
$converterRoot = Resolve-ConverterRoot -PreferredRoot $ConverterRoot -BaseDir $scriptDir
$converterInputDir = Join-Path $converterRoot "input_files"
$converterOutputDir = Join-Path $converterRoot "output_bundle"
$bundleFolderName = "InazumaGantt_v3_Distribution"
$workbookPayloadPath = Join-Path $excelDir "WorkbookPayload.json"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$activeVbaFiles = @(
    "InazumaGantt_v3_UTF8.bas",
    "InazumaGantt_v3_SJIS.bas",
    "WBSParentRollup_UTF8.bas",
    "WBSParentRollup_SJIS.bas",
    "WBSRoadmapReport_UTF8.bas",
    "WBSRoadmapReport_SJIS.bas",
    "WBSSampleShowcase_UTF8.bas",
    "WBSSampleShowcase_SJIS.bas",
    "HierarchyColor_UTF8.bas",
    "HierarchyColor_SJIS.bas",
    "SetupWizard_UTF8.bas",
    "SetupWizard_SJIS.bas",
    "SheetModule_UTF8.bas",
    "SheetModule_SJIS.bas"
)

$scriptFiles = @(
    "BuildInazumaGantt_UTF8.ps1",
    "FixEncoding.ps1",
    "OneClick_CreateLatestWorkbook.ps1",
    "RestoreWorkbookFromPayload.ps1",
    "Run_OneClick_CreateLatestWorkbook.bat",
    "Run_RestoreWorkbookFromPayload.bat",
    "CreateDistributionPackage.ps1",
    "Run_CreateDistributionPackage.bat"
)

$docFiles = @(
    "RestoreGuide.md",
    "高速入力_状況_進捗率仕様レポート.md"
)

$fixEncodingScript = Join-Path $scriptDir "FixEncoding.ps1"
if (-not (Test-Path -LiteralPath $fixEncodingScript)) {
    throw "FixEncoding.ps1 not found: $fixEncodingScript"
}

Write-Host "Synchronizing SJIS modules for distribution..."
& $fixEncodingScript

Invoke-BuildScriptWithWatchdog `
    -BuildScriptPath (Join-Path $scriptDir "BuildInazumaGantt_UTF8.ps1") `
    -WorkingDirectory $scriptDir `
    -ActionName "Create v3 distribution workbook"

$latestFile = Get-ChildItem -Path $outputDir -Filter "InazumaGantt_v3_*.xlsm" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $latestFile) {
    throw "No workbook available for distribution."
}

if (Test-Path -LiteralPath $distributionStagingDir) {
    Remove-PathWithRetry -TargetPath $distributionStagingDir
}
New-Item -ItemType Directory -Path $distributionStagingDir | Out-Null
New-Item -ItemType Directory -Path $scriptsDir | Out-Null
New-Item -ItemType Directory -Path $excelDir | Out-Null
New-Item -ItemType Directory -Path $vbaDir | Out-Null
New-Item -ItemType Directory -Path $docsDir | Out-Null

Copy-Item -LiteralPath $latestFile.FullName -Destination (Join-Path $excelDir $latestFile.Name) -Force

$workbookPayload = [ordered]@{
    fileName = $latestFile.Name
    sourcePath = (".\excel\" + $latestFile.Name)
    byteLength = $latestFile.Length
    sha256 = (Get-FileHash -LiteralPath $latestFile.FullName -Algorithm SHA256).Hash
    generatedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss K")
    contentBase64 = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($latestFile.FullName))
}

$payloadJson = ($workbookPayload | ConvertTo-Json -Depth 4) -replace "`r`n", "`n"
[System.IO.File]::WriteAllText($workbookPayloadPath, $payloadJson + "`n", $utf8NoBom)

foreach ($fileName in $activeVbaFiles) {
    Copy-Item -LiteralPath (Join-Path (Join-Path $scriptDir "vba") $fileName) -Destination (Join-Path $vbaDir $fileName) -Force
}

foreach ($fileName in $scriptFiles) {
    Copy-Item -LiteralPath (Join-Path $scriptDir $fileName) -Destination (Join-Path $scriptsDir $fileName) -Force
}

foreach ($fileName in $docFiles) {
    Copy-Item -LiteralPath (Join-Path (Join-Path $scriptDir "docs") $fileName) -Destination (Join-Path $docsDir $fileName) -Force
}

$vbaModuleLines = $activeVbaFiles | ForEach-Object { "vba\" + $_ }
$docLines = $docFiles | ForEach-Object { "docs\" + $_ }

$bundleLines = @(
    "# InazumaGantt v3 distribution package",
    "",
    ("Created at: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") + " JST"),
    "Author: Codex (GPT-5)",
    "",
    ("Generated at: " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss K")),
    ("Package root: ."),
    ("Converter root: set at runtime"),
    "",
    "[Workbook]",
    ("excel\" + $latestFile.Name),
    "",
    "[One-click launchers]",
    "scripts\Run_OneClick_CreateLatestWorkbook.bat",
    "scripts\OneClick_CreateLatestWorkbook.ps1",
    "scripts\Run_RestoreWorkbookFromPayload.bat",
    "scripts\RestoreWorkbookFromPayload.ps1",
    "scripts\Run_CreateDistributionPackage.bat",
    "scripts\CreateDistributionPackage.ps1",
    "",
    "[Build helpers]",
    "scripts\BuildInazumaGantt_UTF8.ps1",
    "scripts\FixEncoding.ps1",
    "",
    "[VBA modules]"
) + $vbaModuleLines + @(
    "",
    "[Workbook payload]",
    "excel\WorkbookPayload.json",
    "",
    "[Documents]"
) + $docLines + @(
    "",
    "[VBA note]",
    "Use *_UTF8.bas files for editing and manual copy/paste into the VBA editor.",
    "Refresh *_SJIS.bas with scripts\\FixEncoding.ps1, then use them only for Excel VBA import on Windows.",
    "",
    "[Usage]",
    "1. Double-click scripts\Run_OneClick_CreateLatestWorkbook.bat to generate the latest workbook.",
    ("2. Open excel\" + $latestFile.Name + " to review the generated sample workbook."),
    "3. If you restore from the converter bundle, run scripts\Run_RestoreWorkbookFromPayload.bat to recreate the xlsm from WorkbookPayload.json.",
    "4. Use scripts\Run_CreateDistributionPackage.bat to rebuild this distribution folder."
)

$bundleText = [string]::Join("`n", $bundleLines)
[System.IO.File]::WriteAllText($bundleTextPath, $bundleText + "`n", $utf8NoBom)

if (-not (Test-Path -LiteralPath $converterRoot)) {
    throw "Converter root not found: $converterRoot"
}

Ensure-DirectoryExists -Path $converterInputDir
Ensure-DirectoryExists -Path $converterOutputDir

$backupDir = Join-Path ([System.IO.Path]::GetTempPath()) ("inazuma_converter_backup_" + [guid]::NewGuid().ToString("N"))
$stagingDir = Join-Path $converterInputDir $bundleFolderName
$bundleOutputFile = $null

try {
    New-Item -ItemType Directory -Path $backupDir | Out-Null

    Get-ChildItem -LiteralPath $converterInputDir -Force | ForEach-Object {
        Move-PathWithRetry -SourcePath $_.FullName -DestinationPath (Join-Path $backupDir $_.Name)
    }

    Copy-Item -LiteralPath $distributionStagingDir -Destination $stagingDir -Recurse -Force

    $beforeBundlePaths = @(Get-ChildItem -LiteralPath $converterOutputDir -Filter "bundle_*.txt" -File | Select-Object -ExpandProperty FullName)

    Push-Location -LiteralPath $converterRoot
    try {
        & (Join-Path $converterRoot "bundle_system.ps1") -Mode Bundle
    }
    finally {
        Pop-Location
    }

    $bundleOutputFile = Get-ChildItem -LiteralPath $converterOutputDir -Filter "bundle_*.txt" -File |
        Where-Object { $beforeBundlePaths -notcontains $_.FullName } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -eq $bundleOutputFile) {
        $bundleOutputFile = Get-ChildItem -LiteralPath $converterOutputDir -Filter "bundle_*.txt" -File |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
    }

    if ($null -eq $bundleOutputFile) {
        throw "No converter bundle file was created."
    }

    Get-ChildItem -LiteralPath $distributionStagingDir -Filter "bundle_*.txt" -File -ErrorAction SilentlyContinue | Remove-Item -Force
    Get-ChildItem -LiteralPath $distributionStagingDir -Filter "配布内容.txt" -File -ErrorAction SilentlyContinue | Remove-Item -Force
    Copy-Item -LiteralPath $bundleOutputFile.FullName -Destination (Join-Path $distributionStagingDir $bundleOutputFile.Name) -Force
}
finally {
    if (Test-Path -LiteralPath $stagingDir) {
        Remove-PathWithRetry -TargetPath $stagingDir
    }

    Get-ChildItem -LiteralPath $backupDir -Force -ErrorAction SilentlyContinue | ForEach-Object {
        Move-PathWithRetry -SourcePath $_.FullName -DestinationPath (Join-Path $converterInputDir $_.Name)
    }

    if (Test-Path -LiteralPath $backupDir) {
        Remove-PathWithRetry -TargetPath $backupDir
    }
}

Replace-DirectoryFromStaging -StagingPath $distributionStagingDir -TargetPath $distributionDir

Write-Host "Distribution package created: $distributionDir"
Write-Host "Workbook bundled: $($latestFile.Name)"
if ($null -ne $bundleOutputFile) {
    Write-Host "Converter bundle created: $($bundleOutputFile.Name)"
}
