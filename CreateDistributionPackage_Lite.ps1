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

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputDir = Join-Path $scriptDir "output"
$distributionDir = Join-Path $scriptDir "distribution_lite"
$scriptsDir = Join-Path $distributionDir "scripts"
$excelDir = Join-Path $distributionDir "excel"
$vbaDir = Join-Path $distributionDir "vba"
$docsDir = Join-Path $distributionDir "docs"
$bundleTextPath = Join-Path $distributionDir "PackageContents.md"
$converterRoot = Resolve-ConverterRoot -PreferredRoot $ConverterRoot -BaseDir $scriptDir
$converterInputDir = Join-Path $converterRoot "input_files"
$converterOutputDir = Join-Path $converterRoot "output_bundle"
$bundleFolderName = "InazumaGantt_Lite_Distribution"
$workbookPayloadPath = Join-Path $excelDir "WorkbookPayload.json"
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

$activeVbaFiles = @(
    "InazumaGantt_Lite_UTF8.bas",
    "InazumaGantt_Lite_SJIS.bas",
    "WBSParentRollup_Lite_UTF8.bas",
    "WBSParentRollup_Lite_SJIS.bas",
    "WBSRoadmapReport_Lite_UTF8.bas",
    "WBSRoadmapReport_Lite_SJIS.bas",
    "WBSSampleShowcase_Lite_UTF8.bas",
    "WBSSampleShowcase_Lite_SJIS.bas",
    "HierarchyColor_Lite_UTF8.bas",
    "HierarchyColor_Lite_SJIS.bas",
    "SetupWizard_Lite_UTF8.bas",
    "SetupWizard_Lite_SJIS.bas",
    "SheetModule_Lite_UTF8.bas",
    "SheetModule_Lite_SJIS.bas"
)

$scriptFiles = @(
    "BuildInazumaGantt_Lite_UTF8.ps1",
    "FixEncoding.ps1",
    "OneClick_CreateLiteWorkbook.ps1",
    "RestoreWorkbookFromPayload.ps1",
    "Run_OneClick_CreateLiteWorkbook.bat",
    "Run_RestoreWorkbookFromPayload.bat",
    "CreateDistributionPackage_Lite.ps1",
    "Run_CreateDistributionPackage_Lite.bat"
)

$docFiles = @(
    "RestoreGuide_Lite.md"
)

& (Join-Path $scriptDir "BuildInazumaGantt_Lite_UTF8.ps1")

$latestFile = Get-ChildItem -Path $outputDir -Filter "InazumaGantt_Lite_*.xlsm" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $latestFile) {
    throw "No workbook available for distribution."
}

if (Test-Path $distributionDir) {
    Get-ChildItem -Path $distributionDir -Force | Remove-Item -Recurse -Force
}
else {
    New-Item -ItemType Directory -Path $distributionDir | Out-Null
}

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

$payloadJson = $workbookPayload | ConvertTo-Json -Depth 4
[System.IO.File]::WriteAllText($workbookPayloadPath, $payloadJson, $utf8NoBom)

foreach ($fileName in $activeVbaFiles) {
    Copy-Item -LiteralPath (Join-Path (Join-Path $scriptDir "vba") $fileName) -Destination (Join-Path $vbaDir $fileName) -Force
}

foreach ($fileName in $scriptFiles) {
    Copy-Item -LiteralPath (Join-Path $scriptDir $fileName) -Destination (Join-Path $scriptsDir $fileName) -Force
}

foreach ($fileName in $docFiles) {
    Copy-Item -LiteralPath (Join-Path (Join-Path $scriptDir "docs") $fileName) -Destination (Join-Path $docsDir $fileName) -Force
}

$bundleLines = @(
    "# InazumaGantt Lite distribution package",
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
    "scripts\Run_OneClick_CreateLiteWorkbook.bat",
    "scripts\OneClick_CreateLiteWorkbook.ps1",
    "scripts\Run_RestoreWorkbookFromPayload.bat",
    "scripts\RestoreWorkbookFromPayload.ps1",
    "scripts\Run_CreateDistributionPackage_Lite.bat",
    "scripts\CreateDistributionPackage_Lite.ps1",
    "",
    "[Build helpers]",
    "scripts\BuildInazumaGantt_Lite_UTF8.ps1",
    "scripts\FixEncoding.ps1",
    "",
    "[VBA modules]",
    ($activeVbaFiles | ForEach-Object { "vba\" + $_ }),
    "",
    "[Workbook payload]",
    "excel\WorkbookPayload.json",
    "",
    "[Documents]",
    "docs\RestoreGuide_Lite.md",
    "",
    "[VBA note]",
    "Read *_UTF8.bas files when you inspect source text.",
    "Use *_SJIS.bas files only for Excel VBA import on Windows.",
    "",
    "[Usage]",
    "1. Double-click scripts\Run_OneClick_CreateLiteWorkbook.bat to generate the latest workbook.",
    ("2. Open excel\" + $latestFile.Name + " to review the generated sample workbook."),
    "3. If you restore from the converter bundle, run scripts\Run_RestoreWorkbookFromPayload.bat to recreate the xlsm from WorkbookPayload.json.",
    "4. Use scripts\Run_CreateDistributionPackage_Lite.bat to rebuild this distribution folder."
)

[System.IO.File]::WriteAllLines($bundleTextPath, $bundleLines, $utf8NoBom)

if (-not (Test-Path -LiteralPath $converterRoot)) {
    throw "Converter root not found: $converterRoot"
}

$backupDir = Join-Path ([System.IO.Path]::GetTempPath()) ("inazuma_converter_backup_" + [guid]::NewGuid().ToString("N"))
$stagingDir = Join-Path $converterInputDir $bundleFolderName
$bundleOutputFile = $null

try {
    New-Item -ItemType Directory -Path $backupDir | Out-Null

    Get-ChildItem -LiteralPath $converterInputDir -Force | ForEach-Object {
        Move-Item -LiteralPath $_.FullName -Destination $backupDir
    }

    Copy-Item -LiteralPath $distributionDir -Destination $stagingDir -Recurse -Force

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

    Get-ChildItem -LiteralPath $distributionDir -Filter "bundle_*.txt" -File -ErrorAction SilentlyContinue | Remove-Item -Force
    Get-ChildItem -LiteralPath $distributionDir -Filter "配布内容.txt" -File -ErrorAction SilentlyContinue | Remove-Item -Force
    Copy-Item -LiteralPath $bundleOutputFile.FullName -Destination (Join-Path $distributionDir $bundleOutputFile.Name) -Force
}
finally {
    if (Test-Path -LiteralPath $stagingDir) {
        Remove-Item -LiteralPath $stagingDir -Recurse -Force
    }

    Get-ChildItem -LiteralPath $backupDir -Force -ErrorAction SilentlyContinue | ForEach-Object {
        Move-Item -LiteralPath $_.FullName -Destination $converterInputDir
    }

    if (Test-Path -LiteralPath $backupDir) {
        Remove-Item -LiteralPath $backupDir -Recurse -Force
    }
}

Write-Host "Distribution package created: $distributionDir"
Write-Host "Workbook bundled: $($latestFile.Name)"
if ($null -ne $bundleOutputFile) {
    Write-Host "Converter bundle created: $($bundleOutputFile.Name)"
}
