# FixEncoding.ps1
# VBAファイルをUTF-8からShift-JISに正しく変換して再生成するスクリプト

$vbaDir = Join-Path $PSScriptRoot "vba"
$coreModules = @(
    "InazumaGantt_v2",
    "HierarchyColor",
    "SetupWizard",
    "SheetModule"
)

$addonModules = @(
    "addons\DataMigration\DataMigration",
    "addons\DataMigration\WBSParser",
    "addons\DataMigration\DataMigrationWizard",
    "addons\DataMigration\MigrationFormBuilder"
)

foreach ($mod in $coreModules) {
    $utf8Path = Join-Path $vbaDir "${mod}_UTF8.bas"
    $sjisPath = Join-Path $vbaDir "${mod}_SJIS.bas"

    if (Test-Path $utf8Path) {
        Write-Host "Converting ${mod}..."
        Get-Content $utf8Path -Encoding UTF8 | Set-Content $sjisPath -Encoding Default
    } else {
        Write-Warning "Source file not found: $utf8Path"
    }
}

foreach ($mod in $addonModules) {
    $utf8Path = Join-Path $vbaDir "${mod}_UTF8.bas"
    $sjisPath = Join-Path $vbaDir "${mod}_SJIS.bas"

    if (Test-Path $utf8Path) {
        Write-Host "Converting ${mod}..."
        Get-Content $utf8Path -Encoding UTF8 | Set-Content $sjisPath -Encoding Default
    } else {
        Write-Warning "Source file not found: $utf8Path"
    }
}

Write-Host "Encoding Fix Complete!"
