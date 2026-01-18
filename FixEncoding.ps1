# FixEncoding.ps1
# VBAファイルをUTF-8からShift-JISに確実に変換するスクリプト

$vbaDir = Join-Path $PSScriptRoot "vba"
$coreModules = @(
    "InazumaGantt_v3",
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

# エンコーディング定義
$utf8 = [System.Text.Encoding]::UTF8 # 標準のUTF-8 (BOMあり/なし両対応)
$sjis = [System.Text.Encoding]::GetEncoding(932) # Shift-JIS (CP932)

function Convert-ToSjis {
    param(
        [string]$SourcePath,
        [string]$DestPath
    )
    
    try {
        # 既存のみ削除
        if (Test-Path $DestPath) {
            Remove-Item $DestPath -Force
        }
        
        # 読み込み (UTF-8として)
        $content = [System.IO.File]::ReadAllText($SourcePath, $utf8)
        
        # 書き込み (Shift-JISとして)
        [System.IO.File]::WriteAllText($DestPath, $content, $sjis)
        
        Write-Host "Converted: $(Split-Path $SourcePath -Leaf) -> $(Split-Path $DestPath -Leaf)"
    }
    catch {
        Write-Error "Failed to convert $SourcePath : $_"
    }
}

# メイン処理
foreach ($mod in $coreModules) {
    $utf8Path = Join-Path $vbaDir "${mod}_UTF8.bas"
    $sjisPath = Join-Path $vbaDir "${mod}_SJIS.bas"

    if (Test-Path $utf8Path) {
        Convert-ToSjis -SourcePath $utf8Path -DestPath $sjisPath
    } else {
        Write-Warning "Source file not found: $utf8Path"
    }
}

foreach ($mod in $addonModules) {
    $utf8Path = Join-Path $vbaDir "${mod}_UTF8.bas"
    $sjisPath = Join-Path $vbaDir "${mod}_SJIS.bas"

    if (Test-Path $utf8Path) {
        Convert-ToSjis -SourcePath $utf8Path -DestPath $sjisPath
    } else {
        Write-Warning "Source file not found: $utf8Path"
    }
}

Write-Host "Encoding Fix Process Completed."
