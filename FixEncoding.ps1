# FixEncoding.ps1
# VBAファイルをUTF-8からShift-JISに確実に変換するスクリプト

$projectDir = $PSScriptRoot
if (-not (Test-Path (Join-Path $projectDir "vba"))) {
    $parentDir = Split-Path -Parent $projectDir
    if (Test-Path (Join-Path $parentDir "vba")) {
        $projectDir = $parentDir
    }
}

$vbaDir = Join-Path $projectDir "vba"
$coreModules = @(
    "InazumaGantt_v3",
    "InazumaGantt_Lite",
    "WBSParentRollup",
    "WBSParentRollup_Lite",
    "WBSRoadmapReport",
    "WBSRoadmapReport_Lite",
    "WBSSampleShowcase",
    "WBSSampleShowcase_Lite",
    "HierarchyColor",
    "HierarchyColor_Lite",
    "SetupWizard",
    "SetupWizard_Lite",
    "SheetModule",
    "SheetModule_Lite"
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
    }
}

Write-Host "Encoding Fix Process Completed."
