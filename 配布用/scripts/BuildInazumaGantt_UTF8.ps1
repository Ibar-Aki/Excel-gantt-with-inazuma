$errorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = $scriptDir
if (-not (Test-Path (Join-Path $projectDir "vba"))) {
    $parentDir = Split-Path -Parent $projectDir
    if (Test-Path (Join-Path $parentDir "vba")) {
        $projectDir = $parentDir
    }
}

$vbaDir = Join-Path $projectDir "vba"
$outputDir = Join-Path $projectDir "output"
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputFile = Join-Path $outputDir "InazumaGantt_v3_$timestamp.xlsm"

# UTF8編集内容をSJIS import対象へ同期
$fixEncodingScript = Join-Path $scriptDir "FixEncoding.ps1"
if (-not (Test-Path $fixEncodingScript)) {
    $fixEncodingScript = Join-Path $projectDir "FixEncoding.ps1"
}
if (Test-Path $fixEncodingScript) {
    Write-Host "Synchronizing SJIS modules..."
    & $fixEncodingScript
}
else {
    throw "FixEncoding.ps1 not found: $fixEncodingScript"
}

# 出力ディレクトリ作成
if (!(Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# 既存ファイル削除
if (Test-Path $outputFile) { Remove-Item $outputFile -Force }

# Excel起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Creating new workbook..."
    $wb = $excel.Workbooks.Add()
    # インポートするファイルリスト（必須モジュール）
    $coreModules = @(
        "InazumaGantt_v3_SJIS.bas",
        "WBSParentRollup_SJIS.bas",
        "WBSRoadmapReport_SJIS.bas",
        "WBSSampleShowcase_SJIS.bas",
        "HierarchyColor_SJIS.bas",
        "SetupWizard_SJIS.bas"
    )

    # モジュールのインポート
    foreach ($file in $coreModules) {
        $path = Join-Path $vbaDir $file
        if (Test-Path $path) {
            Write-Host "Importing $file..."
            $wb.VBProject.VBComponents.Import($path)
        }
        else {
            Write-Warning "File not found: $path"
        }
    }
    
    # 自動セットアップテスト
    Write-Host "Running SilentSetup..."
    try {
        $excel.Run("SilentSetup", $true)
        Write-Host "SilentSetup completed successfully."
    }
    catch {
        Write-Warning "Failed to run SilentSetup: $($_.Exception.Message)"
    }

    $mainSheet = $wb.Worksheets.Item("InazumaGantt_v3")

    # シートモジュールのコード注入
    $sheetModPath = Join-Path $vbaDir "SheetModule_SJIS.bas"
    if (Test-Path $sheetModPath) {
        Write-Host "Injecting SheetModule code..."
        $code = Get-Content $sheetModPath -Encoding Default -Raw
        $code = $code -replace "Attribute VB_Name = .*`r?`n", ""
        $mainSheetCode = $wb.VBProject.VBComponents.Item($mainSheet.CodeName).CodeModule
        $mainSheetCode.AddFromString($code)
    }

    Write-Host "Saving to $outputFile..."
    $wb.SaveAs($outputFile, 52) # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $wb.Close($false)
    $wb = $null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel -ErrorAction SilentlyContinue

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($outputFile)

    $keepSheets = @("InazumaGantt_v3", "InazumaGantt_説明", "設定マスタ", "WBSサマリ")
    $deleteSheets = @()
    for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
        $candidate = $wb.Worksheets.Item($i)
        if ($keepSheets -notcontains $candidate.Name) {
            $deleteSheets += $candidate.Name
        }
    }

    if ($deleteSheets.Count -gt 0) {
        $wb.Worksheets.Item("InazumaGantt_v3").Activate() | Out-Null
        foreach ($sheetName in $deleteSheets) {
            $wb.Worksheets.Item($sheetName).Delete()
        }
        $wb.Save()
    }

    Write-Host "Build Complete!"
}
catch {
    Write-Error "Error occurred: $_"
}
finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Variable excel -ErrorAction SilentlyContinue
    }
}
