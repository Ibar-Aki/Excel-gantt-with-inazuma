$errorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbaDir = Join-Path $scriptDir "vba"
$outputDir = Join-Path $scriptDir "output"
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputFile = Join-Path $outputDir "InazumaGantt_v3_$timestamp.xlsm"

# エンコーディング修正スクリプトの実行（スキップ）
# $fixEncodingScript = Join-Path $scriptDir "FixEncoding.ps1"
# if (Test-Path $fixEncodingScript) { ... }

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

    $defaultSheet = $wb.Worksheets.Item(1)
    if ($defaultSheet.Name -ne "InazumaGantt_v3" -and $wb.Worksheets.Count -gt 1) {
        if ($excel.WorksheetFunction.CountA($defaultSheet.UsedRange) -eq 0) {
            $defaultSheet.Delete()
        }
    }

    # 保存
    Write-Host "Saving to $outputFile..."
    $wb.SaveAs($outputFile, 52) # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    
    Write-Host "Build Complete!"
}
catch {
    Write-Error "Error occurred: $_"
}
finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    Remove-Variable excel
}
