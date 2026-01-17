# BuildInazumaGantt.ps1
# InazumaGantt v2 + データ移管ウィザード フルビルドスクリプト

$errorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vbaDir = Join-Path $scriptDir "vba"
$outputDir = Join-Path $scriptDir "output"
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$outputFile = Join-Path $outputDir "InazumaGantt_v2_UTF8_$timestamp.xlsm"

# エンコーディング修正スクリプトの実行（スキップ）
# $fixEncodingScript = Join-Path $scriptDir "FixEncoding.ps1"
# if (Test-Path $fixEncodingScript) { ... }

# 出力ディレクトリ作成
if (!(Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# 既存ファイル削除
if (Test-Path $outputFile) { Remove-Item $outputFile -Force }

# Excel起動
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true # デバッグ用に表示（必要に応じて$false）
$excel.DisplayAlerts = $false

try {
    Write-Host "Creating new workbook..."
    $wb = $excel.Workbooks.Add()
    
    # シート名変更 (Sheet1 -> InazumaGantt_v2)
    $mainSheet = $wb.Worksheets.Item(1)
    $mainSheet.Name = "InazumaGantt_v2"
    
    # インポートするファイルリスト（必須モジュール）
    $coreModules = @(
        "InazumaGantt_v2_UTF8.bas",
        "HierarchyColor_UTF8.bas",
        "SetupWizard_UTF8.bas"
    )

    # インポートするファイルリスト（データ移管アドオン）
    $addonModules = @(
        "addons\DataMigration\DataMigration_UTF8.bas",
        "addons\DataMigration\WBSParser_UTF8.bas",
        "addons\DataMigration\DataMigrationWizard_UTF8.bas",
        "addons\DataMigration\MigrationFormBuilder_UTF8.bas"
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

    foreach ($file in $addonModules) {
        $path = Join-Path $vbaDir $file
        if (Test-Path $path) {
            Write-Host "Importing $file..."
            $wb.VBProject.VBComponents.Import($path)
        }
        else {
            Write-Warning "File not found: $path"
        }
    }
    
    # シートモジュールのコード注入
    $sheetModPath = Join-Path $vbaDir "SheetModule_UTF8.bas"
    if (Test-Path $sheetModPath) {
        Write-Host "Injecting SheetModule code..."
        $code = Get-Content $sheetModPath -Encoding Default -Raw
        $code = $code -replace "Attribute VB_Name = .*`r?`n", ""
        $mainSheetCode = $wb.VBProject.VBComponents.Item($mainSheet.CodeName).CodeModule
        $mainSheetCode.AddFromString($code)
    }
    
    # UserFormの生成 (MigrationFormBuilderを実行)
    Write-Host "Generating UserForm..."
    try {
        $excel.Run("CreateMigrationWizardForm")
    }
    catch {
        Write-Warning "Failed to run CreateMigrationWizardForm: $($_.Exception.Message)"
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
