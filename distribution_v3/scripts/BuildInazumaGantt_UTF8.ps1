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

# Sync UTF-8 source modules to SJIS import targets
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

# Ensure output directory exists
if (!(Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }

# Remove target file if it already exists
if (Test-Path $outputFile) { Remove-Item $outputFile -Force }

# Start Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Throw-VbaAccessGuidance([string]$detail) {
    throw ("Excel workbook generation failed during VBA import/injection. " +
           "On another PC, enable 'Trust access to the VBA project object model' in Excel: " +
           "File > Options > Trust Center > Trust Center Settings > Macro Settings. Detail: " + $detail)
}

function Get-WorksheetOrThrow($workbook, [string]$sheetName) {
    try {
        return $workbook.Worksheets.Item($sheetName)
    }
    catch {
        throw "Required worksheet was not created: $sheetName"
    }
}

function Close-WorkbookSafely([ref]$workbookRef, [bool]$saveChanges = $false) {
    if ($null -ne $workbookRef.Value) {
        try {
            $workbookRef.Value.Close($saveChanges)
        }
        catch {
        }
        finally {
            $workbookRef.Value = $null
        }
    }
}

function Close-ExcelSafely([ref]$excelRef) {
    if ($null -ne $excelRef.Value) {
        try {
            $excelRef.Value.Quit()
        }
        catch {
        }
        finally {
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelRef.Value) | Out-Null
            }
            catch {
            }
            $excelRef.Value = $null
        }
    }
}

try {
    Write-Host "Creating new workbook..."
    $wb = $excel.Workbooks.Add()
    # Required standard modules to import
    $coreModules = @(
        "InazumaGantt_v3_SJIS.bas",
        "WBSParentRollup_SJIS.bas",
        "WBSRoadmapReport_SJIS.bas",
        "WBSSampleShowcase_SJIS.bas",
        "HierarchyColor_SJIS.bas",
        "SetupWizard_SJIS.bas"
    )

    # Import standard modules
    foreach ($file in $coreModules) {
        $path = Join-Path $vbaDir $file
        if (Test-Path $path) {
            Write-Host "Importing $file..."
            try {
                $wb.VBProject.VBComponents.Import($path) | Out-Null
            }
            catch {
                Throw-VbaAccessGuidance($_.Exception.Message)
            }
        }
        else {
            Write-Warning "File not found: $path"
        }
    }
    
    # Run setup after module import
    Write-Host "Running SilentSetup..."
    try {
        $excel.Run("SilentSetup", $true)
        Write-Host "SilentSetup completed successfully."
    }
    catch {
        throw "SilentSetup failed: $($_.Exception.Message)"
    }

    $mainSheet = Get-WorksheetOrThrow $wb "InazumaGantt_v3"

    # Inject sheet module code into the main worksheet
    $sheetModPath = Join-Path $vbaDir "SheetModule_SJIS.bas"
    if (Test-Path $sheetModPath) {
        Write-Host "Injecting SheetModule code..."
        $code = Get-Content $sheetModPath -Encoding Default -Raw
        $code = $code -replace "Attribute VB_Name = .*`r?`n", ""
        try {
            $mainSheetCode = $wb.VBProject.VBComponents.Item($mainSheet.CodeName).CodeModule
            $mainSheetCode.AddFromString($code)
        }
        catch {
            Throw-VbaAccessGuidance($_.Exception.Message)
        }
    }

    Write-Host "Saving to $outputFile..."
    $wb.SaveAs($outputFile, 52) # xlOpenXMLWorkbookMacroEnabled
    Close-WorkbookSafely ([ref]$wb) $false
    Close-ExcelSafely ([ref]$excel)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($outputFile)

    $deleteSheets = @()
    for ($i = $wb.Worksheets.Count; $i -ge 1; $i--) {
        $candidate = $wb.Worksheets.Item($i)
        $usedRange = $candidate.UsedRange
        $isSingleCell = ($usedRange.Rows.Count -eq 1 -and $usedRange.Columns.Count -eq 1)
        $cellValue = ""
        if ($isSingleCell -and $null -ne $usedRange.Value2) {
            $cellValue = [string]$usedRange.Value2
        }
        $hasShapes = ($candidate.Shapes.Count -gt 0)

        if ($candidate.Name -ne "InazumaGantt_v3" -and $isSingleCell -and $cellValue -eq "" -and -not $hasShapes) {
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
    throw
}
finally {
    Close-WorkbookSafely ([ref]$wb) $false
    Close-ExcelSafely ([ref]$excel)
}
