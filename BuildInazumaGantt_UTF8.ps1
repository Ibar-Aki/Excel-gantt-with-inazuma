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
$utf8NoBom = [System.Text.UTF8Encoding]::new($false, $true)

function Throw-VbaAccessGuidance([string]$detail) {
    throw ("Excel workbook generation failed during VBA injection. " +
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

function Invoke-WorkbookMacroOrThrow($excelApp, $workbook, [string]$macroName) {
    try {
        $qualifiedMacro = "'{0}'!{1}" -f $workbook.Name, $macroName
        $excelApp.Run($qualifiedMacro)
    }
    catch {
        throw ("Macro smoke test failed for {0}: {1}" -f $macroName, $_.Exception.Message)
    }
}

function Invoke-WorksheetChangeSmoke($excelApp, $workbook, [string]$sheetName) {
    $worksheet = Get-WorksheetOrThrow $workbook $sheetName
    $startRow = 900
    $endRow = $startRow + 2
    $prevEvents = $excelApp.EnableEvents
    $prevScreenUpdating = $excelApp.ScreenUpdating
    $prevAlerts = $excelApp.DisplayAlerts

    try {
        $excelApp.ScreenUpdating = $false
        $excelApp.DisplayAlerts = $false
        $excelApp.EnableEvents = $false
        $worksheet.Range("A${startRow}:O${endRow}").ClearContents() | Out-Null

        $taskValues = New-Object 'object[,]' 3, 4
        $taskValues[0, 0] = "Smoke Parent"
        $taskValues[1, 1] = "Smoke Child A"
        $taskValues[2, 1] = "Smoke Child B"

        $rollupValues = New-Object 'object[,]' 3, 4
        $rollupValues[1, 1] = 1
        $rollupValues[1, 3] = 2
        $rollupValues[2, 1] = 0
        $rollupValues[2, 3] = 2

        $dateValues = New-Object 'object[,]' 3, 2
        $dateValues[1, 0] = "35/01/08"
        $dateValues[1, 1] = "35/01/09"
        $dateValues[2, 0] = "35/01/10"
        $dateValues[2, 1] = "35/01/11"

        $excelApp.EnableEvents = $true
        $worksheet.Range("C${startRow}:F${endRow}").Value2 = $taskValues
        $worksheet.Range("H${startRow}:K${endRow}").Value2 = $rollupValues
        $worksheet.Range("L${startRow}:M${endRow}").Value2 = $dateValues

        $parentProgress = [double]$worksheet.Range("I${startRow}").Value2
        $parentHours = [double]$worksheet.Range("K${startRow}").Value2
        $parentStart = $worksheet.Range("L${startRow}").Value2
        $parentEnd = $worksheet.Range("M${startRow}").Value2
        if ([math]::Abs($parentProgress - 0.5) -gt 0.001 -or [math]::Abs($parentHours - 4) -gt 0.001 -or $null -eq $parentStart -or $null -eq $parentEnd) {
            throw ("Worksheet_Change smoke test failed: parent progress={0}, hours={1}, start={2}, end={3}" -f $parentProgress, $parentHours, $parentStart, $parentEnd)
        }
    }
    catch {
        throw ("Worksheet_Change smoke test failed: {0}" -f $_.Exception.Message)
    }
    finally {
        try {
            $excelApp.EnableEvents = $false
            $worksheet.Range("A${startRow}:O${endRow}").ClearContents() | Out-Null
        }
        catch {
        }
        $excelApp.EnableEvents = $prevEvents
        $excelApp.ScreenUpdating = $prevScreenUpdating
        $excelApp.DisplayAlerts = $prevAlerts
    }
}

function Invoke-CoreSmokeMacros($excelApp, $workbook, [string]$mainSheetName) {
    Invoke-WorkbookMacroOrThrow $excelApp $workbook "RefreshInazumaGantt"
    Invoke-WorksheetChangeSmoke $excelApp $workbook $mainSheetName
    Invoke-WorkbookMacroOrThrow $excelApp $workbook "RefreshInazumaGantt"
    Invoke-WorkbookMacroOrThrow $excelApp $workbook "CreateWbsBackupSheetSilent"
    Invoke-WorkbookMacroOrThrow $excelApp $workbook "ToggleBulkEditMode"
    Invoke-WorkbookMacroOrThrow $excelApp $workbook "ToggleBulkEditMode"
    Invoke-WorkbookMacroOrThrow $excelApp $workbook "RestoreWbsFromBackupSheetSilent"
}

function Remove-WorksheetIfExists($excelApp, $workbook, [string]$sheetName, [string]$fallbackSheetName) {
    $targetSheet = $null
    try {
        $targetSheet = $workbook.Worksheets.Item($sheetName)
    }
    catch {
        return
    }

    $prevAlerts = $excelApp.DisplayAlerts
    try {
        $excelApp.DisplayAlerts = $false
        if (-not [string]::IsNullOrWhiteSpace($fallbackSheetName)) {
            try {
                $workbook.Worksheets.Item($fallbackSheetName).Activate() | Out-Null
            }
            catch {
            }
        }
        $targetSheet.Delete()
    }
    finally {
        $excelApp.DisplayAlerts = $prevAlerts
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
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookRef.Value) | Out-Null
            }
            catch {
            }
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
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }
}

function Read-Utf8Text([string]$path) {
    return [System.IO.File]::ReadAllText($path, $utf8NoBom)
}

function Get-VbaModuleNameFromSource([string]$content, [string]$fallbackName) {
    $match = [regex]::Match($content, '(?m)^Attribute VB_Name = "([^"]+)"')
    if ($match.Success) {
        return $match.Groups[1].Value
    }

    return $fallbackName
}

function Remove-VbaAttributeLines([string]$content) {
    return [regex]::Replace($content, '(?m)^Attribute [^\r\n]*\r?\n', '')
}

function Get-FirstNonAsciiSourceLine([string]$content) {
    $lines = $content -split "`r?`n"
    foreach ($line in $lines) {
        $trimmedLine = $line.Trim()
        if ($trimmedLine -and $trimmedLine -match '[^\u0000-\u007F]') {
            return $trimmedLine
        }
    }

    return $null
}

function Get-CodeModuleText($codeModule) {
    $lineCount = $codeModule.CountOfLines
    if ($lineCount -le 0) {
        return ""
    }

    return $codeModule.Lines(1, $lineCount)
}

function Assert-CodeModuleContainsSentinel($codeModule, [string]$moduleName, [string]$expectedText) {
    if ([string]::IsNullOrWhiteSpace($expectedText)) {
        return
    }

    $moduleText = Get-CodeModuleText $codeModule
    if (-not $moduleText.Contains($expectedText)) {
        throw ("Imported VBA text for module '{0}' did not retain sentinel '{1}'. " +
               "Possible mojibake during workbook generation." -f $moduleName, $expectedText)
    }
}

function Add-StandardModuleFromUtf8Source($workbook, [string]$sourcePath) {
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        Write-Warning "File not found: $sourcePath"
        return
    }

    $rawSource = Read-Utf8Text $sourcePath
    $moduleName = Get-VbaModuleNameFromSource $rawSource ([System.IO.Path]::GetFileNameWithoutExtension($sourcePath))
    $moduleBody = Remove-VbaAttributeLines $rawSource
    $sentinel = Get-FirstNonAsciiSourceLine $moduleBody

    Write-Host ("Injecting {0} as {1}..." -f ([System.IO.Path]::GetFileName($sourcePath)), $moduleName)
    try {
        $component = $workbook.VBProject.VBComponents.Add(1)
        $component.Name = $moduleName
        $component.CodeModule.AddFromString($moduleBody)
        Assert-CodeModuleContainsSentinel $component.CodeModule $moduleName $sentinel
    }
    catch {
        Throw-VbaAccessGuidance($_.Exception.Message)
    }
}

function Get-WorksheetDocumentCodeModule($workbook, [string]$worksheetName) {
    foreach ($component in $workbook.VBProject.VBComponents) {
        if ($component.Type -ne 100) {
            continue
        }

        try {
            if ($component.Properties.Item("Name").Value -eq $worksheetName) {
                return $component.CodeModule
            }
        }
        catch {
        }
    }

    throw "Worksheet VBComponent was not found for sheet: $worksheetName"
}

function Inject-WorksheetModuleFromUtf8Source($workbook, [string]$worksheetName, [string]$sourcePath) {
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        Write-Warning "File not found: $sourcePath"
        return
    }

    $rawSource = Read-Utf8Text $sourcePath
    $moduleName = Get-VbaModuleNameFromSource $rawSource ([System.IO.Path]::GetFileNameWithoutExtension($sourcePath))
    $moduleBody = Remove-VbaAttributeLines $rawSource
    $sentinel = Get-FirstNonAsciiSourceLine $moduleBody

    Write-Host ("Injecting worksheet module from {0}..." -f ([System.IO.Path]::GetFileName($sourcePath)))
    try {
        $codeModule = Get-WorksheetDocumentCodeModule $workbook $worksheetName
        if ($codeModule.CountOfLines -gt 0) {
            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
        }
        $codeModule.AddFromString($moduleBody)
        Assert-CodeModuleContainsSentinel $codeModule $moduleName $sentinel
    }
    catch {
        Throw-VbaAccessGuidance($_.Exception.Message)
    }
}

if (!(Test-Path $outputDir)) { New-Item -ItemType Directory -Path $outputDir | Out-Null }
if (Test-Path $outputFile) { Remove-Item $outputFile -Force }

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    Write-Host "Creating new workbook..."
    $wb = $excel.Workbooks.Add()
    $coreModules = @(
        "InazumaGantt_v3_UTF8.bas",
        "WBSParentRollup_UTF8.bas",
        "WBSRoadmapReport_UTF8.bas",
        "WBSSampleShowcase_UTF8.bas",
        "HierarchyColor_UTF8.bas",
        "SetupWizard_UTF8.bas"
    )

    foreach ($file in $coreModules) {
        Add-StandardModuleFromUtf8Source $wb (Join-Path $vbaDir $file)
    }

    Write-Host "Running SilentSetup..."
    try {
        $excel.Run(("'{0}'!SilentSetup" -f $wb.Name), $true)
        Write-Host "SilentSetup completed successfully."
    }
    catch {
        throw "SilentSetup failed: $($_.Exception.Message)"
    }

    $mainSheet = Get-WorksheetOrThrow $wb "InazumaGantt_v3"
    Inject-WorksheetModuleFromUtf8Source $wb $mainSheet.Name (Join-Path $vbaDir "SheetModule_UTF8.bas")

    Write-Host "Running in-memory smoke tests..."
    Invoke-CoreSmokeMacros $excel $wb "InazumaGantt_v3"
    Remove-WorksheetIfExists $excel $wb "WBS_Backup_v3" "InazumaGantt_v3"

    Write-Host "Saving to $outputFile..."
    $wb.SaveAs($outputFile, 52)
    Close-WorkbookSafely ([ref]$wb) $false
    Close-ExcelSafely ([ref]$excel)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($outputFile)

    Write-Host "Running post-save smoke tests..."
    Invoke-CoreSmokeMacros $excel $wb "InazumaGantt_v3"
    Remove-WorksheetIfExists $excel $wb "WBS_Backup_v3" "InazumaGantt_v3"

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
    }

    $wb.Save()
    Write-Host "Build Complete!"
}
catch {
    Write-Error "Error occurred: $_"
    throw
}
finally {
    Remove-Variable -Name @("candidate", "usedRange", "mainSheet", "component", "codeModule", "targetSheet") -ErrorAction SilentlyContinue
    Close-WorkbookSafely ([ref]$wb) $false
    Close-ExcelSafely ([ref]$excel)
}
