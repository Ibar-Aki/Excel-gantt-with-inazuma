param(
    [string]$ProjectRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [string]$OutputRoot,
    [string]$V3WorkbookPath,
    [string]$LiteWorkbookPath,
    [switch]$SkipRestored,
    [switch]$KeepWorkbooks
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 3.0

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $OutputRoot = Join-Path $ProjectRoot ("output\acceptance_test_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
}

$utf8NoBom = [System.Text.UTF8Encoding]::new($false, $true)
$results = New-Object System.Collections.Generic.List[object]

function Add-TestResult {
    param(
        [string]$Id,
        [string]$Edition,
        [string]$Name,
        [string]$Status,
        [string]$Message
    )

    $script:results.Add([pscustomobject]@{
        id = $Id
        edition = $Edition
        name = $Name
        status = $Status
        message = $Message
    })
}

function Invoke-TestCase {
    param(
        [string]$Id,
        [string]$Edition,
        [string]$Name,
        [scriptblock]$Body
    )

    $progressPath = Join-Path $OutputRoot "acceptance-progress.log"
    Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ("RUN {0} {1} {2} {3}" -f (Get-Date -Format "HH:mm:ss"), $Edition, $Id, $Name)

    try {
        & $Body
        Add-TestResult -Id $Id -Edition $Edition -Name $Name -Status "OK" -Message ""
        Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ("OK  {0} {1} {2} {3}" -f (Get-Date -Format "HH:mm:ss"), $Edition, $Id, $Name)
    }
    catch {
        $details = $_.Exception.Message
        if (-not [string]::IsNullOrWhiteSpace([string]$_.ScriptStackTrace)) {
            $details = $details + " | " + ([string]$_.ScriptStackTrace).Replace("`r", " ").Replace("`n", " ")
        }
        Add-TestResult -Id $Id -Edition $Edition -Name $Name -Status "NG" -Message $details
        Add-Content -LiteralPath $progressPath -Encoding UTF8 -Value ("NG  {0} {1} {2} {3} {4}" -f (Get-Date -Format "HH:mm:ss"), $Edition, $Id, $Name, $details)
    }
}

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    if (-not $Condition) {
        throw $Message
    }
}

function Assert-Near {
    param(
        [double]$Actual,
        [double]$Expected,
        [double]$Tolerance,
        [string]$Message
    )

    if ([math]::Abs($Actual - $Expected) -gt $Tolerance) {
        throw ("{0}: expected={1}, actual={2}" -f $Message, $Expected, $Actual)
    }
}

function Get-LatestWorkbook {
    param(
        [string]$Directory,
        [string]$Filter
    )

    $file = Get-ChildItem -LiteralPath $Directory -Filter $Filter -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -eq $file) {
        throw "Workbook not found: $Directory / $Filter"
    }
    return $file.FullName
}

function Convert-A1Address {
    param([string]$Address)

    if ($Address -notmatch '^([A-Z]+)(\d+)$') {
        throw "Only single-cell A1 addresses are supported: $Address"
    }

    $column = 0
    foreach ($char in $matches[1].ToCharArray()) {
        $column = ($column * 26) + ([int][char]$char - [int][char]'A' + 1)
    }

    return [pscustomobject]@{
        Row = [int]$matches[2]
        Column = $column
    }
}

function Get-CellText {
    param($Worksheet, $Address)
    $cell = Convert-A1Address ([string]$Address)
    $range = $Worksheet.Cells.Item([int]$cell.Row, [int]$cell.Column)
    return [string]$range.Text
}

function Get-CellValue {
    param($Worksheet, $Address)
    $cell = Convert-A1Address ([string]$Address)
    $range = $Worksheet.Cells.Item([int]$cell.Row, [int]$cell.Column)
    return $range.Value2
}

function Set-CellValue {
    param($Worksheet, $Address, $Value)
    $cell = Convert-A1Address ([string]$Address)
    $range = $Worksheet.Cells.Item([int]$cell.Row, [int]$cell.Column)
    if ($Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        $range.FormulaR1C1 = [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
    }
    else {
        $range.Value2 = $Value
    }
}

function Clear-TestRows {
    param($Excel, $Worksheet, [int]$StartRow, [int]$EndRow)

    $previousEvents = $Excel.EnableEvents
    try {
        $Excel.EnableEvents = $false
        $Worksheet.Range("A${StartRow}:O${EndRow}").ClearContents() | Out-Null
    }
    finally {
        $Excel.EnableEvents = $previousEvents
    }
}

function Find-Shape {
    param($Worksheet, [string]$Name)

    foreach ($shape in $Worksheet.Shapes) {
        if ($shape.Name -eq $Name) {
            return $shape
        }
    }
    return $null
}

function Get-ShapeText {
    param($Shape)

    try {
        return [string]$Shape.TextFrame2.TextRange.Text
    }
    catch {
        try {
            return [string]$Shape.TextFrame.Characters().Text
        }
        catch {
            return ""
        }
    }
}

function Count-ShapesByName {
    param($Worksheet, [string]$Name)

    $count = 0
    foreach ($shape in $Worksheet.Shapes) {
        if ($shape.Name -eq $Name) {
            $count++
        }
    }
    return $count
}

function Get-ManagedShapeCounts {
    param($Worksheet)

    $buttonCount = 0
    $barCount = 0
    $todayCount = 0
    $inazumaCount = 0
    foreach ($shape in $Worksheet.Shapes) {
        $name = [string]$shape.Name
        if ($name.StartsWith("Btn_")) { $buttonCount++ }
        if ($name.StartsWith("Bar_")) { $barCount++ }
        if ($name.StartsWith("Today_")) { $todayCount++ }
        if ($name.StartsWith("Inazuma_")) { $inazumaCount++ }
    }

    return [pscustomobject]@{
        buttons = $buttonCount
        bars = $barCount
        today = $todayCount
        inazuma = $inazumaCount
    }
}

function Test-ButtonLayout {
    param($Worksheet, [string]$Edition)

    $expectedNames = @("Btn_Refresh", "Btn_ToggleWeekend", "Btn_BulkEdit", "Btn_ShiftDates", "Btn_ExportPDF")
    foreach ($name in $expectedNames) {
        Assert-True ((Count-ShapesByName $Worksheet $name) -eq 1) "${Edition}: $name count is not 1"
        $shape = Find-Shape $Worksheet $name
        Assert-True ($null -ne $shape) "${Edition}: $name not found"
        Assert-True ($shape.Width -gt 30 -and $shape.Height -gt 10) "${Edition}: $name size is invalid"
        Assert-True ((Get-ShapeText $shape).Trim().Length -gt 0) "${Edition}: $name text is empty"
    }

    $buttons = @()
    foreach ($shape in $Worksheet.Shapes) {
        if ([string]$shape.Name -like "Btn_*") {
            $buttons += $shape
        }
    }

    for ($i = 0; $i -lt $buttons.Count; $i++) {
        for ($j = $i + 1; $j -lt $buttons.Count; $j++) {
            $a = $buttons[$i]
            $b = $buttons[$j]
            $overlapX = ($a.Left -lt ($b.Left + $b.Width)) -and (($a.Left + $a.Width) -gt $b.Left)
            $overlapY = ($a.Top -lt ($b.Top + $b.Height)) -and (($a.Top + $a.Height) -gt $b.Top)
            Assert-True (-not ($overlapX -and $overlapY)) "${Edition}: button overlap detected: $($a.Name) / $($b.Name)"
        }
    }
}

function Set-BulkEditState {
    param($Excel, $Workbook, $Settings, [bool]$Enabled)

    $current = [bool]$Settings.Range("B9").Value2
    if ($current -ne $Enabled) {
        $Excel.Run(("'{0}'!ToggleBulkEditMode" -f $Workbook.Name))
    } elseif ($Enabled) {
        $Excel.EnableEvents = $false
    } else {
        $Excel.EnableEvents = $true
    }
    Assert-True (([bool]$Settings.Range("B9").Value2) -eq $Enabled) "Bulk edit setting did not become $Enabled"
    Assert-True ($Excel.EnableEvents -eq (-not $Enabled)) "EnableEvents did not match bulk edit state $Enabled"
}

function Set-AutomationMode {
    param($Excel, $Workbook, $Settings, [bool]$Enabled)

    $Excel.Run(("'{0}'!SetAutomationMode" -f $Workbook.Name), $Enabled)
    Assert-True (([bool]$Settings.Range("B12").Value2) -eq $Enabled) "Automation mode did not become $Enabled"
}

function Count-LogEvents {
    param($Workbook, [string]$EventName)

    $logSheet = $null
    try {
        $logSheet = $Workbook.Worksheets.Item("_InazumaGantt_Log")
    }
    catch {
        return 0
    }

    $lastRow = $logSheet.Cells($logSheet.Rows.Count, "B").End(-4162).Row
    $count = 0
    for ($row = 2; $row -le $lastRow; $row++) {
        if ([string]$logSheet.Cells($row, "B").Value2 -eq $EventName) {
            $count++
        }
    }
    return $count
}

function Prepare-GanttRows {
    param($Excel, $Worksheet, [int]$StartRow, [bool]$HasActual)

    $previousEvents = $Excel.EnableEvents
    try {
        $Excel.EnableEvents = $false
        Clear-TestRows $Excel $Worksheet $StartRow ($StartRow + 7)
        $today = Get-Date
        Set-CellValue $Worksheet "L2" ($today.AddDays(-7).ToOADate())

        Set-CellValue $Worksheet "C${StartRow}" "Acceptance Parent"
        Set-CellValue $Worksheet ("D" + ($StartRow + 1)) "Acceptance Child A"
        Set-CellValue $Worksheet ("D" + ($StartRow + 2)) "Acceptance Child B"
        Set-CellValue $Worksheet ("D" + ($StartRow + 3)) "Acceptance Child C"
        Set-CellValue $Worksheet ("D" + ($StartRow + 4)) "Acceptance Child D"

        for ($i = 0; $i -lt 5; $i++) {
            $row = $StartRow + $i
            Set-CellValue $Worksheet "L${row}" ($today.AddDays($i - 3).ToOADate())
            Set-CellValue $Worksheet "M${row}" ($today.AddDays($i + 4).ToOADate())
            Set-CellValue $Worksheet "I${row}" (0.2 + ($i * 0.15))
            Set-CellValue $Worksheet "K${row}" (1 + $i)
            if ($HasActual) {
                Set-CellValue $Worksheet "N${row}" ($today.AddDays($i - 2).ToOADate())
                Set-CellValue $Worksheet "O${row}" ($today.AddDays($i + 2).ToOADate())
            }
        }
    }
    finally {
        $Excel.EnableEvents = $previousEvents
    }
}

function Test-GanttShapeAlignment {
    param($Worksheet, [int]$StartRow, [int]$EndRow, [bool]$HasActual, [string]$Edition)

    $ganttStartCol = $Worksheet.Columns("P").Column
    $ganttEndCol = $ganttStartCol + 120 - 1
    $leftLimit = [double]$Worksheet.Cells($StartRow, $ganttStartCol).Left - 1
    $rightLimit = [double]($Worksheet.Cells($StartRow, $ganttEndCol).Left + $Worksheet.Cells($StartRow, $ganttEndCol).Width) + 1
    $topLimit = [double]$Worksheet.Cells($StartRow, $ganttStartCol).Top - 1
    $bottomLimit = [double]($Worksheet.Cells($EndRow, $ganttStartCol).Top + $Worksheet.Cells($EndRow, $ganttStartCol).Height) + 1

    $barCount = 0
    $actualCount = 0
    foreach ($shape in $Worksheet.Shapes) {
        $name = [string]$shape.Name
        $targetBar = $false
        if ($name -match '^Bar_(Plan|Actual)_(\d+)$') {
            $shapeRow = [int]$matches[2]
            $targetBar = ($shapeRow -ge $StartRow -and $shapeRow -le $EndRow)
        }
        elseif ($name.StartsWith("Today_") -or $name.StartsWith("Inazuma_")) {
            Assert-True ([double]$shape.Left -ge $leftLimit) "${Edition}: $name is left of gantt area"
            Assert-True (([double]$shape.Left + [double]$shape.Width) -le ($rightLimit + 5)) "${Edition}: $name is right of gantt area"
        }

        if ($targetBar) {
            Assert-True ([double]$shape.Left -ge $leftLimit) "${Edition}: $name is left of gantt area"
            Assert-True (([double]$shape.Left + [double]$shape.Width) -le ($rightLimit + 5)) "${Edition}: $name is right of gantt area"
            Assert-True ([double]$shape.Top -ge ($topLimit - 5)) "${Edition}: $name is above data area"
            Assert-True (([double]$shape.Top + [double]$shape.Height) -le ($bottomLimit + 20)) "${Edition}: $name is below expected data area"
        }
        if ($targetBar -and $name.StartsWith("Bar_Plan_")) {
            $barCount++
        }
        if ($targetBar -and $name.StartsWith("Bar_Actual_")) {
            $actualCount++
        }
    }

    Assert-True ($barCount -ge 4) "${Edition}: expected at least four plan bars, got $barCount"

    if ($HasActual) {
        Assert-True ($actualCount -ge 1) "${Edition}: expected actual bars"
    }
}

function Invoke-WorkbookAcceptance {
    param(
        $Excel,
        [string]$WorkbookPath,
        [string]$Edition,
        [string]$SheetName,
        [bool]$HasActual,
        [string]$WorkDir
    )

    $workbook = $null
    $settings = $null
    try {
        $workbook = $Excel.Workbooks.Open($WorkbookPath)
        $worksheet = $workbook.Worksheets.Item($SheetName)
        $settings = $workbook.Worksheets.Item("設定マスタ")
        $worksheet.Activate() | Out-Null
        $Excel.EnableEvents = $true
        Set-AutomationMode $Excel $workbook $settings $true

        Invoke-TestCase "TC-01" $Edition "初期表示とボタン配置" {
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Test-ButtonLayout $worksheet $Edition
            $indicator = Get-CellText $worksheet "A3"
            Assert-True ($indicator -like "*高速入力 OFF*") "${Edition}: initial indicator is not OFF: $indicator"
            Assert-True (-not [bool]$worksheet.Range("H9").Validation.ShowInput) "${Edition}: status input message is shown"
            Assert-True (-not [bool]$worksheet.Range("I9").Validation.ShowInput) "${Edition}: progress input message is shown"
            Assert-True (-not [bool]$worksheet.Range("H20").Validation.ShowInput) "${Edition}: status input message is shown outside first row"
            Assert-True (-not [bool]$worksheet.Range("I20").Validation.ShowInput) "${Edition}: progress input message is shown outside first row"
            $runtimeState = [string]$Excel.Run(("'{0}'!CheckInazumaRuntimeState" -f $workbook.Name), $true)
            Assert-True ($runtimeState -eq "OK") "${Edition}: runtime state is $runtimeState"
        }

        Invoke-TestCase "TC-02" $Edition "高速入力 ON ボタン押下" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 938 939
            $Excel.Run(("'{0}'!ToggleBulkEditMode" -f $workbook.Name))
            $button = Find-Shape $worksheet "Btn_BulkEdit"
            Assert-True ((Get-ShapeText $button) -like "*ON*") "${Edition}: bulk button text is not ON"
            Assert-True ((Get-CellText $worksheet "A3") -like "*高速入力 ON*") "${Edition}: indicator is not ON"
            Assert-True (-not $Excel.EnableEvents) "${Edition}: EnableEvents should be false in ON"
            Set-CellValue $worksheet "D938" "Bulk Level 2 Hint"
            Set-CellValue $worksheet "F939" "Bulk Level 4 Hint"
            $worksheet.Calculate()
            Assert-Near ([double](Get-CellValue $worksheet "A938")) 2 0.001 "${Edition}: bulk LV hint for D column"
            Assert-Near ([double](Get-CellValue $worksheet "A939")) 4 0.001 "${Edition}: bulk LV hint for F column"
        }

        Invoke-TestCase "TC-03" $Edition "高速入力 ON 中の Undo 保護状態" {
            Set-BulkEditState $Excel $workbook $settings $true
            $before = Get-CellValue $worksheet "C950"
            Set-CellValue $worksheet "C950" "Undo Smoke"
            Assert-True (-not $Excel.EnableEvents) "${Edition}: EnableEvents changed during bulk input"
            Assert-True (([bool]$settings.Range("B9").Value2) -eq $true) "${Edition}: bulk setting changed during bulk input"
            Clear-TestRows $Excel $worksheet 950 950
            if ($null -ne $before) {
                Set-CellValue $worksheet "C950" $before
            }
        }

        Invoke-TestCase "TC-04" $Edition "高速入力 OFF 復帰時の一括再整合" {
            Set-BulkEditState $Excel $workbook $settings $true
            Prepare-GanttRows $Excel $worksheet 930 $HasActual
            $Excel.Run(("'{0}'!ToggleBulkEditMode" -f $workbook.Name))
            Assert-True (-not ([bool]$settings.Range("B9").Value2)) "${Edition}: bulk setting is still ON"
            Assert-True ($Excel.EnableEvents) "${Edition}: EnableEvents was not restored"
            Assert-True ((Get-CellText $worksheet "A3") -like "*高速入力 OFF*") "${Edition}: indicator is not OFF"
            Assert-True ((Count-ShapesByName $worksheet "Btn_BulkEdit") -eq 1) "${Edition}: bulk button duplicated"
            Test-GanttShapeAlignment $worksheet 930 934 $HasActual $Edition
        }

        Invoke-TestCase "TC-05" $Edition "高速入力ボタン連打" {
            Set-BulkEditState $Excel $workbook $settings $false
            for ($i = 0; $i -lt 10; $i++) {
                $Excel.Run(("'{0}'!ToggleBulkEditMode" -f $workbook.Name))
            }
            Assert-True (-not ([bool]$settings.Range("B9").Value2)) "${Edition}: even toggle count did not end OFF"
            Assert-True ($Excel.EnableEvents) "${Edition}: EnableEvents not true after even toggle count"
            Test-ButtonLayout $worksheet $Edition
        }

        Invoke-TestCase "TC-06" $Edition "状況列の直接変更" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 940 945
            Set-CellValue $worksheet "C940" "Status Not Started"
            Set-CellValue $worksheet "C941" "Status In Progress"
            Set-CellValue $worksheet "C942" "Status Completed"
            Set-CellValue $worksheet "C943" "Status Hold"
            Set-CellValue $worksheet "I943" 0.7
            Set-CellValue $worksheet "H940" "未着手"
            Set-CellValue $worksheet "H941" "進行中"
            Set-CellValue $worksheet "H942" "完了"
            Set-CellValue $worksheet "H943" "保留"
            Assert-True ((Get-CellText $worksheet "H940") -eq "未着手") "${Edition}: H940"
            Assert-Near ([double](Get-CellValue $worksheet "I940")) 0 0.001 "${Edition}: I940"
            Assert-True ((Get-CellText $worksheet "H941") -eq "進行中") "${Edition}: H941"
            Assert-Near ([double](Get-CellValue $worksheet "I941")) 0.5 0.001 "${Edition}: I941"
            Assert-True ((Get-CellText $worksheet "H942") -eq "完了") "${Edition}: H942"
            Assert-Near ([double](Get-CellValue $worksheet "I942")) 1 0.001 "${Edition}: I942"
            Assert-True ((Get-CellText $worksheet "H943") -eq "保留") "${Edition}: H943"
            Assert-Near ([double](Get-CellValue $worksheet "I943")) 0.7 0.001 "${Edition}: I943"
        }

        Invoke-TestCase "TC-07" $Edition "進捗率列の直接変更" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 946 952
            for ($i = 0; $i -lt 6; $i++) {
                Set-CellValue $worksheet ("C" + (946 + $i)) "Progress Case $i"
            }
            $inputs = @(0, "0%", 0.7, 70, "70%", "100%")
            for ($i = 0; $i -lt $inputs.Count; $i++) {
                Set-CellValue $worksheet ("I" + (946 + $i)) $inputs[$i]
            }
            Assert-True ((Get-CellText $worksheet "H946") -eq "未着手") "${Edition}: 0 status"
            Assert-True ((Get-CellText $worksheet "H947") -eq "未着手") "${Edition}: 0% status"
            Assert-True ((Get-CellText $worksheet "H948") -eq "進行中") "${Edition}: 0.7 status"
            Assert-True ((Get-CellText $worksheet "H949") -eq "進行中") "${Edition}: 70 status"
            Assert-True ((Get-CellText $worksheet "H950") -eq "進行中") "${Edition}: 70% status"
            Assert-True ((Get-CellText $worksheet "H951") -eq "完了") "${Edition}: 100% status"
            Assert-Near ([double](Get-CellValue $worksheet "I949")) 0.7 0.001 "${Edition}: 70 normalized"
            Assert-Near ([double](Get-CellValue $worksheet "I951")) 1 0.001 "${Edition}: 100 normalized"
        }

        Invoke-TestCase "TC-08" $Edition "不正な進捗率入力後の復旧" {
            Set-AutomationMode $Excel $workbook $settings $true
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 953 954
            Set-CellValue $worksheet "C953" "Invalid Progress Minus"
            Set-CellValue $worksheet "C954" "Invalid Progress Text"
            $worksheet.Range("I953:I954").Validation.Delete()
            $beforeInvalidLogCount = Count-LogEvents $workbook "InvalidProgressInput"
            Set-CellValue $worksheet "I953" -1
            Set-CellValue $worksheet "I954" "not-a-progress"
            Assert-True ($Excel.EnableEvents) "${Edition}: EnableEvents was not restored after invalid progress"
            Assert-True ((Get-CellText $worksheet "H953") -eq "未着手") "${Edition}: invalid -1 status"
            Assert-Near ([double](Get-CellValue $worksheet "I953")) 0 0.001 "${Edition}: invalid -1 progress"
            Assert-True ((Get-CellText $worksheet "H954") -eq "未着手") "${Edition}: invalid text status"
            Assert-Near ([double](Get-CellValue $worksheet "I954")) 0 0.001 "${Edition}: invalid text progress"
            $afterInvalidLogCount = Count-LogEvents $workbook "InvalidProgressInput"
            Assert-True ($afterInvalidLogCount -ge ($beforeInvalidLogCount + 2)) "${Edition}: invalid progress was not logged"
        }

        Invoke-TestCase "TC-09" $Edition "高速入力 ON 中に状況・進捗率を編集してから OFF" {
            Clear-TestRows $Excel $worksheet 955 958
            Set-BulkEditState $Excel $workbook $settings $true
            for ($i = 0; $i -lt 4; $i++) { Set-CellValue $worksheet ("C" + (955 + $i)) "Bulk HI $i" }
            Set-CellValue $worksheet "H955" "進行中"
            Set-CellValue $worksheet "H956" "保留"
            Set-CellValue $worksheet "I957" "70%"
            Set-CellValue $worksheet "I958" "100%"
            Set-BulkEditState $Excel $workbook $settings $false
            Assert-True ((Get-CellText $worksheet "H955") -eq "進行中") "${Edition}: bulk H955"
            Assert-Near ([double](Get-CellValue $worksheet "I955")) 0.5 0.001 "${Edition}: bulk I955"
            Assert-True ((Get-CellText $worksheet "H956") -eq "保留") "${Edition}: bulk H956"
            Assert-True ((Get-CellText $worksheet "H957") -eq "進行中") "${Edition}: bulk H957"
            Assert-True ((Get-CellText $worksheet "H958") -eq "完了") "${Edition}: bulk H958"
        }

        Invoke-TestCase "TC-10" $Edition "ダブルクリック完了と高速入力の関係" {
            Clear-TestRows $Excel $worksheet 959 960
            Set-BulkEditState $Excel $workbook $settings $false
            Set-CellValue $worksheet "C959" "Double Click Equivalent"
            $Excel.EnableEvents = $false
            Set-CellValue $worksheet "I959" 1
            Set-CellValue $worksheet "H959" "完了"
            $Excel.EnableEvents = $true
            Assert-True ((Get-CellText $worksheet "H959") -eq "完了") "${Edition}: completion status"
            Assert-Near ([double](Get-CellValue $worksheet "I959")) 1 0.001 "${Edition}: completion progress"
            Set-BulkEditState $Excel $workbook $settings $true
            Set-CellValue $worksheet "C960" "Bulk Double Click Guard"
            Assert-True (([bool]$settings.Range("B9").Value2) -eq $true) "${Edition}: bulk mode changed unexpectedly"
            Set-BulkEditState $Excel $workbook $settings $false
        }

        Invoke-TestCase "TC-11" $Edition "ガント更新ボタン後の図形ずれ" {
            Set-BulkEditState $Excel $workbook $settings $false
            Prepare-GanttRows $Excel $worksheet 930 $HasActual
            $beforeGanttLogCount = Count-LogEvents $workbook "GanttRefresh"
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Test-GanttShapeAlignment $worksheet 930 934 $HasActual $Edition
            $afterGanttLogCount = Count-LogEvents $workbook "GanttRefresh"
            Assert-True ($afterGanttLogCount -gt $beforeGanttLogCount) "${Edition}: gantt refresh was not logged"
        }

        Invoke-TestCase "TC-12" $Edition "土日切替後の図形ずれ" {
            Set-BulkEditState $Excel $workbook $settings $false
            Prepare-GanttRows $Excel $worksheet 930 $HasActual
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            $Excel.Run(("'{0}'!ToggleWeekends" -f $workbook.Name))
            $Excel.Run(("'{0}'!ToggleWeekends" -f $workbook.Name))
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Test-GanttShapeAlignment $worksheet 930 934 $HasActual $Edition
        }

        Invoke-TestCase "TC-13" $Edition "開始日変更後のヘッダー・図形整合" {
            Set-BulkEditState $Excel $workbook $settings $false
            Prepare-GanttRows $Excel $worksheet 930 $HasActual
            Set-CellValue $worksheet "L2" ((Get-Date).AddDays(-14).ToOADate())
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Test-GanttShapeAlignment $worksheet 930 934 $HasActual $Edition
        }

        Invoke-TestCase "TC-14" $Edition "保存・再オープン後の高速入力状態" {
            Set-BulkEditState $Excel $workbook $settings $true
            $copyPath = Join-Path $WorkDir ("reopen_state_" + $Edition + ".xlsm")
            if (Test-Path -LiteralPath $copyPath) { Remove-Item -LiteralPath $copyPath -Force }
            $workbook.SaveCopyAs($copyPath)
            Set-BulkEditState $Excel $workbook $settings $false
            $Excel.EnableEvents = $true
            $reopened = $Excel.Workbooks.Open($copyPath)
            try {
                $reWs = $reopened.Worksheets.Item($SheetName)
                $reSettings = $reopened.Worksheets.Item("設定マスタ")
                $reWs.Activate() | Out-Null
                $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $reopened.Name))
                Assert-True (-not ([bool]$reSettings.Range("B9").Value2)) "${Edition}: reopened workbook did not repair bulk setting"
                Assert-True ($Excel.EnableEvents) "${Edition}: reopened workbook did not restore events"
            }
            finally {
                $reopened.Close($false) | Out-Null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($reopened) | Out-Null
            }
        }

        Invoke-TestCase "TC-15" $Edition "バックアップ・復元ボタンとの組み合わせ" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 961 962
            Set-CellValue $worksheet "C961" "Backup Original"
            Set-CellValue $worksheet "I961" "70%"
            $Excel.Run(("'{0}'!CreateWbsBackupSheetSilent" -f $workbook.Name))
            Set-BulkEditState $Excel $workbook $settings $true
            Set-CellValue $worksheet "C961" "Backup Changed"
            Set-CellValue $worksheet "I961" "100%"
            Set-BulkEditState $Excel $workbook $settings $false
            $Excel.Run(("'{0}'!RestoreWbsFromBackupSheetSilent" -f $workbook.Name))
            Assert-True ((Get-CellText $worksheet "C961") -eq "Backup Original") "${Edition}: backup restore did not restore task"
            Assert-True ((Get-CellText $worksheet "H961") -eq "進行中") "${Edition}: backup restore did not restore status"
            Assert-Near ([double](Get-CellValue $worksheet "I961")) 0.7 0.001 "${Edition}: backup restore did not restore progress"
            Assert-True (-not ([bool]$settings.Range("B9").Value2)) "${Edition}: bulk setting not OFF after restore"
        }

        Invoke-TestCase "TC-17" $Edition "WBSサマリの半月予定なし表示" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 963 965
            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                Set-CellValue $worksheet "C963" "Roadmap First Half"
                Set-CellValue $worksheet "L963" ([datetime]"2026-06-02").ToOADate()
                Set-CellValue $worksheet "M963" ([datetime]"2026-06-10").ToOADate()
                Set-CellValue $worksheet "C964" "Roadmap Second Half"
                Set-CellValue $worksheet "L964" ([datetime]"2026-06-22").ToOADate()
                Set-CellValue $worksheet "M964" ([datetime]"2026-06-26").ToOADate()
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            $Excel.Run(("'{0}'!CreateRoadmapOverviewSheet" -f $workbook.Name), [datetime]"2026-06-15")
            $roadmap = $workbook.Worksheets.Item("WBSサマリ")
            $juneColumn = 0
            $julyColumn = 0
            for ($col = 8; $col -le 20; $col++) {
                $headerText = [string]$roadmap.Cells.Item(5, $col).Text
                if ($headerText -eq "6月") { $juneColumn = $col }
                if ($headerText -eq "7月") { $julyColumn = $col }
            }
            Assert-True ($juneColumn -gt 0) "${Edition}: June column was not found in WBS summary"
            Assert-True ($julyColumn -gt 0) "${Edition}: July column was not found in WBS summary"

            $firstHalfRow = 0
            $secondHalfRow = 0
            for ($row = 6; $row -le 200; $row++) {
                $itemText = [string]$roadmap.Cells.Item($row, 2).Text
                if ($itemText -like "*Roadmap First Half*") { $firstHalfRow = $row }
                if ($itemText -like "*Roadmap Second Half*") { $secondHalfRow = $row }
            }
            Assert-True ($firstHalfRow -gt 0) "${Edition}: first-half roadmap row was not found"
            Assert-True ($secondHalfRow -gt 0) "${Edition}: second-half roadmap row was not found"

            $firstHalfCell = $roadmap.Cells.Item($firstHalfRow, $juneColumn)
            $secondHalfCell = $roadmap.Cells.Item($secondHalfRow, $juneColumn)
            Assert-True ([string]$firstHalfCell.Text -eq "□□") "${Edition}: first half roadmap text should have two square slots"
            Assert-True ([string]$secondHalfCell.Text -eq "□□") "${Edition}: second half roadmap text should have two square slots"
            Assert-True ([int]$firstHalfCell.Characters(1, 1).Font.Color -eq 10921638) "${Edition}: first-half row first slot should be planned square"
            Assert-True ([int]$firstHalfCell.Characters(2, 1).Font.Color -eq 16777215) "${Edition}: first-half row second slot should be white empty square"
            Assert-True ([int]$secondHalfCell.Characters(1, 1).Font.Color -eq 16777215) "${Edition}: second-half row first slot should be white empty square"
            Assert-True ([int]$secondHalfCell.Characters(2, 1).Font.Color -eq 10921638) "${Edition}: second-half row second slot should be planned square"
            Assert-True ([string]$roadmap.Cells.Item($firstHalfRow, $julyColumn).Text -eq "") "${Edition}: month without schedule should be blank"
        }

        Invoke-TestCase "TC-18" $Edition "高速入力ON中のLVライブ表示網羅" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 966 969
            $Excel.Run(("'{0}'!ToggleBulkEditMode" -f $workbook.Name))
            Assert-True (-not $Excel.EnableEvents) "${Edition}: EnableEvents should be false in ON for LV live test"

            Set-CellValue $worksheet "C966" "Bulk LV1 Live"
            Set-CellValue $worksheet "D967" "Bulk LV2 Live"
            Set-CellValue $worksheet "E968" "Bulk LV3 Live"
            Set-CellValue $worksheet "F969" "Bulk LV4 Live"
            $worksheet.Calculate()

            Assert-Near ([double](Get-CellValue $worksheet "A966")) 1 0.001 "${Edition}: bulk LV hint for C column"
            Assert-Near ([double](Get-CellValue $worksheet "A967")) 2 0.001 "${Edition}: bulk LV hint for D column"
            Assert-Near ([double](Get-CellValue $worksheet "A968")) 3 0.001 "${Edition}: bulk LV hint for E column"
            Assert-Near ([double](Get-CellValue $worksheet "A969")) 4 0.001 "${Edition}: bulk LV hint for F column"
            Assert-True ([bool]$worksheet.Range("A966").HasFormula) "${Edition}: live LV formula was not installed while ON"

            Set-BulkEditState $Excel $workbook $settings $false
            Assert-True ($Excel.EnableEvents) "${Edition}: EnableEvents was not restored after LV live OFF"
            Assert-Near ([double](Get-CellValue $worksheet "A966")) 1 0.001 "${Edition}: LV1 after OFF"
            Assert-Near ([double](Get-CellValue $worksheet "A967")) 2 0.001 "${Edition}: LV2 after OFF"
            Assert-Near ([double](Get-CellValue $worksheet "A968")) 3 0.001 "${Edition}: LV3 after OFF"
            Assert-Near ([double](Get-CellValue $worksheet "A969")) 4 0.001 "${Edition}: LV4 after OFF"
            Assert-True (-not [bool]$worksheet.Range("A966").HasFormula) "${Edition}: LV1 formula remained after OFF"
            Assert-True (-not [bool]$worksheet.Range("A967").HasFormula) "${Edition}: LV2 formula remained after OFF"
            Assert-True (-not [bool]$worksheet.Range("A968").HasFormula) "${Edition}: LV3 formula remained after OFF"
            Assert-True (-not [bool]$worksheet.Range("A969").HasFormula) "${Edition}: LV4 formula remained after OFF"
        }

        Invoke-TestCase "TC-19" $Edition "下位タスク警告の上位伝播" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 970 975
            $today = (Get-Date).Date
            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                Set-CellValue $worksheet "C970" "Alert Parent LV1"
                Set-CellValue $worksheet "D971" "Alert Parent LV2"
                Set-CellValue $worksheet "E972" "Alert Child LV3 Delayed"
                Set-CellValue $worksheet "L972" ($today.AddDays(-3).ToOADate())
                Set-CellValue $worksheet "M972" ($today.AddDays(-1).ToOADate())
                Set-CellValue $worksheet "H972" "未着手"
                Set-CellValue $worksheet "I972" 0

                Set-CellValue $worksheet "C973" "This Week Parent LV1"
                Set-CellValue $worksheet "D974" "This Week Child LV2"
                Set-CellValue $worksheet "L974" ($today.ToOADate())
                Set-CellValue $worksheet "M974" ($today.ToOADate())
                Set-CellValue $worksheet "H974" "未着手"
                Set-CellValue $worksheet "I974" 0

                Set-CellValue $worksheet "C975" "Standalone LV1 Delayed"
                Set-CellValue $worksheet "L975" ($today.AddDays(-4).ToOADate())
                Set-CellValue $worksheet "M975" ($today.AddDays(-2).ToOADate())
                Set-CellValue $worksheet "H975" "未着手"
                Set-CellValue $worksheet "I975" 0
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }

            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Assert-True ((Get-CellText $worksheet "C972") -eq "!!") "${Edition}: delayed child did not get !!"
            Assert-True ((Get-CellText $worksheet "C971") -eq "!!") "${Edition}: LV2 parent did not inherit !!"
            Assert-True ((Get-CellText $worksheet "C970") -eq "Alert Parent LV1") "${Edition}: LV1 parent should not show !! prefix"
            Assert-True ((Get-CellText $worksheet "C974") -eq "!") "${Edition}: this-week child did not get !"
            Assert-True ((Get-CellText $worksheet "C973") -eq "This Week Parent LV1") "${Edition}: LV1 parent should not show ! prefix"
            Assert-True ((Get-CellText $worksheet "C975") -eq "Standalone LV1 Delayed") "${Edition}: standalone LV1 should not show !! prefix"
        }

        Invoke-TestCase "TC-20" $Edition "親予定日の自動入力と手入力保持" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 976 980
            $base = [datetime]"2026-06-01"
            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                Set-CellValue $worksheet "C976" "Parent Date Rollup LV1"
                Set-CellValue $worksheet "D977" "Parent Date Rollup LV2"
                Set-CellValue $worksheet "E978" "Parent Date Child A"
                Set-CellValue $worksheet "E979" "Parent Date Child B"
                Set-CellValue $worksheet "L978" ($base.AddDays(2).ToOADate())
                Set-CellValue $worksheet "M978" ($base.AddDays(5).ToOADate())
                Set-CellValue $worksheet "L979" ($base.AddDays(4).ToOADate())
                Set-CellValue $worksheet "M979" ($base.AddDays(8).ToOADate())
                Set-CellValue $worksheet "H978" "未着手"
                Set-CellValue $worksheet "H979" "未着手"
                Set-CellValue $worksheet "I978" 0
                Set-CellValue $worksheet "I979" 0
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }

            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Assert-Near ([double](Get-CellValue $worksheet "L977")) ([double]$base.AddDays(2).ToOADate()) 0.001 "${Edition}: LV2 parent start was not auto-filled"
            Assert-Near ([double](Get-CellValue $worksheet "M977")) ([double]$base.AddDays(8).ToOADate()) 0.001 "${Edition}: LV2 parent end was not auto-filled"
            Assert-Near ([double](Get-CellValue $worksheet "L976")) ([double]$base.AddDays(2).ToOADate()) 0.001 "${Edition}: LV1 parent start was not auto-filled"
            Assert-Near ([double](Get-CellValue $worksheet "M976")) ([double]$base.AddDays(8).ToOADate()) 0.001 "${Edition}: LV1 parent end was not auto-filled"

            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                Set-CellValue $worksheet "L978" ($base.AddDays(1).ToOADate())
                Set-CellValue $worksheet "M979" ($base.AddDays(9).ToOADate())
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }

            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Assert-Near ([double](Get-CellValue $worksheet "L977")) ([double]$base.AddDays(1).ToOADate()) 0.001 "${Edition}: auto-filled start did not follow child changes"
            Assert-Near ([double](Get-CellValue $worksheet "M977")) ([double]$base.AddDays(9).ToOADate()) 0.001 "${Edition}: auto-filled end did not follow child changes"

            Set-CellValue $worksheet "L977" ($base.AddDays(-3).ToOADate())
            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                Set-CellValue $worksheet "L978" ($base.AddDays(6).ToOADate())
                Set-CellValue $worksheet "L979" ($base.AddDays(4).ToOADate())
                Set-CellValue $worksheet "M979" ($base.AddDays(10).ToOADate())
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }

            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Assert-Near ([double](Get-CellValue $worksheet "L977")) ([double]$base.AddDays(-3).ToOADate()) 0.001 "${Edition}: manual parent start was overwritten"
            Assert-Near ([double](Get-CellValue $worksheet "M977")) ([double]$base.AddDays(10).ToOADate()) 0.001 "${Edition}: non-manual parent end did not keep following children"

            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                $worksheet.Range("L977").ClearContents() | Out-Null
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }

            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            Assert-Near ([double](Get-CellValue $worksheet "L977")) ([double]$base.AddDays(4).ToOADate()) 0.001 "${Edition}: cleared manual parent start did not return to auto"
        }

        Invoke-TestCase "TC-21" $Edition "親ステータス遅延判定と連続更新復帰" {
            Set-BulkEditState $Excel $workbook $settings $false
            Clear-TestRows $Excel $worksheet 982 988
            $today = (Get-Date).Date
            $previousEvents = $Excel.EnableEvents
            try {
                $Excel.EnableEvents = $false
                Set-CellValue $worksheet "C982" "Visible Future Root LV1"
                Set-CellValue $worksheet "D983" "Visible Future Parent LV2"
                Set-CellValue $worksheet "E984" "Past Child Under Future Parent"
                Set-CellValue $worksheet "L983" ($today.AddDays(2).ToOADate())
                Set-CellValue $worksheet "M983" ($today.AddDays(5).ToOADate())
                Set-CellValue $worksheet "L984" ($today.AddDays(-5).ToOADate())
                Set-CellValue $worksheet "M984" ($today.AddDays(-2).ToOADate())
                Set-CellValue $worksheet "H984" "未着手"
                Set-CellValue $worksheet "I984" 0

                Set-CellValue $worksheet "C986" "Auto Overdue Root LV1"
                Set-CellValue $worksheet "D987" "Auto Overdue Parent LV2"
                Set-CellValue $worksheet "E988" "Auto Overdue Child"
                Set-CellValue $worksheet "L988" ($today.AddDays(-5).ToOADate())
                Set-CellValue $worksheet "M988" ($today.AddDays(-2).ToOADate())
                Set-CellValue $worksheet "H988" "未着手"
                Set-CellValue $worksheet "I988" 0
            }
            finally {
                $Excel.EnableEvents = $previousEvents
            }

            $previousCalculation = $Excel.Calculation
            $previousScreenUpdating = [bool]$Excel.ScreenUpdating
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))
            $Excel.Run(("'{0}'!RefreshInazumaGantt" -f $workbook.Name))

            Assert-True ((Get-CellText $worksheet "C982") -eq "Visible Future Root LV1") "${Edition}: LV1 future parent should not show alert marker"
            Assert-True ((Get-CellText $worksheet "H983") -ne "遅延") "${Edition}: future visible parent end was marked delayed"
            Assert-True ((Get-CellText $worksheet "H987") -eq "遅延") "${Edition}: overdue visible parent end was not marked delayed"
            Assert-True ($Excel.EnableEvents) "${Edition}: EnableEvents was not restored after repeated refresh"
            Assert-True ([int]$Excel.Calculation -eq [int]$previousCalculation) "${Edition}: Calculation mode was not restored after repeated refresh"
            Assert-True ([bool]$Excel.ScreenUpdating -eq $previousScreenUpdating) "${Edition}: ScreenUpdating was not restored after repeated refresh"
        }

    }
    finally {
        if ($null -ne $workbook) {
            try {
                Set-AutomationMode $Excel $workbook $settings $false
            }
            catch {
            }
            try {
                $workbook.Close($false) | Out-Null
            }
            catch {
            }
            try {
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            catch {
            }
        }
    }
}

function Restore-WorkbookFromPayload {
    param(
        [string]$DistributionRoot,
        [string]$DestinationExcelDir
    )

    New-Item -ItemType Directory -Path $DestinationExcelDir -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $DistributionRoot "excel\WorkbookPayload.json") -Destination (Join-Path $DestinationExcelDir "WorkbookPayload.json") -Force
    $restoreScript = Join-Path $DistributionRoot "scripts\RestoreWorkbookFromPayload.ps1"
    & pwsh -NoProfile -ExecutionPolicy Bypass -File $restoreScript -PayloadPath (Join-Path $DestinationExcelDir "WorkbookPayload.json") | Out-Null
    return (Get-ChildItem -LiteralPath $DestinationExcelDir -Filter "*.xlsm" -File | Select-Object -First 1).FullName
}

New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($V3WorkbookPath)) {
    $v3Source = Get-LatestWorkbook (Join-Path $ProjectRoot "distribution_v3\excel") "InazumaGantt_v3_*.xlsm"
}
else {
    $v3Source = (Resolve-Path -LiteralPath $V3WorkbookPath).Path
}

if ([string]::IsNullOrWhiteSpace($LiteWorkbookPath)) {
    $liteSource = Get-LatestWorkbook (Join-Path $ProjectRoot "distribution_lite\excel") "InazumaGantt_Lite_*.xlsm"
}
else {
    $liteSource = (Resolve-Path -LiteralPath $LiteWorkbookPath).Path
}

$v3Copy = Join-Path $OutputRoot (Split-Path -Leaf $v3Source)
$liteCopy = Join-Path $OutputRoot (Split-Path -Leaf $liteSource)
Copy-Item -LiteralPath $v3Source -Destination $v3Copy -Force
Copy-Item -LiteralPath $liteSource -Destination $liteCopy -Force

$excel = $null
$excelVersion = ""
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excelVersion = [string]$excel.Version

    Invoke-WorkbookAcceptance $excel $v3Copy "v3" "InazumaGantt_v3" $true $OutputRoot
    Invoke-WorkbookAcceptance $excel $liteCopy "Lite" "InazumaGantt_Lite" $false $OutputRoot

    if (-not $SkipRestored) {
        Invoke-TestCase "TC-16" "v3" "配布 bundle 復元後の製品近似確認" {
            $restoredPath = Restore-WorkbookFromPayload (Join-Path $ProjectRoot "distribution_v3") (Join-Path $OutputRoot "restored_v3\excel")
            Invoke-WorkbookAcceptance $excel $restoredPath "v3-restored" "InazumaGantt_v3" $true $OutputRoot
        }

        Invoke-TestCase "TC-16" "Lite" "配布 bundle 復元後の製品近似確認" {
            $restoredPath = Restore-WorkbookFromPayload (Join-Path $ProjectRoot "distribution_lite") (Join-Path $OutputRoot "restored_lite\excel")
            Invoke-WorkbookAcceptance $excel $restoredPath "Lite-restored" "InazumaGantt_Lite" $false $OutputRoot
        }
    }
}
finally {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$summary = [ordered]@{
    generatedAt = (Get-Date -Format "yyyy-MM-dd HH:mm:ss K")
    projectRoot = $ProjectRoot
    outputRoot = $OutputRoot
    excelVersion = $excelVersion
    sourceWorkbooks = @($v3Source, $liteSource)
    total = $results.Count
    ok = @($results | Where-Object { $_.status -eq "OK" }).Count
    ng = @($results | Where-Object { $_.status -eq "NG" }).Count
    manual = @($results | Where-Object { $_.status -eq "MANUAL" }).Count
    results = $results
}

$jsonPath = Join-Path $OutputRoot "acceptance-results.json"
$mdPath = Join-Path $OutputRoot "acceptance-results.md"
$summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

$lines = New-Object System.Collections.Generic.List[string]
$lines.Add("# InazumaGantt acceptance test results")
$lines.Add("")
$lines.Add(("- generatedAt: {0}" -f $summary.generatedAt))
$lines.Add(("- excelVersion: {0}" -f $summary.excelVersion))
$lines.Add(("- total: {0}" -f $summary.total))
$lines.Add(("- ok: {0}" -f $summary.ok))
$lines.Add(("- ng: {0}" -f $summary.ng))
$lines.Add(("- manual: {0}" -f $summary.manual))
$lines.Add("")
$lines.Add("| ID | Edition | Name | Status | Message |")
$lines.Add("|---|---|---|---|---|")
foreach ($result in $results) {
    $message = ([string]$result.message).Replace("|", "\|").Replace("`r", " ").Replace("`n", " ")
    $lines.Add(("| {0} | {1} | {2} | {3} | {4} |" -f $result.id, $result.edition, $result.name, $result.status, $message))
}
[System.IO.File]::WriteAllLines($mdPath, $lines, $utf8NoBom)

if (-not $KeepWorkbooks) {
    Get-ChildItem -LiteralPath $OutputRoot -Filter "*.xlsm" -File -Recurse -ErrorAction SilentlyContinue | Remove-Item -Force
}

Write-Host ("Acceptance tests complete. OK={0}, NG={1}" -f $summary.ok, $summary.ng)
Write-Host "Result JSON: $jsonPath"
Write-Host "Result MD: $mdPath"
if ($summary.ng -gt 0) {
    exit 1
}
