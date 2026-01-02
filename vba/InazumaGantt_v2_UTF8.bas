Attribute VB_Name = "InazumaGantt_v2"
Option Explicit

' ==========================================
'  繧､繝翫ぜ繝槭ぎ繝ｳ繝医メ繝｣繝ｼ繝・- 險ｭ螳壹お繝ｪ繧｢
' ==========================================
' 繝ｬ繧､繧｢繧ｦ繝・
' A: LV(髫主ｱ､) | B: No. | C: TASK | D-F: (繧ｿ繧ｹ繧ｯ逕ｨ繧ｹ繝壹・繧ｹ)
' G: 繧ｿ繧ｹ繧ｯ縺ｮ隧ｳ邏ｰ | H: 迥ｶ豕・| I: 騾ｲ謐礼紫 | J: 諡・ｽ・
' K: 髢句ｧ倶ｺ亥ｮ・| L: 螳御ｺ・ｺ亥ｮ・| M: 髢句ｧ句ｮ溽ｸｾ | N: 螳御ｺ・ｮ溽ｸｾ
' O莉･髯・ 繧ｬ繝ｳ繝医メ繝｣繝ｼ繝磯伜沺 (譌･莉・

Public Const COL_HIERARCHY As String = "A"   ' LV(髫主ｱ､)
Public Const COL_NO As String = "B"          ' No.
Public Const COL_TASK As String = "C"        ' TASK
' D-F蛻励・繧ｿ繧ｹ繧ｯ逕ｨ縺ｮ繧ｹ繝壹・繧ｹ・亥ｹ・ｺ・ａ・・
Public Const COL_TASK_DETAIL As String = "G" ' 繧ｿ繧ｹ繧ｯ縺ｮ隧ｳ邏ｰ
Public Const COL_STATUS As String = "H"      ' 迥ｶ豕・
Public Const COL_PROGRESS As String = "I"    ' 騾ｲ謐礼紫
Public Const COL_ASSIGNEE As String = "J"    ' 諡・ｽ・
Public Const COL_START_PLAN As String = "K"  ' 髢句ｧ倶ｺ亥ｮ・
Public Const COL_END_PLAN As String = "L"    ' 螳御ｺ・ｺ亥ｮ・
Public Const COL_START_ACTUAL As String = "M" ' 髢句ｧ句ｮ溽ｸｾ
Public Const COL_END_ACTUAL As String = "N"  ' 螳御ｺ・ｮ溽ｸｾ

Public Const COL_GANTT_START As String = "O"  ' 繧ｬ繝ｳ繝医メ繝｣繝ｼ繝医・髢句ｧ句・
Public Const ROW_TITLE As Long = 1            ' 繧ｿ繧､繝医Ν陦・
Public Const ROW_WEEK_HEADER As Long = 6      ' 騾ｱ繝倥ャ繝繝ｼ陦・
Public Const ROW_DATE_HEADER As Long = 7      ' 譌･莉倩｡鯉ｼ医ぎ繝ｳ繝茨ｼ・
Public Const ROW_HEADER As Long = 8           ' 譖懈律陦鯉ｼ医ぎ繝ｳ繝茨ｼ・ 鬆・岼繝倥ャ繝繝ｼ陦鯉ｼ・-N蛻暦ｼ・
Public Const ROW_DATA_START As Long = 9       ' 繝・・繧ｿ髢句ｧ玖｡・
Public Const GANTT_DAYS As Long = 120         ' 繧ｬ繝ｳ繝医メ繝｣繝ｼ繝医・譌･謨ｰ
Public Const DATA_ROWS_DEFAULT As Long = 200  ' 蛻晄悄蜈･蜉帷ｯ・峇縺ｮ陦梧焚
Public Const HOLIDAY_SHEET_NAME As String = "逾晄律繝槭せ繧ｿ"
Public Const GUIDE_SHEET_NAME As String = "InazumaGantt_隱ｬ譏・
Public Const MAIN_SHEET_NAME As String = "InazumaGantt_v2"
Public Const GUIDE_LEGEND_START_CELL As String = "E1"
Public Const CELL_PROJECT_START As String = "K3"
Public Const CELL_DISPLAY_WEEK As String = "K4"
Public Const CELL_TODAY As String = "M3"

' 濶ｲ險ｭ螳・
Public Const COLOR_PLAN As Long = 230& + 230& * 256& + 230& * 65536&
Public Const COLOR_PROGRESS As Long = 31& + 78& * 256& + 121& * 65536&
Public Const COLOR_HOLIDAY As Long = 242& + 242& * 256& + 242& * 65536&
Public Const COLOR_ROW_BAND As Long = 248& + 248& * 256& + 248& * 65536&
Public Const COLOR_ACTUAL As Long = 0& + 176& * 256& + 80& * 65536&
Public Const COLOR_TODAY As Long = 255& + 0& * 256& + 0& * 65536&
Public Const COLOR_WARN As Long = 255& + 242& * 256& + 204& * 65536&
Public Const COLOR_ERROR As Long = 255& + 199& * 256& + 206& * 65536&
Public Const COLOR_INAZUMA As Long = 255& + 165& * 256& + 0& * 65536&
Public Const COLOR_HEADER_BG As Long = 68& + 114& * 256& + 196& * 65536&
Public Const COLOR_GANTT_HEADER As Long = 128& + 128& * 256& + 128& * 65536&
Public Const TODAY_LINE_WEIGHT As Double = 2
Public Const ACTUAL_LINE_WEIGHT As Double = 4

' ==========================================
'  蛻晄悄繧ｻ繝・ヨ繧｢繝・・ (繝倥ャ繝繝ｼ菴懈・・・嶌蠑剰ｨｭ螳・
' ==========================================
Sub SetupInazumaGantt()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.Name <> MAIN_SHEET_NAME Then
        On Error Resume Next
        ws.Name = MAIN_SHEET_NAME
        If Err.Number <> 0 Then
            MsgBox "繧ｷ繝ｼ繝亥錐繧・'" & MAIN_SHEET_NAME & "' 縺ｫ螟画峩縺ｧ縺阪∪縺帙ｓ縺ｧ縺励◆縲・ & vbCrLf & "譌｢縺ｫ蜷悟錐縺ｮ繧ｷ繝ｼ繝医′蟄伜惠縺吶ｋ蜿ｯ閭ｽ諤ｧ縺後≠繧翫∪縺吶・, vbExclamation
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 繧ｿ繧､繝医Ν繝ｻ諠・ｱ繧ｨ繝ｪ繧｢
    ws.Range("A" & ROW_TITLE).Value = "繧､繝翫ぜ繝槭ぎ繝ｳ繝医メ繝｣繝ｼ繝・
    ws.Range("A" & ROW_TITLE).Font.Bold = True
    ws.Range("A" & ROW_TITLE).Font.Size = 16
    ws.Range("A2").Value = "莨夂､ｾ蜷・
    ws.Range("A3").Value = "繝励Ο繧ｸ繧ｧ繧ｯ繝井ｸｻ莉ｻ"
    ws.Range("J3").Value = "繝励Ο繧ｸ繧ｧ繧ｯ繝医・髢句ｧ・"
    ws.Range("J4").Value = "騾ｱ陦ｨ遉ｺ:"
    ws.Range("L3").Value = "莉頑律:"
    
    ' 繝倥ャ繝繝ｼ險ｭ螳・(ROW_HEADER = 8陦檎岼縺ｫ邨ｱ荳)
    ws.Range(COL_HIERARCHY & ROW_HEADER).Value = "LV"
    ws.Range(COL_NO & ROW_HEADER).Value = "No."
    ' C・曦蛻励・縺吶∋縺ｦTASK蜈･蜉帛・・亥・蜉帑ｽ咲ｽｮ縺ｧ髫主ｱ､縺梧ｱｺ螳夲ｼ・
    ws.Range("C" & ROW_HEADER).Value = "TASK(LV1)"
    ws.Range("D" & ROW_HEADER).Value = "TASK(LV2)"
    ws.Range("E" & ROW_HEADER).Value = "TASK(LV3)"
    ws.Range("F" & ROW_HEADER).Value = "TASK(LV4)"
    ws.Range(COL_TASK_DETAIL & ROW_HEADER).Value = "繧ｿ繧ｹ繧ｯ隧ｳ邏ｰ"
    ws.Range(COL_STATUS & ROW_HEADER).Value = "迥ｶ豕・
    ws.Range(COL_PROGRESS & ROW_HEADER).Value = "騾ｲ謐礼紫"
    ws.Range(COL_ASSIGNEE & ROW_HEADER).Value = "諡・ｽ・
    ws.Range(COL_START_PLAN & ROW_HEADER).Value = "髢句ｧ倶ｺ亥ｮ・
    ws.Range(COL_END_PLAN & ROW_HEADER).Value = "螳御ｺ・ｺ亥ｮ・
    ws.Range(COL_START_ACTUAL & ROW_HEADER).Value = "髢句ｧ句ｮ溽ｸｾ"
    ws.Range(COL_END_ACTUAL & ROW_HEADER).Value = "螳御ｺ・ｮ溽ｸｾ"
    
    ' 繝倥ャ繝繝ｼ陦後・繧ｹ繧ｿ繧､繝ｫ・・陦檎岼縲、・朦蛻暦ｼ・
    With ws.Range("A" & ROW_HEADER & ":N" & ROW_HEADER)
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER_BG
        .Font.Color = RGB(255, 255, 255)
    End With

    EnsureHolidaySheet
    EnsureGuideSheet
    
    ' 譌･莉倬幕蟋区律繧貞・蜉帙＆縺帙ｋ
    Dim startDateInput As Variant
    startDateInput = Application.InputBox("繧ｬ繝ｳ繝医メ繝｣繝ｼ繝医・髢句ｧ区律繧貞・蜉帙＠縺ｦ縺上□縺輔＞ (萓・ 24/12/25)", "髢句ｧ区律險ｭ螳・, Format(Date, "yy/mm/dd"), Type:=2)
    
    If startDateInput = False Then
        startDateInput = Date
    End If
    
    Dim ganttStartDate As Date
    If IsDate(startDateInput) Then
        ganttStartDate = CDate(startDateInput)
    Else
        ganttStartDate = Date
    End If
    
    ws.Range(CELL_PROJECT_START).Value = ganttStartDate
    ws.Range(CELL_PROJECT_START).NumberFormat = "yyyy/mm/dd"
    ws.Range(CELL_DISPLAY_WEEK).Value = 1
    ws.Range(CELL_TODAY).Value = Date
    ws.Range(CELL_TODAY).NumberFormat = "yyyy/mm/dd"
    
    ' 譌･莉伜・縺ｮ逕滓・
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column

    Dim todayDate As Date
    todayDate = Date
    If IsDate(ws.Range(CELL_TODAY).Value) Then
        todayDate = CDate(ws.Range(CELL_TODAY).Value)
    End If
    
    ' 騾ｱ繝ｻ譌･莉倥・譖懈律繝倥ャ繝繝ｼ縺ｮ菴懈・
    Dim weekStartCol As Long
    Dim weekEndCol As Long
    Dim currentDate As Date
    Dim colIndex As Long
    Dim i As Long
    
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        
        ' 7陦檎岼: 譌･莉假ｼ域律縺ｮ縺ｿ・・
        ws.Cells(ROW_DATE_HEADER, colIndex).Value = Day(currentDate)
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Size = 9
        ws.Cells(ROW_DATE_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_DATE_HEADER, colIndex).Interior.Color = COLOR_GANTT_HEADER
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Color = RGB(255, 255, 255)
        
        ' 8陦檎岼: 譖懈律
        ws.Cells(ROW_HEADER, colIndex).Value = Format$(currentDate, "aaa")
        ws.Cells(ROW_HEADER, colIndex).Font.Size = 8
        ws.Cells(ROW_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_HEADER, colIndex).Interior.Color = COLOR_GANTT_HEADER
        ws.Cells(ROW_HEADER, colIndex).Font.Color = RGB(255, 255, 255)
        
        ' 蝨滓律縺ｮ濶ｲ蛻・￠
        If Weekday(currentDate, vbMonday) >= 6 Then
            ws.Cells(ROW_DATE_HEADER, colIndex).Interior.Color = COLOR_HOLIDAY
            ws.Cells(ROW_DATE_HEADER, colIndex).Font.Color = RGB(128, 128, 128)
            ws.Cells(ROW_HEADER, colIndex).Interior.Color = COLOR_HOLIDAY
            ws.Cells(ROW_HEADER, colIndex).Font.Color = RGB(128, 128, 128)
        End If
        
        ' 蛻怜ｹ・ｒ險ｭ螳・
        ws.Columns(colIndex).ColumnWidth = 3
        
        ' 6陦檎岼: 騾ｱ繝倥ャ繝繝ｼ・・譌･蜊倅ｽ搾ｼ・
        If (i - 1) Mod 7 = 0 Then
            weekStartCol = colIndex
            weekEndCol = Application.WorksheetFunction.Min(ganttStartCol + GANTT_DAYS - 1, weekStartCol + 6)
            With ws.Range(ws.Cells(ROW_WEEK_HEADER, weekStartCol), ws.Cells(ROW_WEEK_HEADER, weekEndCol))
                .Merge
                .Value = Format$(currentDate, "yyyy/m/d")
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 9
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Weight = xlThin
            End With
        End If
    Next i
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then
        lastRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    End If

    ApplyGanttBorders ws, lastRow
    DrawWeekSeparators ws, lastRow
    ApplyDataValidationAndFormats ws, lastRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "繧ｻ繝・ヨ繧｢繝・・螳御ｺ・ｼ・ & vbCrLf & "繝・・繧ｿ繧貞・蜉帛ｾ後ヽefreshInazumaGantt 繧貞ｮ溯｡後＠縺ｦ縺上□縺輔＞縲・, vbInformation, "繧､繝翫ぜ繝槭ぎ繝ｳ繝・
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub

' ==========================================
'  逾晄律繝槭せ繧ｿ縺ｮ遒ｺ菫・
' ==========================================
Private Sub EnsureHolidaySheet()
    Dim wsHoliday As Worksheet
    On Error Resume Next
    Set wsHoliday = ThisWorkbook.Worksheets(HOLIDAY_SHEET_NAME)
    On Error GoTo 0
    
    If wsHoliday Is Nothing Then
        Set wsHoliday = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsHoliday.Name = HOLIDAY_SHEET_NAME
        wsHoliday.Range("A1").Value = "逾晄律"
        wsHoliday.Range("A1").Font.Bold = True
        wsHoliday.Columns("A").NumberFormat = "yy/mm/dd"
    End If
End Sub

' ==========================================
'  蜈･蜉幄ｦ丞援縺ｨ譌･莉俶嶌蠑上・驕ｩ逕ｨ
' ==========================================
Private Sub ApplyDataValidationAndFormats(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    ' 騾ｲ謐礼紫縺ｮ繝峨Ο繝・・繝繧ｦ繝ｳ
    With ws.Range(COL_PROGRESS & ROW_DATA_START & ":" & COL_PROGRESS & lastRow)
        .NumberFormat = "0%"
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="0%,10%,20%,30%,40%,50%,60%,70%,80%,90%,100%"
            .InCellDropdown = True
        End With
    End With
    
    ' 迥ｶ豕√・繝峨Ο繝・・繝繧ｦ繝ｳ
    With ws.Range(COL_STATUS & ROW_DATA_START & ":" & COL_STATUS & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="譛ｪ逹謇・騾ｲ陦御ｸｭ,螳御ｺ・菫晉蕗"
        .InCellDropdown = True
    End With
    
    ' 譌･莉伜・縺ｮ譖ｸ蠑・
    ws.Range(COL_START_PLAN & ROW_DATA_START & ":" & COL_END_ACTUAL & lastRow).NumberFormat = "yy/mm/dd"
End Sub

' ==========================================
'  繝・・繧ｿ譛邨り｡後・蜿門ｾ・
' ==========================================
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ROW_HEADER
    
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_TASK_DETAIL).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_START_PLAN).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_END_PLAN).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_START_ACTUAL).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_END_ACTUAL).End(xlUp).Row)
    
    GetLastDataRow = lastRow
End Function

Private Function MaxRow(ByVal a As Long, ByVal b As Long) As Long
    If b > a Then
        MaxRow = b
    Else
        MaxRow = a
    End If
End Function

' ==========================================
'  隱ｬ譏弱す繝ｼ繝医・菴懈・
' ==========================================
Private Sub EnsureGuideSheet()
    Dim wsGuide As Worksheet
    On Error Resume Next
    Set wsGuide = ThisWorkbook.Worksheets(GUIDE_SHEET_NAME)
    On Error GoTo 0
    
    If wsGuide Is Nothing Then
        Set wsGuide = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsGuide.Name = GUIDE_SHEET_NAME
    Else
        wsGuide.Cells.Clear
    End If
    
    wsGuide.Cells(1, 1).Value = "InazumaGantt 隱ｬ譏・
    wsGuide.Cells(1, 1).Font.Bold = True
    wsGuide.Cells(3, 1).Value = "1) SetupInazumaGantt 繧貞ｮ溯｡後＠縺ｦ蛻晄悄險ｭ螳・
    wsGuide.Cells(4, 1).Value = "2) 繧ｿ繧ｹ繧ｯ繧貞・蜉幢ｼ・-F蛻暦ｼ・
    wsGuide.Cells(5, 1).Value = "3) RefreshInazumaGantt 繧貞ｮ溯｡後＠縺ｦ繧ｬ繝ｳ繝域峩譁ｰ"
    wsGuide.Columns(1).ColumnWidth = 50
End Sub

' ==========================================
'  繧ｬ繝ｳ繝亥・菴薙・鄂ｫ邱・
' ==========================================
Private Sub ApplyGanttBorders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttEndCol As Long
    ganttEndCol = ganttStartCol + GANTT_DAYS - 1
    
    Dim borderRange As Range
    Set borderRange = ws.Range(ws.Cells(ROW_DATE_HEADER, 1), ws.Cells(lastRow, ganttEndCol))
    
    With borderRange.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(217, 217, 217)
    End With
    With borderRange.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(217, 217, 217)
    End With
    With borderRange.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(217, 217, 217)
    End With
    With borderRange.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(217, 217, 217)
    End With
    With borderRange.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(217, 217, 217)
    End With
    With borderRange.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(217, 217, 217)
    End With
End Sub

' ==========================================
'  騾ｱ縺ｮ蛹ｺ蛻・ｊ邱・
' ==========================================
Private Sub DrawWeekSeparators(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim colIndex As Long
    Dim weekRange As Range
    
    For colIndex = ganttStartCol To ganttStartCol + GANTT_DAYS - 1 Step 7
        Set weekRange = ws.Range(ws.Cells(ROW_WEEK_HEADER, colIndex), ws.Cells(lastRow, colIndex))
        With weekRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .Color = RGB(191, 191, 191)
        End With
    Next colIndex
End Sub

' ==========================================
'  繧ｬ繝ｳ繝医ヰ繝ｼ謠冗判
' ==========================================
Sub DrawGanttBars()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttRange As Range
    Set ganttRange = ws.Range(ws.Cells(ROW_DATA_START, ganttStartCol), ws.Cells(lastRow, ganttStartCol + GANTT_DAYS - 1))
    
    ' 譌｢蟄倥・譖ｸ蠑上ｒ繧ｯ繝ｪ繧｢
    ganttRange.Interior.ColorIndex = xlNone
    ganttRange.FormatConditions.Delete
    
    ' 譚｡莉ｶ莉倥″譖ｸ蠑・ 蝨滓律繝ｻ逾晄律
    Dim cfHoliday As FormatCondition
    Set cfHoliday = ganttRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=OR(WEEKDAY(" & COL_GANTT_START & "$" & ROW_DATE_HEADER & ",2)>=6,COUNTIF('" & HOLIDAY_SHEET_NAME & "'!$A:$A," & COL_GANTT_START & "$" & ROW_DATE_HEADER & ")>0)")
    cfHoliday.Interior.Color = COLOR_HOLIDAY
    cfHoliday.StopIfTrue = False
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DrawGanttBars 繧ｨ繝ｩ繝ｼ: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub

' ==========================================
'  蜈ｨ謠冗判螳溯｡・
' ==========================================
Sub RefreshInazumaGantt()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    ApplyGanttBorders ws, lastRow
    DrawWeekSeparators ws, lastRow
    ApplyDataValidationAndFormats ws, lastRow
    
    Call DrawGanttBars
    
    MsgBox "繧､繝翫ぜ繝槭ぎ繝ｳ繝域峩譁ｰ螳御ｺ・ｼ・, vbInformation, "繧､繝翫ぜ繝槭ぎ繝ｳ繝・
    Exit Sub
    
ErrorHandler:
    MsgBox "譖ｴ譁ｰ荳ｭ縺ｫ繧ｨ繝ｩ繝ｼ縺檎匱逕溘＠縺ｾ縺励◆: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub

' ==========================================
'  繧ｿ繧ｹ繧ｯ蛻励・髢句ｧ倶ｽ咲ｽｮ繧貞叙蠕暦ｼ磯嚴螻､繝ｬ繝吶Ν縺九ｉ・・
' ==========================================
Public Function GetTaskColumnByLevel(ByVal level As Long) As String
    Select Case level
        Case 1
            GetTaskColumnByLevel = "C"
        Case 2
            GetTaskColumnByLevel = "D"
        Case 3
            GetTaskColumnByLevel = "E"
        Case 4
            GetTaskColumnByLevel = "F"
        Case Else
            GetTaskColumnByLevel = "C"
    End Select
End Function

' ==========================================
'  繧ｿ繧ｹ繧ｯ蜈･蜉帛・縺九ｉ髫主ｱ､繧定・蜍募愛螳・
' ==========================================
Public Sub AutoDetectTaskLevel(Optional ByVal targetRow As Long = 0)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim startRow As Long, endRow As Long
    
    If targetRow > 0 Then
        If targetRow < ROW_DATA_START Then Exit Sub
        startRow = targetRow
        endRow = targetRow
    Else
        startRow = ROW_DATA_START
        endRow = GetLastDataRow(ws)
        If endRow < ROW_DATA_START Then endRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    End If
    
    Application.EnableEvents = False
    
    Dim r As Long
    Dim taskLevel As Long
    
    For r = startRow To endRow
        taskLevel = 0
        
        If Trim$(CStr(ws.Cells(r, "F").Value)) <> "" Then
            taskLevel = 4
        ElseIf Trim$(CStr(ws.Cells(r, "E").Value)) <> "" Then
            taskLevel = 3
        ElseIf Trim$(CStr(ws.Cells(r, "D").Value)) <> "" Then
            taskLevel = 2
        ElseIf Trim$(CStr(ws.Cells(r, "C").Value)) <> "" Then
            taskLevel = 1
        End If
        
        If taskLevel > 0 Then
            ws.Cells(r, COL_HIERARCHY).Value = taskLevel
        Else
            ws.Cells(r, COL_HIERARCHY).ClearContents
        End If
    Next r
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "髫主ｱ､閾ｪ蜍募愛螳壹お繝ｩ繝ｼ: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub
