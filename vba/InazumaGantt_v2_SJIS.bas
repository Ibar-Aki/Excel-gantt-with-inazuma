Attribute VB_Name = "InazumaGantt_v2"
Option Explicit

' ==========================================
'  イナズマガントチャーチE- 設定エリア
' ==========================================
' レイアウチE
' A: LV(階層) | B: No. | C: TASK | D-F: (タスク用スペ�Eス)
' G: タスクの詳細 | H: 状況E| I: 進捗率 | J: 拁E��E
' K: 開始予宁E| L: 完亁E��宁E| M: 開始実績 | N: 完亁E��績
' O以陁E ガントチャート領域 (日仁E

Public Const COL_HIERARCHY As String = "A"   ' LV(階層)
Public Const COL_NO As String = "B"          ' No.
Public Const COL_TASK As String = "C"        ' TASK
' D-F列�Eタスク用のスペ�Eス�E�幁E��E���E�E
Public Const COL_TASK_DETAIL As String = "G" ' タスクの詳細
Public Const COL_STATUS As String = "H"      ' 状況E
Public Const COL_PROGRESS As String = "I"    ' 進捗率
Public Const COL_ASSIGNEE As String = "J"    ' 拁E��E
Public Const COL_START_PLAN As String = "K"  ' 開始予宁E
Public Const COL_END_PLAN As String = "L"    ' 完亁E��宁E
Public Const COL_START_ACTUAL As String = "M" ' 開始実績
Public Const COL_END_ACTUAL As String = "N"  ' 完亁E��績

Public Const COL_GANTT_START As String = "O"  ' ガントチャート�E開始�E
Public Const ROW_TITLE As Long = 1            ' タイトル衁E
Public Const ROW_WEEK_HEADER As Long = 6      ' 週ヘッダー衁E
Public Const ROW_DATE_HEADER As Long = 7      ' 日付行（ガント！E
Public Const ROW_HEADER As Long = 8           ' 曜日行（ガント！E 頁E��ヘッダー行！E-N列！E
Public Const ROW_DATA_START As Long = 9       ' チE�Eタ開始衁E
Public Const GANTT_DAYS As Long = 120         ' ガントチャート�E日数
Public Const DATA_ROWS_DEFAULT As Long = 200  ' 初期入力篁E��の行数
Public Const HOLIDAY_SHEET_NAME As String = "祝日マスタ"
Public Const GUIDE_SHEET_NAME As String = "InazumaGantt_説昁E
Public Const MAIN_SHEET_NAME As String = "InazumaGantt_v2"
Public Const GUIDE_LEGEND_START_CELL As String = "E1"
Public Const CELL_PROJECT_START As String = "K3"
Public Const CELL_DISPLAY_WEEK As String = "K4"
Public Const CELL_TODAY As String = "M3"

' 色設宁E
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
'  初期セチE��アチE�E (ヘッダー作�E�E�E��式設宁E
' ==========================================
Sub SetupInazumaGantt()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.Name <> MAIN_SHEET_NAME Then
        On Error Resume Next
        ws.Name = MAIN_SHEET_NAME
        If Err.Number <> 0 Then
            MsgBox "シート名めE'" & MAIN_SHEET_NAME & "' に変更できませんでした、E & vbCrLf & "既に同名のシートが存在する可能性があります、E, vbExclamation
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' タイトル・惁E��エリア
    ws.Range("A" & ROW_TITLE).Value = "イナズマガントチャーチE
    ws.Range("A" & ROW_TITLE).Font.Bold = True
    ws.Range("A" & ROW_TITLE).Font.Size = 16
    ws.Range("A2").Value = "会社吁E
    ws.Range("A3").Value = "プロジェクト主任"
    ws.Range("J3").Value = "プロジェクト�E開姁E"
    ws.Range("J4").Value = "週表示:"
    ws.Range("L3").Value = "今日:"
    
    ' ヘッダー設宁E(ROW_HEADER = 8行目に統一)
    ws.Range(COL_HIERARCHY & ROW_HEADER).Value = "LV"
    ws.Range(COL_NO & ROW_HEADER).Value = "No."
    ' C�E�F列�EすべてTASK入力�E�E��E力位置で階層が決定！E
    ws.Range("C" & ROW_HEADER).Value = "TASK(LV1)"
    ws.Range("D" & ROW_HEADER).Value = "TASK(LV2)"
    ws.Range("E" & ROW_HEADER).Value = "TASK(LV3)"
    ws.Range("F" & ROW_HEADER).Value = "TASK(LV4)"
    ws.Range(COL_TASK_DETAIL & ROW_HEADER).Value = "タスク詳細"
    ws.Range(COL_STATUS & ROW_HEADER).Value = "状況E
    ws.Range(COL_PROGRESS & ROW_HEADER).Value = "進捗率"
    ws.Range(COL_ASSIGNEE & ROW_HEADER).Value = "拁E��E
    ws.Range(COL_START_PLAN & ROW_HEADER).Value = "開始予宁E
    ws.Range(COL_END_PLAN & ROW_HEADER).Value = "完亁E��宁E
    ws.Range(COL_START_ACTUAL & ROW_HEADER).Value = "開始実績"
    ws.Range(COL_END_ACTUAL & ROW_HEADER).Value = "完亁E��績"
    
    ' ヘッダー行�Eスタイル�E�E行目、A�E�N列！E
    With ws.Range("A" & ROW_HEADER & ":N" & ROW_HEADER)
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER_BG
        .Font.Color = RGB(255, 255, 255)
    End With

    EnsureHolidaySheet
    EnsureGuideSheet
    
    ' 日付開始日を�E力させる
    Dim startDateInput As Variant
    startDateInput = Application.InputBox("ガントチャート�E開始日を�E力してください (侁E 24/12/25)", "開始日設宁E, Format(Date, "yy/mm/dd"), Type:=2)
    
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
    
    ' 日付�Eの生�E
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column

    Dim todayDate As Date
    todayDate = Date
    If IsDate(ws.Range(CELL_TODAY).Value) Then
        todayDate = CDate(ws.Range(CELL_TODAY).Value)
    End If
    
    ' 週・日付�E曜日ヘッダーの作�E
    Dim weekStartCol As Long
    Dim weekEndCol As Long
    Dim currentDate As Date
    Dim colIndex As Long
    Dim i As Long
    
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        
        ' 7行目: 日付（日のみ�E�E
        ws.Cells(ROW_DATE_HEADER, colIndex).Value = Day(currentDate)
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Size = 9
        ws.Cells(ROW_DATE_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_DATE_HEADER, colIndex).Interior.Color = COLOR_GANTT_HEADER
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Color = RGB(255, 255, 255)
        
        ' 8行目: 曜日
        ws.Cells(ROW_HEADER, colIndex).Value = Format$(currentDate, "aaa")
        ws.Cells(ROW_HEADER, colIndex).Font.Size = 8
        ws.Cells(ROW_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_HEADER, colIndex).Interior.Color = COLOR_GANTT_HEADER
        ws.Cells(ROW_HEADER, colIndex).Font.Color = RGB(255, 255, 255)
        
        ' 土日の色刁E��
        If Weekday(currentDate, vbMonday) >= 6 Then
            ws.Cells(ROW_DATE_HEADER, colIndex).Interior.Color = COLOR_HOLIDAY
            ws.Cells(ROW_DATE_HEADER, colIndex).Font.Color = RGB(128, 128, 128)
            ws.Cells(ROW_HEADER, colIndex).Interior.Color = COLOR_HOLIDAY
            ws.Cells(ROW_HEADER, colIndex).Font.Color = RGB(128, 128, 128)
        End If
        
        ' 列幁E��設宁E
        ws.Columns(colIndex).ColumnWidth = 3
        
        ' 6行目: 週ヘッダー�E�E日単位！E
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
    
    MsgBox "セチE��アチE�E完亁E��E & vbCrLf & "チE�Eタを�E力後、RefreshInazumaGantt を実行してください、E, vbInformation, "イナズマガンチE
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  祝日マスタの確俁E
' ==========================================
Private Sub EnsureHolidaySheet()
    Dim wsHoliday As Worksheet
    On Error Resume Next
    Set wsHoliday = ThisWorkbook.Worksheets(HOLIDAY_SHEET_NAME)
    On Error GoTo 0
    
    If wsHoliday Is Nothing Then
        Set wsHoliday = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsHoliday.Name = HOLIDAY_SHEET_NAME
        wsHoliday.Range("A1").Value = "祝日"
        wsHoliday.Range("A1").Font.Bold = True
        wsHoliday.Columns("A").NumberFormat = "yy/mm/dd"
    End If
End Sub

' ==========================================
'  入力規則と日付書式�E適用
' ==========================================
Private Sub ApplyDataValidationAndFormats(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    ' 進捗率のドロチE�Eダウン
    With ws.Range(COL_PROGRESS & ROW_DATA_START & ":" & COL_PROGRESS & lastRow)
        .NumberFormat = "0%"
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="0%,10%,20%,30%,40%,50%,60%,70%,80%,90%,100%"
            .InCellDropdown = True
        End With
    End With
    
    ' 状況�EドロチE�Eダウン
    With ws.Range(COL_STATUS & ROW_DATA_START & ":" & COL_STATUS & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="未着扁E進行中,完亁E保留"
        .InCellDropdown = True
    End With
    
    ' 日付�Eの書弁E
    ws.Range(COL_START_PLAN & ROW_DATA_START & ":" & COL_END_ACTUAL & lastRow).NumberFormat = "yy/mm/dd"
End Sub

' ==========================================
'  チE�Eタ最終行�E取征E
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
'  説明シート�E作�E
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
    
    wsGuide.Cells(1, 1).Value = "InazumaGantt 説昁E
    wsGuide.Cells(1, 1).Font.Bold = True
    wsGuide.Cells(3, 1).Value = "1) SetupInazumaGantt を実行して初期設宁E
    wsGuide.Cells(4, 1).Value = "2) タスクを�E力！E-F列！E
    wsGuide.Cells(5, 1).Value = "3) RefreshInazumaGantt を実行してガント更新"
    wsGuide.Columns(1).ColumnWidth = 50
End Sub

' ==========================================
'  ガント�E体�E罫緁E
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
'  週の区刁E��緁E
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
'  ガントバー描画
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
    
    ' 既存�E書式をクリア
    ganttRange.Interior.ColorIndex = xlNone
    ganttRange.FormatConditions.Delete
    
    ' 条件付き書弁E 土日・祝日
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
    MsgBox "DrawGanttBars エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  全描画実衁E
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
    
    MsgBox "イナズマガント更新完亁E��E, vbInformation, "イナズマガンチE
    Exit Sub
    
ErrorHandler:
    MsgBox "更新中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  タスク列�E開始位置を取得（階層レベルから�E�E
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
'  タスク入力�Eから階層を�E動判宁E
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
    MsgBox "階層自動判定エラー: " & Err.Description, vbCritical, "エラー"
End Sub
