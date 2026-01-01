Attribute VB_Name = "InazumaGantt_v2"
Option Explicit

' ==========================================
'  イナズマガントチャート - 設定エリア
' ==========================================
' レイアウト:
' A: LV(階層) | B: No. | C: TASK | D-F: (タスク用スペース)
' G: タスクの詳細 | H: 状況 | I: 進捗率 | J: 担当
' K: 開始予定 | L: 完了予定 | M: 開始実績 | N: 完了実績
' O以降: ガントチャート領域 (日付)

Public Const COL_HIERARCHY As String = "A"   ' LV(階層)
Public Const COL_NO As String = "B"          ' No.
Public Const COL_TASK As String = "C"        ' TASK
' D-F列はタスク用のスペース（幅広め）
Public Const COL_TASK_DETAIL As String = "G" ' タスクの詳細
Public Const COL_STATUS As String = "H"      ' 状況
Public Const COL_PROGRESS As String = "I"    ' 進捗率
Public Const COL_ASSIGNEE As String = "J"    ' 担当
Public Const COL_START_PLAN As String = "K"  ' 開始予定
Public Const COL_END_PLAN As String = "L"    ' 完了予定
Public Const COL_START_ACTUAL As String = "M" ' 開始実績
Public Const COL_END_ACTUAL As String = "N"  ' 完了実績

Public Const COL_GANTT_START As String = "O"  ' ガントチャートの開始列
Public Const ROW_TITLE As Long = 1            ' タイトル行
Public Const ROW_WEEK_HEADER As Long = 6      ' 週ヘッダー行
Public Const ROW_HEADER_LABEL As Long = 7     ' 見出し行（No.～完了実績、曜日表示）
Public Const ROW_HEADER As Long = 8           ' 日付行（ガント基準）
Public Const ROW_DATA_START As Long = 9       ' データ開始行
Public Const GANTT_DAYS As Long = 120         ' ガントチャートの日数 (マジックナンバー対策)
Public Const DATA_ROWS_DEFAULT As Long = 200  ' 初期入力範囲の行数
Public Const HOLIDAY_SHEET_NAME As String = "祝日マスタ"
Public Const GUIDE_SHEET_NAME As String = "InazumaGantt_説明"
Public Const MAIN_SHEET_NAME As String = "InazumaGantt_v2"
Public Const GUIDE_LEGEND_START_CELL As String = "E1"
Public Const CELL_PROJECT_START As String = "K3"
Public Const CELL_DISPLAY_WEEK As String = "K4"
Public Const CELL_TODAY As String = "M3"

' 色設定
Public Const COLOR_PLAN As Long = 230& + 230& * 256& + 230& * 65536&
Public Const COLOR_PROGRESS As Long = 31& + 78& * 256& + 121& * 65536&
Public Const COLOR_HOLIDAY As Long = 242& + 242& * 256& + 242& * 65536&
Public Const COLOR_ROW_BAND As Long = 248& + 248& * 256& + 248& * 65536&
Public Const COLOR_ACTUAL As Long = 0& + 176& * 256& + 80& * 65536&
Public Const COLOR_TODAY As Long = 255& + 0& * 256& + 0& * 65536&
Public Const COLOR_WARN As Long = 255& + 242& * 256& + 204& * 65536&
Public Const COLOR_ERROR As Long = 255& + 199& * 256& + 206& * 65536&
Public Const COLOR_INAZUMA As Long = 255& + 165& * 256& + 0& * 65536&  ' RGB(255,165,0) オレンジ
Public Const TODAY_LINE_WEIGHT As Double = 2
Public Const ACTUAL_LINE_WEIGHT As Double = 4

' ==========================================
'  初期セットアップ (ヘッダー作成＆書式設定)
' ==========================================
Sub SetupInazumaGantt()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.Name <> MAIN_SHEET_NAME Then
        On Error Resume Next
        ws.Name = MAIN_SHEET_NAME
        If Err.Number <> 0 Then
            MsgBox "シート名を '" & MAIN_SHEET_NAME & "' に変更できませんでした。" & vbCrLf & "既に同名のシートが存在する可能性があります。", vbExclamation
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' タイトル・情報エリア
    ws.Range("A" & ROW_TITLE).Value = "イナズマガントチャート"
    ws.Range("A" & ROW_TITLE).Font.Bold = True
    ws.Range("A" & ROW_TITLE).Font.Size = 16
    ws.Range("A2").Value = "会社名"
    ws.Range("A3").Value = "プロジェクト主任"
    ws.Range("J3").Value = "プロジェクトの開始:"
    ws.Range("J4").Value = "週表示:"
    ws.Range("L3").Value = "今日:"
    
    ' ヘッダー設定 (ROW_HEADER_LABELに統一)
    ws.Range(COL_HIERARCHY & ROW_HEADER_LABEL).Value = "LV"
    ws.Range(COL_NO & ROW_HEADER_LABEL).Value = "No."
    ' C～F列はすべてTASK入力列（入力位置で階層が決定）
    ws.Range("C" & ROW_HEADER_LABEL).Value = "TASK(LV1)"
    ws.Range("D" & ROW_HEADER_LABEL).Value = "TASK(LV2)"
    ws.Range("E" & ROW_HEADER_LABEL).Value = "TASK(LV3)"
    ws.Range("F" & ROW_HEADER_LABEL).Value = "TASK(LV4)"
    ws.Range(COL_TASK_DETAIL & ROW_HEADER_LABEL).Value = "タスク詳細"
    ws.Range(COL_STATUS & ROW_HEADER_LABEL).Value = "状況"
    ws.Range(COL_PROGRESS & ROW_HEADER_LABEL).Value = "進捗率"
    ws.Range(COL_ASSIGNEE & ROW_HEADER_LABEL).Value = "担当"
    ws.Range(COL_START_PLAN & ROW_HEADER_LABEL).Value = "開始予定"
    ws.Range(COL_END_PLAN & ROW_HEADER_LABEL).Value = "完了予定"
    ws.Range(COL_START_ACTUAL & ROW_HEADER_LABEL).Value = "開始実績"
    ws.Range(COL_END_ACTUAL & ROW_HEADER_LABEL).Value = "完了実績"
    
    ' ヘッダー行のスタイル
    With ws.Range("A" & ROW_HEADER_LABEL & ":N" & ROW_HEADER_LABEL)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    EnsureHolidaySheet
    EnsureGuideSheet
    
    ' 日付開始日を入力させる
    Dim startDateInput As Variant
    startDateInput = Application.InputBox("ガントチャートの開始日を入力してください (例: 24/12/25)", "開始日設定", Format(Date, "yy/mm/dd"), Type:=2)
    
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
    
    ' 日付列の生成 (一括書き込み)
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column

    Dim todayDate As Date
    todayDate = Date
    If IsDate(ws.Range(CELL_TODAY).Value) Then
        todayDate = CDate(ws.Range(CELL_TODAY).Value)
    End If
    
    ' 週・曜日・日付ヘッダーの作成
    Dim weekStartCol As Long
    Dim weekEndCol As Long
    Dim currentDate As Date
    Dim colIndex As Long
    
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        
        
        ' 日付と曜日をROW_HEADERに表示
        ws.Cells(ROW_HEADER, colIndex).Value = Format$(currentDate, "d(aaa)")
        ws.Cells(ROW_HEADER, colIndex).Font.Size = 8
        ws.Cells(ROW_HEADER, colIndex).HorizontalAlignment = xlCenter
        
        ' 列幅を設定
        ws.Columns(colIndex).ColumnWidth = 3
        
        ' 週ヘッダー（7日単位）
        If (i - 1) Mod 7 = 0 Then
            weekStartCol = colIndex
            weekEndCol = Application.WorksheetFunction.Min(ganttStartCol + GANTT_DAYS - 1, weekStartCol + 6)
            With ws.Range(ws.Cells(ROW_WEEK_HEADER, weekStartCol), ws.Cells(ROW_WEEK_HEADER, weekEndCol))
                .Merge
                .Value = Format$(currentDate, "yyyy/m/d")
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                .Font.Size = 9
                .Interior.Color = COLOR_ROW_BAND
            End With
        End If
    Next i
    
    ' 週の区切り線（左罫線を太く）
    Dim r As Long
    ' 週の区切り線は後でデータ行数に合わせて描画する
    
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
    
    MsgBox "セットアップ完了！" & vbCrLf & "データを入力後、RefreshInazumaGantt を実行してください。", vbInformation, "イナズマガント"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  祝日マスタの確保
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
'  入力規則と日付書式の適用
' ==========================================
Private Sub ApplyDataValidationAndFormats(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    ' 進捗率のドロップダウン (数値として保存されるよう書式設定)
    With ws.Range(COL_PROGRESS & ROW_DATA_START & ":" & COL_PROGRESS & lastRow)
        .NumberFormat = "0%"
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="0%,10%,20%,30%,40%,50%,60%,70%,80%,90%,100%"
            .InCellDropdown = True
        End With
    End With
    
    ' 状況のドロップダウン
    With ws.Range(COL_STATUS & ROW_DATA_START & ":" & COL_STATUS & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="未着手,進行中,完了,保留"
        .InCellDropdown = True
    End With
    
    ' 日付列の書式
    ws.Range(COL_START_PLAN & ROW_DATA_START & ":" & COL_END_ACTUAL & lastRow).NumberFormat = "yy/mm/dd"

    ' 進捗未入力の警告
    Dim rngProgress As Range
    Set rngProgress = ws.Range(COL_PROGRESS & ROW_DATA_START & ":" & COL_PROGRESS & lastRow)
    rngProgress.FormatConditions.Delete
    Dim cfWarn As FormatCondition
    Set cfWarn = rngProgress.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND($" & COL_START_PLAN & ROW_DATA_START & "<>"""",$" & COL_END_PLAN & ROW_DATA_START & "<>"""",$" & COL_PROGRESS & ROW_DATA_START & "="""")")
    cfWarn.Interior.Color = COLOR_WARN
    cfWarn.StopIfTrue = False

    ' 開始予定 > 完了予定 の警告
    Dim rngDates As Range
    Set rngDates = ws.Range(COL_START_PLAN & ROW_DATA_START & ":" & COL_END_PLAN & lastRow)
    rngDates.FormatConditions.Delete
    Dim cfInvalid As FormatCondition
    Set cfInvalid = rngDates.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND($" & COL_START_PLAN & ROW_DATA_START & "<>"""",$" & COL_END_PLAN & ROW_DATA_START & "<>"""",$" & COL_START_PLAN & ROW_DATA_START & ">$" & COL_END_PLAN & ROW_DATA_START & ")")
    cfInvalid.Interior.Color = COLOR_ERROR
    cfInvalid.StopIfTrue = False
End Sub

' ==========================================
'  データ最終行の取得
' ==========================================
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ROW_HEADER_LABEL
    
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
'  Vertex42 互換の名前定義（進捗計算用）
' ==========================================
Private Sub EnsureGanttNames(ByVal ws As Worksheet)
    Dim sheetName As String
    sheetName = ws.Name
    Dim sheetRef As String
    sheetRef = "'" & sheetName & "'"
    
    EnsureOrUpdateName "task_start", "=" & sheetRef & "!$" & COL_START_PLAN & "1"
    EnsureOrUpdateName "task_end", "=" & sheetRef & "!$" & COL_END_PLAN & "1"
    
    ' 進捗率は文字列(例: "70%")も吸収
    EnsureOrUpdateName "task_progress", _
        "=IFERROR(IF(ISTEXT(" & sheetRef & "!$" & COL_PROGRESS & "1)," & _
        "VALUE(SUBSTITUTE(" & sheetRef & "!$" & COL_PROGRESS & "1,CHAR(37),""""))/100," & _
        sheetRef & "!$" & COL_PROGRESS & "1),0)"
End Sub

' ==========================================
'  説明シートの作成
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
    
    Dim linesText As String
    linesText = "InazumaGantt 作成者向け説明（初心者向け）"
    linesText = linesText & vbLf & ""
    linesText = linesText & vbLf & "1) 目的"
    linesText = linesText & vbLf & "Excel上でガントチャート＋進捗（オレンジ線）を描画するためのVBAです。"
    linesText = linesText & vbLf & ""
    linesText = linesText & vbLf & "2) 使い方（設定手順）"
    linesText = linesText & vbLf & "① SetupInazumaGantt を実行"
    linesText = linesText & vbLf & "② 開始日を入力"
    linesText = linesText & vbLf & "③ 祝日を使う場合は 祝日マスタ のA列に日付を入力"
    linesText = linesText & vbLf & "④ データ入力後、RefreshInazumaGantt を実行"
    linesText = linesText & vbLf & ""
    linesText = linesText & vbLf & "3) 主なマクロ"
    linesText = linesText & vbLf & "- SetupInazumaGantt : 初期設定"
    linesText = linesText & vbLf & "- RefreshInazumaGantt : 画面更新（バー・線描画）"
    linesText = linesText & vbLf & "- DrawGanttBars : 予定/進捗バーの描画"
    linesText = linesText & vbLf & "- DrawActualBars : 実績バー（緑の太線）"
    linesText = linesText & vbLf & "- DrawInazumaLine : オレンジ線（進捗線）"
    linesText = linesText & vbLf & ""
    linesText = linesText & vbLf & "4) オレンジ線のルール"
    linesText = linesText & vbLf & "- 進行中：進捗率の位置"
    linesText = linesText & vbLf & "- 完了＋予定通り：今日線と同じ位置"
    linesText = linesText & vbLf & "- 完了＋進んでいる：完了予定日に固定"
    linesText = linesText & vbLf & "- 未着手＋遅れ：開始予定日に固定"
    linesText = linesText & vbLf & ""
    linesText = linesText & vbLf & "5) よくある注意"
    linesText = linesText & vbLf & "- 列順を変えると動作しません。"
    linesText = linesText & vbLf & "- 進捗率は数値/パーセントどちらでもOK。"
    linesText = linesText & vbLf & "- basファイルはインポート機能で取り込むのが確実。"
    linesText = linesText & vbLf & ""
    linesText = linesText & vbLf & "6) カスタマイズ"
    linesText = linesText & vbLf & "- 色変更：DrawGanttBars / DrawActualBars / DrawInazumaLine のRGB"
    linesText = linesText & vbLf & "- 日数変更：GANTT_DAYS を変更"

    Dim lines As Variant
    lines = Split(linesText, vbLf)
    
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        wsGuide.Cells(i + 1, 1).Value = lines(i)
    Next i
    
    With wsGuide.Columns(1)
        .ColumnWidth = 70
        .WrapText = True
    End With
    wsGuide.Range("A1").Font.Bold = True
    EnsureLegendOnGuide wsGuide
End Sub

' ==========================================
'  凡例の作成（説明シート）
' ==========================================
Private Sub EnsureLegendOnGuide(ByVal wsGuide As Worksheet)
    Dim startCell As Range
    Set startCell = wsGuide.Range(GUIDE_LEGEND_START_CELL)
    
    Dim labels As Variant
    labels = Array("凡例", "進捗", "予定", "実績", "今日", "土日・祝日", "行の縞模様", "警告(進捗未入力)", "警告(日付逆転)")
    Dim colors As Variant
    colors = Array(vbWhite, COLOR_PROGRESS, COLOR_PLAN, COLOR_ACTUAL, COLOR_TODAY, COLOR_HOLIDAY, COLOR_ROW_BAND, COLOR_WARN, COLOR_ERROR)
    
    Dim i As Long
    For i = LBound(labels) To UBound(labels)
        wsGuide.Cells(startCell.Row + i, startCell.Column).Value = labels(i)
        wsGuide.Cells(startCell.Row + i, startCell.Column + 1).Interior.Color = colors(i)
        wsGuide.Cells(startCell.Row + i, startCell.Column + 1).Value = " "
    Next i
    
    wsGuide.Cells(startCell.Row, startCell.Column).Font.Bold = True
    wsGuide.Columns(startCell.Column).ColumnWidth = 20
    wsGuide.Columns(startCell.Column + 1).ColumnWidth = 4
End Sub

Private Sub EnsureOrUpdateName(ByVal nameText As String, ByVal refersToFormula As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).RefersTo = refersToFormula
    If Err.Number <> 0 Then
        Err.Clear
        ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersToFormula
    End If
    On Error GoTo 0
End Sub

' ==========================================
'  ガントバー描画 (条件付き書式を使用 - 高速化)
' ==========================================
Sub DrawGanttBars()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim stage As String
    stage = "init"
    EnsureHolidaySheet
    EnsureGanttNames ws
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttRange As Range
    stage = "ganttRange"
    Set ganttRange = ws.Range(ws.Cells(ROW_DATA_START, ganttStartCol), ws.Cells(lastRow, ganttStartCol + GANTT_DAYS - 1))
    
    ' 既存の書式をクリア
    stage = "clearFormats"
    ganttRange.Interior.ColorIndex = xlNone
    ganttRange.FormatConditions.Delete
    ' 条件付き書式: 進捗バー（Vertex42 互換）
    stage = "cfProgress"
    Dim cfProgress As FormatCondition
    Set cfProgress = ganttRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(task_start<=" & COL_GANTT_START & "$" & ROW_HEADER & "," & _
                  "ROUNDDOWN((task_end-task_start+1)*task_progress,0)+task_start-1>=" & COL_GANTT_START & "$" & ROW_HEADER & ")")
    cfProgress.Interior.Color = COLOR_PROGRESS
    cfProgress.StopIfTrue = False
    cfProgress.Priority = 1

    ' 条件付き書式: 予定バー（Vertex42 互換）
    stage = "cfPlan"
    Dim cfPlan As FormatCondition
    Set cfPlan = ganttRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND(task_end>=" & COL_GANTT_START & "$" & ROW_HEADER & ",task_start<" & COL_GANTT_START & "$" & ROW_HEADER & "+1)")
    cfPlan.Interior.Color = COLOR_PLAN
    cfPlan.StopIfTrue = False
    cfPlan.Priority = 2

    ' 条件付き書式: 土日・祝日 (灰色)
    stage = "cfHoliday"
    Dim cfHoliday As FormatCondition
    Set cfHoliday = ganttRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=OR(WEEKDAY(" & COL_GANTT_START & "$" & ROW_HEADER & ",2)>=6,COUNTIF('" & HOLIDAY_SHEET_NAME & "'!$A:$A," & COL_GANTT_START & "$" & ROW_HEADER & ")>0)")
    cfHoliday.Interior.Color = COLOR_HOLIDAY
    cfHoliday.StopIfTrue = False
    cfHoliday.Priority = 3

    ' 条件付き書式: 行の縞模様（見やすさ向上）
    stage = "cfRowBand"
    Dim cfRowBand As FormatCondition
    Set cfRowBand = ganttRange.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=MOD(ROW(),2)=0")
    cfRowBand.Interior.Color = COLOR_ROW_BAND
    cfRowBand.StopIfTrue = False
    cfRowBand.Priority = 4
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' MsgBoxは RefreshInazumaGantt からのみ表示
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DrawGanttBars エラー(" & Err.Number & "): " & Err.Description & vbCrLf & "Stage: " & stage, vbCritical, "エラー"
End Sub

' ==========================================
'  実績バー描画（図形の太線）
' ==========================================
Sub DrawActualBars()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim stage As String
    stage = "init"
    
    Application.ScreenUpdating = False
    
    ' 既存の実績バー図形を削除
    Dim shp As Shape
    Dim shapesToDelete As Collection
    Set shapesToDelete = New Collection
    
    stage = "deleteShapes"
    For Each shp In ws.Shapes
        If Left(shp.Name, 9) = "ActualBar" Then
            shapesToDelete.Add shp
        End If
    Next shp
    
    Dim s As Variant
    For Each s In shapesToDelete
        s.Delete
    Next s
    
    Dim lastRow As Long
    stage = "lastRow"
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    Dim ganttStartCol As Long
    stage = "ganttStartCol"
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttStartDate As Date
    stage = "ganttStartDate"
    ganttStartDate = GetGanttStartDate(ws, ganttStartCol)
    
    Dim ganttEndCol As Long
    ganttEndCol = ganttStartCol + GANTT_DAYS - 1
    
    Dim r As Long
    Dim startActual As Variant
    Dim endActual As Variant
    
    stage = "loop"
    For r = ROW_DATA_START To lastRow
        startActual = ws.Cells(r, COL_START_ACTUAL).Value
        endActual = ws.Cells(r, COL_END_ACTUAL).Value
        
        If IsDate(startActual) And IsDate(endActual) Then
            Dim colStart As Long
            Dim colEnd As Long
            colStart = DateToGanttCol(CDate(startActual), ganttStartDate, ganttStartCol)
            colEnd = DateToGanttCol(CDate(endActual), ganttStartDate, ganttStartCol)
            
            If colEnd < colStart Then GoTo NextRow
            If colEnd < ganttStartCol Or colStart > ganttEndCol Then GoTo NextRow
            If colStart < ganttStartCol Then colStart = ganttStartCol
            If colEnd > ganttEndCol Then colEnd = ganttEndCol
            
            Dim startCell As Range
            Dim endCell As Range
            stage = "lineCells"
            Set startCell = ws.Cells(r, colStart)
            Set endCell = ws.Cells(r, colEnd)
            
            Dim y As Double
            Dim x1 As Double
            Dim x2 As Double
            y = startCell.Top + startCell.Height / 2
            x1 = startCell.Left + 1
            x2 = endCell.Left + endCell.Width - 1
            
            Dim actualLine As Shape
            stage = "addLine"
            Set actualLine = ws.Shapes.AddLine(x1, y, x2, y)
            actualLine.Name = "ActualBar" & r
            actualLine.Line.ForeColor.RGB = COLOR_ACTUAL
            actualLine.Line.Weight = ACTUAL_LINE_WEIGHT
        End If
NextRow:
    Next r
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "DrawActualBars エラー(" & Err.Number & "): " & Err.Description & vbCrLf & "Stage: " & stage, vbCritical, "エラー"
End Sub

Private Function GetGanttStartDate(ByVal ws As Worksheet, ByVal ganttStartCol As Long) As Date
    Dim d As Variant
    d = ws.Cells(ROW_HEADER, ganttStartCol).Value
    If IsDate(d) Then
        GetGanttStartDate = CDate(d)
    Else
        GetGanttStartDate = Date
    End If
End Function

Private Function DateToGanttCol(ByVal targetDate As Date, ByVal ganttStartDate As Date, ByVal ganttStartCol As Long) As Long
    DateToGanttCol = ganttStartCol + CLng(targetDate) - CLng(ganttStartDate)
End Function

' ==========================================
'  イナズマ線描画
' ==========================================
Sub DrawInazumaLine()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Dim stage As String
    stage = "init"

    Dim todayDate As Date
    todayDate = Date
    If IsDate(ws.Range(CELL_TODAY).Value) Then
        todayDate = CDate(ws.Range(CELL_TODAY).Value)
    End If
    
    Application.ScreenUpdating = False
    
    ' 既存のイナズマ線（図形）を削除 (逆順で削除して問題回避)
    Dim shp As Shape
    Dim shapesToDelete As Collection
    Set shapesToDelete = New Collection
    
    stage = "deleteShapes"
    For Each shp In ws.Shapes
        If Left(shp.Name, 7) = "Inazuma" Then
            shapesToDelete.Add shp
        End If
    Next shp
    
    Dim s As Variant
    For Each s In shapesToDelete
        s.Delete
    Next s
    
    Dim lastRow As Long
    stage = "lastRow"
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then
        Application.ScreenUpdating = True
        MsgBox "タスクデータがありません。", vbExclamation
        Exit Sub
    End If
    
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    ' 今日の日付の列を特定
    Dim todayCol As Long
    todayCol = -1
    Dim c As Long
    Dim cellDate As Variant
    
    stage = "findToday"
    For c = ganttStartCol To ganttStartCol + GANTT_DAYS - 1
        cellDate = ws.Cells(ROW_HEADER, c).Value
        If IsDate(cellDate) Then
            If CLng(cellDate) = CLng(todayDate) Then
                todayCol = c
                Exit For
            End If
        End If
    Next c
    
    If todayCol = -1 Then
        Application.ScreenUpdating = True
        MsgBox "今日の日付がガントチャート範囲内にありません。" & vbCrLf & "SetupInazumaGantt で日付範囲を再設定してください。", vbExclamation
        Exit Sub
    End If
    
    Dim ganttStartDate As Date
    ganttStartDate = GetGanttStartDate(ws, ganttStartCol)
    
    ' 今日線を描画（縦の赤線）
    Dim todayCell As Range
    stage = "drawTodayLine"
    Set todayCell = ws.Cells(ROW_DATA_START, todayCol)
    Dim todayLine As Shape
    Set todayLine = ws.Shapes.AddLine( _
        todayCell.Left + todayCell.Width / 2, todayCell.Top, _
        todayCell.Left + todayCell.Width / 2, ws.Cells(lastRow, todayCol).Top + ws.Cells(lastRow, todayCol).Height)
    todayLine.Name = "InazumaToday"
    todayLine.Line.ForeColor.RGB = COLOR_TODAY
    todayLine.Line.Weight = TODAY_LINE_WEIGHT
    
    ' イナズマ線（各タスクの進捗ポイントを結ぶ）
    Dim points() As Double
    ReDim points(1 To (lastRow - ROW_DATA_START + 1) * 2, 1 To 2)
    
    Dim r As Long
    Dim idx As Long
    idx = 0
    
    Dim startPlan As Variant
    Dim endPlan As Variant
    Dim progressRate As Double
    Dim expectedProgress As Double
    Dim deviation As Double
    Dim totalDays As Long
    Dim deviationDays As Long
    Dim progressCol As Long
    Dim targetCell As Range
    Dim statusText As String
    
    stage = "buildPoints"
    For r = ROW_DATA_START To lastRow
        startPlan = ws.Cells(r, COL_START_PLAN).Value
        endPlan = ws.Cells(r, COL_END_PLAN).Value
        statusText = Trim$(CStr(ws.Cells(r, COL_STATUS).Value))
        
        ' 進捗率を安全に取得（"70%" などの文字列も許容）
        progressRate = GetProgressRate(ws.Cells(r, COL_PROGRESS).Value)
        
        If IsDate(startPlan) And IsDate(endPlan) Then
            idx = idx + 1
            
            totalDays = CLng(endPlan) - CLng(startPlan)
            
            ' ゼロ除算ガード
            If totalDays <= 0 Then
                If todayDate < CDate(startPlan) Then
                    expectedProgress = 0
                Else
                    expectedProgress = 1
                End If
            Else
                If todayDate < startPlan Then
                    expectedProgress = 0
                ElseIf todayDate > endPlan Then
                    expectedProgress = 1
                Else
                    expectedProgress = (todayDate - startPlan) / totalDays
                End If
            End If
            
            If statusText = "進行中" Then
                ' 進行中は進捗率に応じて位置を決める
                If totalDays <= 0 Then
                    progressCol = DateToGanttCol(CDate(startPlan), ganttStartDate, ganttStartCol)
                Else
                    progressCol = DateToGanttCol(CDate(startPlan) + (totalDays * progressRate), ganttStartDate, ganttStartCol)
                End If
            ElseIf statusText = "完了" And progressRate >= expectedProgress Then
                ' 完了かつ進捗が進んでいる場合
                If Abs(progressRate - expectedProgress) < 0.0001 Then
                    ' 予定通り完了なら今日線と同じ位置
                    progressCol = todayCol
                Else
                    ' 進んでいるなら完了予定日に固定
                    progressCol = DateToGanttCol(CDate(endPlan), ganttStartDate, ganttStartCol)
                End If
            ElseIf statusText = "未着手" And progressRate < expectedProgress Then
                ' 未着手かつ遅れている場合は開始予定日に固定
                progressCol = DateToGanttCol(CDate(startPlan), ganttStartDate, ganttStartCol)
            Else
                deviation = progressRate - expectedProgress
                
                If totalDays > 0 Then
                    deviationDays = CLng(deviation * totalDays)
                Else
                    deviationDays = 0
                End If
                
                progressCol = todayCol + deviationDays
            End If
            If progressCol < ganttStartCol Then progressCol = ganttStartCol
            If progressCol > ganttStartCol + GANTT_DAYS - 1 Then progressCol = ganttStartCol + GANTT_DAYS - 1
            
            Set targetCell = ws.Cells(r, progressCol)
            points(idx, 1) = targetCell.Left + targetCell.Width / 2
            points(idx, 2) = targetCell.Top + targetCell.Height / 2
        End If
    Next r
    
    ' ポイントを結ぶ折れ線を描画 (1点でも描画できるようにマーカー追加)
    If idx >= 1 Then
        stage = "drawLine"
        If idx = 1 Then
            ' 1点の場合はマーカーとして円を描画
            Dim marker As Shape
            Set marker = ws.Shapes.AddShape(msoShapeOval, points(1, 1) - 3, points(1, 2) - 3, 6, 6)
            marker.Name = "InazumaMarker"
            marker.Fill.ForeColor.RGB = COLOR_INAZUMA
            marker.Line.Visible = msoFalse
        Else
            ' 2点以上の場合は折れ線
            Dim freeformBuilder As FreeformBuilder
            Set freeformBuilder = ws.Shapes.BuildFreeform(msoEditingAuto, points(1, 1), points(1, 2))
            
            Dim i As Long
            For i = 2 To idx
                freeformBuilder.AddNodes msoSegmentLine, msoEditingAuto, points(i, 1), points(i, 2)
            Next i
            
            Dim lightning As Shape
            Set lightning = freeformBuilder.ConvertToShape
            lightning.Name = "InazumaLine"
            lightning.Line.ForeColor.RGB = COLOR_INAZUMA
            lightning.Line.Weight = 2.5
            lightning.Fill.Visible = msoFalse
        End If
    End If
    
    Application.ScreenUpdating = True
    
    ' MsgBoxは RefreshInazumaGantt からのみ表示
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "DrawInazumaLine エラー(" & Err.Number & "): " & Err.Description & vbCrLf & "Stage: " & stage, vbCritical, "エラー"
End Sub

Private Function GetProgressRate(ByVal progressValue As Variant) As Double
    Dim rate As Double
    rate = 0
    
    If IsNumeric(progressValue) Then
        rate = CDbl(progressValue)
    ElseIf VarType(progressValue) = vbString Then
        Dim cleaned As String
        cleaned = Trim$(CStr(progressValue))
        cleaned = Replace$(cleaned, "％", "")
        cleaned = Replace$(cleaned, "%", "")
        If IsNumeric(cleaned) Then
            rate = CDbl(cleaned)
        End If
    End If
    
    ' 100超の値はパーセント表記とみなして変換
    If rate > 1 Then rate = rate / 100
    ' 0～1の範囲に正規化
    If rate < 0 Then rate = 0
    If rate > 1 Then rate = 1
    
    GetProgressRate = rate
End Function

' ==========================================
'  全描画実行 (MsgBoxはここでのみ表示)
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
    NormalizeProgressValues ws, lastRow
    UpdateStatusFromProgress ws, lastRow
    
    Call DrawGanttBars
    Call DrawActualBars
    Call DrawInazumaLine
    
    MsgBox "イナズマガント更新完了！" & vbCrLf & vbCrLf & _
           "【見方】" & vbCrLf & _
           "・赤線 = 今日" & vbCrLf & _
           "・オレンジ線 = 各タスクの進捗位置" & vbCrLf & _
           "・オレンジが赤より左 → 遅れ", vbInformation, "イナズマガント"
    Exit Sub
    
ErrorHandler:
    MsgBox "更新中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  ガント全体の罫線（見やすさ向上）
' ==========================================
Private Sub ApplyGanttBorders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttEndCol As Long
    ganttEndCol = ganttStartCol + GANTT_DAYS - 1
    
    Dim borderRange As Range
    Set borderRange = ws.Range(ws.Cells(ROW_HEADER_LABEL, 1), ws.Cells(lastRow, ganttEndCol))
    
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
'  週の区切り線（左罫線を太く）
' ==========================================
Private Sub DrawWeekSeparators(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim colIndex As Long
    Dim weekRange As Range
    
    ' 列ごとに一括設定（パフォーマンス改善）
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
'  進捗率の自動補正（0%～100%に丸めて保存）
' ==========================================
Private Sub NormalizeProgressValues(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    Dim progressValue As Variant
    Dim rate As Double
    
    For r = ROW_DATA_START To lastRow
        progressValue = ws.Cells(r, COL_PROGRESS).Value
        If Trim$(CStr(progressValue)) <> "" Then
            rate = GetProgressRate(progressValue)
            If rate < 0 Then rate = 0
            If rate > 1 Then rate = 1
            ws.Cells(r, COL_PROGRESS).Value = rate
        End If
    Next r
End Sub

' ==========================================
'  進捗率から状況を自動更新
' ==========================================
Private Sub UpdateStatusFromProgress(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    Dim progressValue As Variant
    Dim rate As Double
    
    For r = ROW_DATA_START To lastRow
        progressValue = ws.Cells(r, COL_PROGRESS).Value
        If Trim$(CStr(progressValue)) = "" Then
            ws.Cells(r, COL_STATUS).Value = "未着手"
        Else
            rate = GetProgressRate(progressValue)
            If rate >= 0.999 Then
                ws.Cells(r, COL_STATUS).Value = "完了"
            ElseIf rate <= 0 Then
                ws.Cells(r, COL_STATUS).Value = "未着手"
            Else
                ws.Cells(r, COL_STATUS).Value = "進行中"
            End If
        End If
    Next r
End Sub

' ==========================================
'  ダブルクリックでタスクを完了（シートイベントから呼び出し）
' ==========================================
Public Sub CompleteTaskByDoubleClick(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = Target.Worksheet
    
    ' データ範囲内のセルかチェック（A列～N列、データ開始行以降）
    Dim targetRow As Long
    targetRow = Target.Row
    
    If targetRow < ROW_DATA_START Then Exit Sub
    
    Dim targetCol As Long
    targetCol = Target.Column
    
    Dim hierCol As Long, endActualCol As Long
    hierCol = ws.Columns(COL_HIERARCHY).Column
    endActualCol = ws.Columns(COL_END_ACTUAL).Column
    
    If targetCol < hierCol Or targetCol > endActualCol Then Exit Sub
    
    ' タスクが既に完了している場合は何もしない
    Dim currentStatus As String
    currentStatus = Trim$(CStr(ws.Cells(targetRow, COL_STATUS).Value))
    
    If currentStatus = "完了" Then
        MsgBox "このタスクは既に完了しています。", vbInformation, "タスク完了"
        Exit Sub
    End If
    
    ' 確認ダイアログ
    Dim taskName As String
    taskName = ws.Cells(targetRow, COL_TASK).Value
    
    Dim result As VbMsgBoxResult
    result = MsgBox("タスク: " & taskName & vbCrLf & vbCrLf & _
                    "このタスクを完了にしますか？" & vbCrLf & _
                    "・進捗率 → 100%" & vbCrLf & _
                    "・状況 → 完了" & vbCrLf & _
                    "・完了実績 → 今日の日付（開始実績がある場合のみ）", _
                    vbYesNo + vbQuestion, "タスク完了")
    
    If result = vbNo Then Exit Sub
    
    Application.EnableEvents = False
    
    ' 進捗率を100%に
    ws.Cells(targetRow, COL_PROGRESS).Value = 1
    
    ' 状況を「完了」に
    ws.Cells(targetRow, COL_STATUS).Value = "完了"
    
    ' 開始実績がある場合は完了実績に今日の日付を設定
    Dim startActual As Variant
    startActual = ws.Cells(targetRow, COL_START_ACTUAL).Value
    
    If IsDate(startActual) Then
        ws.Cells(targetRow, COL_END_ACTUAL).Value = Date
        MsgBox "タスクを完了しました！" & vbCrLf & vbCrLf & _
               "完了実績: " & Format(Date, "yyyy/mm/dd"), vbInformation, "完了"
    Else
        MsgBox "タスクを完了しました！" & vbCrLf & vbCrLf & _
               "（開始実績が未入力のため、完了実績は設定されませんでした）", vbInformation, "完了"
    End If
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  タスク入力列から階層を自動判定
' ==========================================
' C列 → LV1, D列 → LV2, E列 → LV3, F列 → LV4
' targetRow を指定した場合はその行のみ処理、0の場合は全行処理
Public Sub AutoDetectTaskLevel(Optional ByVal targetRow As Long = 0)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim startRow As Long, endRow As Long
    
    If targetRow > 0 Then
        ' 指定行のみ処理
        If targetRow < ROW_DATA_START Then Exit Sub
        startRow = targetRow
        endRow = targetRow
    Else
        ' 全行処理
        startRow = ROW_DATA_START
        endRow = GetLastDataRow(ws)
        If endRow < ROW_DATA_START Then endRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    End If
    
    Application.EnableEvents = False
    
    Dim r As Long
    Dim taskLevel As Long
    Dim taskText As String
    
    For r = startRow To endRow
        taskLevel = 0
        
        ' C～F列をチェック（優先順位: F→E→D→C）
        ' 最も右の列（最も深い階層）を優先
        If Trim$(CStr(ws.Cells(r, "F").Value)) <> "" Then
            taskLevel = 4
        ElseIf Trim$(CStr(ws.Cells(r, "E").Value)) <> "" Then
            taskLevel = 3
        ElseIf Trim$(CStr(ws.Cells(r, "D").Value)) <> "" Then
            taskLevel = 2
        ElseIf Trim$(CStr(ws.Cells(r, "C").Value)) <> "" Then
            taskLevel = 1
        End If
        
        ' A列（LV列）に階層レベルを設定
        If taskLevel > 0 Then
            ws.Cells(r, COL_HIERARCHY).Value = taskLevel
        Else
            ' タスクが何も入力されていない場合はLVをクリア
            ws.Cells(r, COL_HIERARCHY).ClearContents
        End If
    Next r
    
    Application.EnableEvents = True
    
    If targetRow = 0 Then
        MsgBox "全タスクの階層レベルを自動設定しました。", vbInformation, "階層自動判定"
    End If
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "階層自動判定エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  タスク列の開始位置を取得（階層レベルから）
' ==========================================
' 階層別の色塗りで使用: タスクが入力された列の文字を取得
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
            GetTaskColumnByLevel = "C" ' デフォルト
    End Select
End Function
