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
Public Const ROW_DATE_HEADER As Long = 7      ' 日付行（ガント）
Public Const ROW_HEADER As Long = 8           ' 曜日行（ガント）/ 項目ヘッダー行（A-N列）
Public Const ROW_DATA_START As Long = 9       ' データ開始行
Public Const GANTT_DAYS As Long = 120         ' ガントチャートの日数
Public Const DATA_ROWS_DEFAULT As Long = 200  ' 初期入力範囲の行数
Public Const HOLIDAY_SHEET_NAME As String = "祝日マスタ"
Public Const GUIDE_SHEET_NAME As String = "InazumaGantt_説明"
Public Const MAIN_SHEET_NAME As String = "InazumaGantt_v2"
Public Const GUIDE_LEGEND_START_CELL As String = "E1"
Public Const CELL_PROJECT_START As String = "L2"
Public Const CELL_DISPLAY_WEEK As String = "L3"
Public Const CELL_TODAY As String = "L4"

' 色設定
Public Const COLOR_PLAN As Long = 16119285       ' RGB(245,245,245) 限りなく白に近い灰色
Public Const COLOR_PROGRESS As Long = 9851952    ' RGB(48,84,150) 紺色
Public Const COLOR_HOLIDAY As Long = 5263430     ' RGB(70,70,80) 濃い灰色（休日祝日）
Public Const COLOR_ROW_BAND As Long = 16316664
Public Const COLOR_ACTUAL As Long = 5287936      ' RGB(0,176,80) 緑色
Public Const COLOR_TODAY As Long = 255           ' RGB(255,0,0) 赤
Public Const COLOR_WARN As Long = 13434879
Public Const COLOR_ERROR As Long = 13553151
Public Const COLOR_INAZUMA As Long = 42495       ' RGB(255,165,0) オレンジ
Public Const COLOR_HEADER_BG As Long = 12874308
Public Const COLOR_GANTT_HEADER As Long = 8421504
Public Const COLOR_WEEKEND As Long = 5263430     ' RGB(70,70,80) 濃い灰色
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
    ws.Range("A4").Value = "メモ："
    
    ' 情報エリア（K-L列）
    ws.Range("K2").Value = "開始日："
    ws.Range("K3").Value = "週表示:"
    ws.Range("K4").Value = "今日："
    
    ' ヘッダー設定 (ROW_HEADER = 8行目に統一)
    ws.Range(COL_HIERARCHY & ROW_HEADER).Value = "LV"
    ws.Range(COL_NO & ROW_HEADER).Value = "No."
    ws.Range("C" & ROW_HEADER).Value = "TASK"
    ' D-F列はタスク入力用（ヘッダーなし）
    ws.Range(COL_TASK_DETAIL & ROW_HEADER).Value = "タスク詳細"
    ws.Range(COL_STATUS & ROW_HEADER).Value = "状況"
    ws.Range(COL_PROGRESS & ROW_HEADER).Value = "進捗率"
    ws.Range(COL_ASSIGNEE & ROW_HEADER).Value = "担当"
    ws.Range(COL_START_PLAN & ROW_HEADER).Value = "開始予定"
    ws.Range(COL_END_PLAN & ROW_HEADER).Value = "完了予定"
    ws.Range(COL_START_ACTUAL & ROW_HEADER).Value = "開始実績"
    ws.Range(COL_END_ACTUAL & ROW_HEADER).Value = "完了実績"
    
    ' ヘッダー行のスタイル（8行目、A～N列）
    With ws.Range("A" & ROW_HEADER & ":N" & ROW_HEADER)
        .Font.Bold = True
        .Interior.Color = COLOR_HEADER_BG
        .Font.Color = RGB(255, 255, 255)
    End With
    
    ' 列幅設定（改善メモ仕様に準拠）
    ws.Columns("A").ColumnWidth = 3     ' LV
    ws.Columns("B").ColumnWidth = 4     ' No.
    ws.Columns("C").ColumnWidth = 4     ' TASK Lv1
    ws.Columns("D").ColumnWidth = 4     ' TASK Lv2
    ws.Columns("E").ColumnWidth = 4     ' TASK Lv3
    ws.Columns("F").ColumnWidth = 15    ' TASK Lv4
    ws.Columns("G").ColumnWidth = 20    ' タスク補足
    ws.Columns("H").ColumnWidth = 7     ' 状況
    ws.Columns("I").ColumnWidth = 7     ' 進捗率
    ws.Columns("J").ColumnWidth = 7     ' 担当
    ws.Columns("K").ColumnWidth = 8.7   ' 開始予定
    ws.Columns("L").ColumnWidth = 8.7   ' 完了予定
    ws.Columns("M").ColumnWidth = 8.7   ' 開始実績
    ws.Columns("N").ColumnWidth = 8.7   ' 完了実績
    
    ' 行高さ統一（22）
    ws.Rows.RowHeight = 22


    EnsureHolidaySheet
    EnsureGuideSheet
    
    ' 日付開始日を入力させる（キャンセル時はロールバック）
    Dim startDateInput As Variant
    startDateInput = Application.InputBox("ガントチャートの開始日を入力してください (例: 24/12/25)", "開始日設定", Format(Date, "yy/mm/dd"), Type:=2)
    
    ' キャンセル処理（ロールバック）
    If startDateInput = False Or VarType(startDateInput) = vbBoolean Then
        ' シートの内容をクリア（ロールバック）: 想定範囲のみ
        Dim rollbackEndCol As Long
        rollbackEndCol = ws.Columns(COL_GANTT_START).Column + GANTT_DAYS - 1
        Dim rollbackEndRow As Long
        rollbackEndRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
        ws.Range(ws.Cells(1, 1), ws.Cells(rollbackEndRow, rollbackEndCol)).Clear
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "セットアップがキャンセルされました。", vbInformation, "キャンセル"
        Exit Sub
    End If
    
    Dim ganttStartDate As Date
    If IsDate(startDateInput) Then
        ganttStartDate = CDate(startDateInput)
    Else
        ganttStartDate = Date
    End If

    
    ws.Range(CELL_PROJECT_START).Value = ganttStartDate
    ws.Range(CELL_PROJECT_START).NumberFormat = "yy/mm/dd"
    ws.Range(CELL_DISPLAY_WEEK).Value = 1
    ws.Range(CELL_TODAY).Value = Date
    ws.Range(CELL_TODAY).NumberFormat = "yy/mm/dd"
    
    ' 日付列の生成
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column

    Dim todayDate As Date
    todayDate = Date
    If IsDate(ws.Range(CELL_TODAY).Value) Then
        todayDate = CDate(ws.Range(CELL_TODAY).Value)
    End If
    
    ' 週・日付・曜日ヘッダーの作成
    Dim weekStartCol As Long
    Dim weekEndCol As Long
    Dim currentDate As Date
    Dim colIndex As Long
    Dim i As Long
    
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        
        ' 7行目: 日付（日のみ）- ヘッダーと同じ色
        ws.Cells(ROW_DATE_HEADER, colIndex).Value = Day(currentDate)
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Size = 9
        ws.Cells(ROW_DATE_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_DATE_HEADER, colIndex).Interior.Color = COLOR_HEADER_BG
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Color = RGB(255, 255, 255)
        
        ' 8行目: 曜日 - ヘッダーと同じ色
        ws.Cells(ROW_HEADER, colIndex).Value = Format$(currentDate, "aaa")
        ws.Cells(ROW_HEADER, colIndex).Font.Size = 8
        ws.Cells(ROW_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_HEADER, colIndex).Interior.Color = COLOR_HEADER_BG
        ws.Cells(ROW_HEADER, colIndex).Font.Color = RGB(255, 255, 255)
        
        ' 列幅を設定
        ws.Columns(colIndex).ColumnWidth = 3
        
        ' 6行目: 週ヘッダー（7日単位）
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
    ApplyWeekendColors ws, lastRow, ganttStartDate, ganttStartCol
    ApplyDataValidationAndFormats ws, lastRow
    
    ' 目盛線をオフ
    ActiveWindow.DisplayGridlines = False
    
    ' フィルタ自動設定 (8行目A-N列)
    If Not ws.AutoFilterMode Then
        ws.Range("A" & ROW_HEADER & ":N" & ROW_HEADER).AutoFilter
    End If
    
    ' No.1〜400の初期採番
    Dim noRow As Long
    For noRow = ROW_DATA_START To ROW_DATA_START + 399
        ws.Cells(noRow, COL_NO).Value = noRow - ROW_DATA_START + 1
    Next noRow
    
    ' コントロールボタンの作成
    CreateControlButtons ws
    
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
    Else
        wsHoliday.Cells.Clear
    End If
    
    ' ヘッダー
    wsHoliday.Range("A1").Value = "祝日"
    wsHoliday.Range("A1").Font.Bold = True
    wsHoliday.Range("A1").Interior.Color = RGB(48, 84, 150)
    wsHoliday.Range("A1").Font.Color = RGB(255, 255, 255)
    wsHoliday.Columns("A").NumberFormat = "yy/mm/dd"
    wsHoliday.Columns("A").ColumnWidth = 12
    
    ' 入力エリアの罫線（A2:A30）
    With wsHoliday.Range("A2:A30").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 48
    End With
    
    ' 説明テキスト
    wsHoliday.Range("C1").Value = "【祝日マスタの使い方】"
    wsHoliday.Range("C1").Font.Bold = True
    wsHoliday.Range("C2").Value = "A列に祝日の日付を入力してください。"
    wsHoliday.Range("C3").Value = "入力した日付はガントチャート上でピンク色で表示されます。"
    wsHoliday.Range("C4").Value = ""
    wsHoliday.Range("C5").Value = "例: 26/01/01, 26/01/13, 26/02/11 ..."
    wsHoliday.Range("C6").Value = ""
    wsHoliday.Range("C7").Value = "※ ガント更新後に反映されます。"
    wsHoliday.Columns("C").ColumnWidth = 40
    
    ' 目盛線オフ
    ActiveWindow.DisplayGridlines = False
End Sub

' ==========================================
'  入力規則と日付書式の適用
' ==========================================
Private Sub ApplyDataValidationAndFormats(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    ' 進捗率のドロップダウン
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
End Sub

' ==========================================
'  データ最終行の取得
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
    
    ' 目盛線オフ
    ActiveWindow.DisplayGridlines = False
    
    ' タイトル
    wsGuide.Cells(1, 1).Value = "進捗管理表 操作マニュアル"
    wsGuide.Cells(1, 1).Font.Bold = True
    wsGuide.Cells(1, 1).Font.Size = 14
    
    ' ボタン機能
    wsGuide.Cells(3, 1).Value = "■ ボタン機能"
    wsGuide.Cells(3, 1).Font.Bold = True
    wsGuide.Cells(4, 1).Value = "【ガント更新】"
    wsGuide.Cells(4, 2).Value = "ガントチャートを最新状態に再描画します。"
    wsGuide.Cells(5, 2).Value = "進捗率や日付を変更した後は必ずクリックしてください。"
    wsGuide.Cells(6, 1).Value = "【土日切替】"
    wsGuide.Cells(6, 2).Value = "土日列の表示/非表示を切替えます。"
    wsGuide.Cells(7, 2).Value = "画面を広く使いたい時に便利です。"
    wsGuide.Cells(8, 1).Value = "【書式リセット】"
    wsGuide.Cells(8, 2).Value = "崩れた罫線・書式を修復します。"
    wsGuide.Cells(9, 2).Value = "表示がおかしくなった時に使用してください。"
    
    ' ダブルクリック完了
    wsGuide.Cells(11, 1).Value = "■ ダブルクリックでタスク完了"
    wsGuide.Cells(11, 1).Font.Bold = True
    wsGuide.Cells(12, 1).Value = "No.列(B列) または LV列(A列) をダブルクリックすると、"
    wsGuide.Cells(13, 1).Value = "そのタスクが完了になります。"
    wsGuide.Cells(14, 1).Value = ""
    wsGuide.Cells(15, 1).Value = "  ・ 状況 → 「完了」"
    wsGuide.Cells(16, 1).Value = "  ・ 進捗率 → 100%"
    wsGuide.Cells(17, 1).Value = "  ・ 完了実績 → 今日の日付"
    wsGuide.Cells(18, 1).Value = ""
    wsGuide.Cells(19, 1).Value = "※ すでに完了しているタスクは変更されません。"
    
    ' 列幅設定
    wsGuide.Columns(1).ColumnWidth = 35
    wsGuide.Columns(2).ColumnWidth = 45
End Sub

' ==========================================
'  ガント全体の罫線（罫線サマリに基づく詳細パターン適用）
' ==========================================
Private Sub ApplyGanttBorders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttEndCol As Long
    ganttEndCol = ganttStartCol + GANTT_DAYS - 1
    
    ' 罫線をクリア
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ganttEndCol)).Borders.LineStyle = xlNone
    
    ' --- P1: 1行目 (K:L 下罫線) ---
    ApplyBorder ws.Range("K1:L1"), xlEdgeBottom, xlContinuous, xlThin, xlColorIndexAutomatic
    
    ' --- P2: 2-4行目 (K:L 上下左右罫線) ---
    Dim r As Long
    For r = 2 To 4
        ApplyBorder ws.Range("K" & r & ":L" & r), xlEdgeTop, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range("K" & r & ":L" & r), xlEdgeBottom, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range("J" & r), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range("L" & r), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range("K" & r), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range("M" & r), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
    Next r
    
    ' --- P3: 5行目 (K:L 上, O:BA 下) ---
    ApplyBorder ws.Range("K5:L5"), xlEdgeTop, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(5, ganttStartCol), ws.Cells(5, ganttEndCol)), xlEdgeBottom, xlContinuous, xlThin, xlColorIndexAutomatic
    
    ' --- P4: 6行目 (週ヘッダー行) ---
    ' 上: O, V, AC, AJ, AQ, AX (7列おき)
    Dim weekCol As Long
    For weekCol = ganttStartCol To ganttEndCol Step 7
        ApplyBorder ws.Cells(6, weekCol), xlEdgeTop, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Cells(6, weekCol), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Cells(6, weekCol), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
    Next weekCol
    ' 下: A:O + 週区切り (中太)
    ApplyBorder ws.Range(ws.Cells(6, 1), ws.Cells(6, ganttStartCol)), xlEdgeBottom, xlContinuous, xlMedium, xlColorIndexAutomatic
    For weekCol = ganttStartCol To ganttEndCol Step 7
        ApplyBorder ws.Cells(6, weekCol), xlEdgeBottom, xlContinuous, xlMedium, xlColorIndexAutomatic
    Next weekCol
    ApplyBorder ws.Range("N6"), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
    
    ' --- P5: 7行目 (日付行) ---
    ' 7行目の背景色をヘッダーと同じ色で塗りつぶし
    ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)).Interior.Color = COLOR_HEADER_BG
    ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)).Font.Color = RGB(255, 255, 255)
    
    ApplyBorder ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)), xlEdgeTop, xlContinuous, xlMedium, xlColorIndexAutomatic
    ' 7行目下部に黒色の太線
    ApplyBorder ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)), xlEdgeBottom, xlContinuous, xlMedium, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(7, 14), ws.Cells(7, ganttEndCol)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range("A7"), xlEdgeLeft, xlContinuous, xlMedium, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(7, ganttStartCol), ws.Cells(7, ganttEndCol)), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
    
    ' 7行目のO列より右のガントチャート部は太字
    ws.Range(ws.Cells(7, ganttStartCol), ws.Cells(7, ganttEndCol)).Font.Bold = True
    
    ' --- P6: 8行目 (ヘッダー行) ---
    ApplyBorder ws.Range(ws.Cells(8, 1), ws.Cells(8, ganttEndCol)), xlEdgeTop, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(8, 1), ws.Cells(8, ganttEndCol)), xlEdgeBottom, xlContinuous, xlMedium, xlColorIndexAutomatic
    ApplyBorder ws.Range("A8:B8"), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(8, 6), ws.Cells(8, ganttEndCol)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range("A8"), xlEdgeLeft, xlContinuous, xlMedium, xlColorIndexAutomatic
    ApplyBorder ws.Range("B8:C8"), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(8, 7), ws.Cells(8, ganttEndCol)), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
    
    ' --- P7/P8: 9行目以降 (データ行パターン、9行目も10行目以降と同じ) ---
    If lastRow >= ROW_DATA_START Then
        Dim dataRange As Range
        Set dataRange = ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, ganttEndCol))
        
        ' 上下: ColorIndex 48 (薄い灰色)
        ApplyBorderWithColorIndex dataRange, xlEdgeTop, xlContinuous, xlThin, 48
        ApplyBorderWithColorIndex dataRange, xlEdgeBottom, xlContinuous, xlThin, 48
        ApplyBorderWithColorIndex ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, ganttEndCol)), xlInsideHorizontal, xlContinuous, xlThin, 48
        
        ' C-E列: 極細 ColorIndex 15
        ApplyBorderWithColorIndex ws.Range(ws.Cells(ROW_DATA_START, 3), ws.Cells(lastRow, 5)), xlEdgeRight, xlContinuous, xlHairline, 15
        ApplyBorderWithColorIndex ws.Range(ws.Cells(ROW_DATA_START, 4), ws.Cells(lastRow, 6)), xlEdgeLeft, xlContinuous, xlHairline, 15
        ApplyBorderWithColorIndex ws.Range(ws.Cells(ROW_DATA_START, 3), ws.Cells(lastRow, 5)), xlInsideVertical, xlContinuous, xlHairline, 15
        
        ' ガントチャート部(O列以降)にもC-D間と同じ縦罫線
        ApplyBorderWithColorIndex ws.Range(ws.Cells(ROW_DATE_HEADER, ganttStartCol), ws.Cells(lastRow, ganttEndCol)), xlInsideVertical, xlContinuous, xlHairline, 15
        
        ' A-B, F-N列: 細線 自動
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, 2)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ' A列B列間は黒細線
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, 1)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 6), ws.Cells(lastRow, 14)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 6), ws.Cells(lastRow, 14)), xlInsideVertical, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, 3)), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 7), ws.Cells(lastRow, ganttStartCol)), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic
    End If
End Sub

' ==========================================
'  罫線適用ヘルパー（自動色）
' ==========================================
Private Sub ApplyBorder(ByVal rng As Range, ByVal borderIndex As XlBordersIndex, _
                        ByVal lineStyle As XlLineStyle, ByVal weight As XlBorderWeight, _
                        ByVal colorIndex As Long)
    On Error Resume Next
    With rng.Borders(borderIndex)
        .LineStyle = lineStyle
        .Weight = weight
        .ColorIndex = colorIndex
    End With
    On Error GoTo 0
End Sub

' ==========================================
'  罫線適用ヘルパー（ColorIndex指定）
' ==========================================
Private Sub ApplyBorderWithColorIndex(ByVal rng As Range, ByVal borderIndex As XlBordersIndex, _
                                      ByVal lineStyle As XlLineStyle, ByVal weight As XlBorderWeight, _
                                      ByVal colorIdx As Long)
    On Error Resume Next
    With rng.Borders(borderIndex)
        .LineStyle = lineStyle
        .Weight = weight
        .ColorIndex = colorIdx
    End With
    On Error GoTo 0
End Sub


' ==========================================
'  週の区切り線
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
'  土日列の色塗り（曜日・日付行とデータ行を含む）
' ==========================================
Private Sub ApplyWeekendColors(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal ganttStartDate As Date, ByVal ganttStartCol As Long)
    Dim colIndex As Long
    Dim currentDate As Date
    Dim i As Long
    
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        
        ' 土日（土=6, 日=7）の列を薄い灰色で塗りつぶす（日付行、曜日行、データ行すべて）
        If Weekday(currentDate, vbMonday) >= 6 Then
            ws.Range(ws.Cells(ROW_DATE_HEADER, colIndex), ws.Cells(lastRow, colIndex)).Interior.Color = COLOR_HOLIDAY
        End If
    Next i
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
    
    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        ganttStartDate = Date
    End If
    
    ' 既存のシェイプを削除
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "Bar_" Or Left(shp.Name, 6) = "Today_" Or Left(shp.Name, 8) = "Inazuma_" Then
            shp.Delete
        End If
    Next shp
    
    ' 各行のバーを描画
    Dim r As Long
    Dim startPlan As Variant, endPlan As Variant
    Dim startActual As Variant, endActual As Variant
    Dim progress As Double
    Dim startCol As Long, endCol As Long, progressCol As Long
    Dim cellTop As Double, cellLeft As Double, cellWidth As Double, cellHeight As Double
    Dim barHeight As Double
    
    barHeight = 12  ' バーの高さ
    
    Dim inazumaPoints() As Variant
    ReDim inazumaPoints(1 To lastRow - ROW_DATA_START + 1, 1 To 2)
    Dim inazumaCount As Long
    inazumaCount = 0
    
    For r = ROW_DATA_START To lastRow
        ' 日付を取得
        startPlan = ws.Cells(r, COL_START_PLAN).Value
        endPlan = ws.Cells(r, COL_END_PLAN).Value
        startActual = ws.Cells(r, COL_START_ACTUAL).Value
        endActual = ws.Cells(r, COL_END_ACTUAL).Value
        
        ' 進捗率を取得
        progress = 0
        If IsNumeric(ws.Cells(r, COL_PROGRESS).Value) Then
            progress = CDbl(ws.Cells(r, COL_PROGRESS).Value)
            If progress > 1 Then progress = progress / 100
            If progress < 0 Then progress = 0
            If progress > 1 Then progress = 1
        End If
        
        ' 予定バーを描画
        If IsDate(startPlan) And IsDate(endPlan) Then
            startCol = DateToColumn(ganttStartDate, CDate(startPlan), ganttStartCol)
            endCol = DateToColumn(ganttStartDate, CDate(endPlan), ganttStartCol)
            
            If startCol >= ganttStartCol And startCol <= ganttStartCol + GANTT_DAYS - 1 Then
                If endCol > ganttStartCol + GANTT_DAYS - 1 Then endCol = ganttStartCol + GANTT_DAYS - 1
                If endCol >= startCol Then
                    cellTop = ws.Cells(r, startCol).Top + 2
                    cellLeft = ws.Cells(r, startCol).Left
                    cellWidth = ws.Cells(r, endCol).Left + ws.Cells(r, endCol).Width - cellLeft
                    barHeight = 6  ' 予定バーの高さ
                    
                    ' 予定バー（薄い灰色 + 黒枠線）
                    Set shp = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, cellWidth, barHeight)
                    shp.Name = "Bar_Plan_" & r
                    shp.Fill.ForeColor.RGB = COLOR_PLAN
                    shp.Line.Visible = msoTrue
                    shp.Line.ForeColor.RGB = RGB(0, 0, 0)  ' 黒枠線
                    shp.Line.Weight = 1
                    
                    ' 進捗バー（紺色 + 黒枠線）
                    If progress > 0 Then
                        progressCol = startCol + CLng((endCol - startCol + 1) * progress) - 1
                        If progressCol < startCol Then progressCol = startCol
                        If progressCol >= startCol Then
                            Dim progressWidth As Double
                            progressWidth = ws.Cells(r, progressCol).Left + ws.Cells(r, progressCol).Width - cellLeft
                            If progressWidth < ws.Cells(r, startCol).Width Then progressWidth = ws.Cells(r, startCol).Width
                            If progress >= 1 Then progressWidth = cellWidth
                            
                            Set shp = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, progressWidth, barHeight)
                            shp.Name = "Bar_Progress_" & r
                            shp.Fill.ForeColor.RGB = COLOR_PROGRESS
                            shp.Line.Visible = msoTrue
                            shp.Line.ForeColor.RGB = RGB(0, 0, 0)  ' 黒枠線
                            shp.Line.Weight = 1
                        End If
                    End If
                    
                    ' イナズマ線用のポイントを記録（今日基準型）
                    ' 条件: 開始予定日が今日以前のタスクのみ対象
                    If CDate(startPlan) <= Date Then
                        Dim inazumaX As Double
                        Dim useTodayPosition As Boolean
                        useTodayPosition = False
                        
                        ' 今日列のX座標を計算
                        Dim todayColForInazuma As Long
                        todayColForInazuma = DateToColumn(ganttStartDate, Date, ganttStartCol)
                        Dim todayX As Double
                        If todayColForInazuma >= ganttStartCol And todayColForInazuma <= ganttStartCol + GANTT_DAYS - 1 Then
                            todayX = ws.Cells(r, todayColForInazuma).Left + ws.Cells(r, todayColForInazuma).Width / 2
                        Else
                            todayX = 0
                        End If
                        
                        If progress >= 1 Then
                            ' 完了済み
                            If CDate(endPlan) < Date Then
                                ' 完了予定日が今日より前の場合は今日の位置で結ぶ
                                useTodayPosition = True
                            Else
                                ' 完了予定日が今日以降の場合は完了予定位置で結ぶ
                                inazumaX = ws.Cells(r, endCol).Left + ws.Cells(r, endCol).Width
                            End If
                        Else
                            ' 進行中または未着手: 進捗率に応じた位置
                            Dim progressPosition As Long
                            progressPosition = startCol + CLng((endCol - startCol + 1) * progress) - 1
                            If progressPosition < startCol Then progressPosition = startCol
                            inazumaX = ws.Cells(r, progressPosition).Left + ws.Cells(r, progressPosition).Width * progress
                            If progress = 0 Then inazumaX = cellLeft
                        End If
                        
                        ' 今日の位置を使用する場合
                        If useTodayPosition And todayX > 0 Then
                            inazumaX = todayX
                        End If
                        
                        inazumaCount = inazumaCount + 1
                        inazumaPoints(inazumaCount, 1) = inazumaX
                        inazumaPoints(inazumaCount, 2) = cellTop + barHeight / 2
                    End If
                End If
            End If
        End If
        
        ' 実績バー（緑色の塗りつぶしバー、予定の下に配置、右端は進捗バーと揃える）
        If IsDate(startActual) And IsDate(startPlan) And IsDate(endPlan) Then
            ' 実績バーの右端は進捗バーの右端と揃える
            Dim actualStartCol As Long
            Dim actualEndCol As Long
            Dim planStartCol As Long
            Dim planEndCol As Long
            
            actualStartCol = DateToColumn(ganttStartDate, CDate(startActual), ganttStartCol)
            planStartCol = DateToColumn(ganttStartDate, CDate(startPlan), ganttStartCol)
            planEndCol = DateToColumn(ganttStartDate, CDate(endPlan), ganttStartCol)
            
            ' 進捗バーの右端位置を計算
            Dim progressEndCol As Long
            If progress >= 1 Then
                progressEndCol = planEndCol
            Else
                progressEndCol = planStartCol + CLng((planEndCol - planStartCol + 1) * progress) - 1
                If progressEndCol < planStartCol Then progressEndCol = planStartCol
            End If
            
            If actualStartCol >= ganttStartCol And actualStartCol <= ganttStartCol + GANTT_DAYS - 1 Then
                If progressEndCol > ganttStartCol + GANTT_DAYS - 1 Then progressEndCol = ganttStartCol + GANTT_DAYS - 1
                If progressEndCol >= actualStartCol Then
                    Dim actualBarHeight As Double
                    actualBarHeight = 6  ' 実績バーの高さ（予定より細め）
                    cellTop = ws.Cells(r, actualStartCol).Top + 10  ' 予定バーの下に配置
                    cellLeft = ws.Cells(r, actualStartCol).Left
                    ' 右端は進捗バーの右端と揃える
                    cellWidth = ws.Cells(r, progressEndCol).Left + ws.Cells(r, progressEndCol).Width - cellLeft
                    
                    Set shp = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, cellWidth, actualBarHeight)
                    shp.Name = "Bar_Actual_" & r
                    shp.Fill.ForeColor.RGB = COLOR_ACTUAL
                    shp.Line.Visible = msoFalse
                End If
            End If
        End If
    Next r
    
    ' 今日線を描画（9行目スタート）
    Dim todayCol As Long
    todayCol = DateToColumn(ganttStartDate, Date, ganttStartCol)
    
    If todayCol >= ganttStartCol And todayCol <= ganttStartCol + GANTT_DAYS - 1 Then
        ' 今日にあたる日付(7行目)を赤字にする
        ws.Cells(ROW_DATE_HEADER, todayCol).Font.Color = COLOR_TODAY
        
        ' 今日線（9行目から開始）
        Dim todayLeft As Double, todayTop As Double, todayBottom As Double
        todayLeft = ws.Cells(ROW_DATA_START, todayCol).Left + ws.Cells(ROW_DATA_START, todayCol).Width / 2
        todayTop = ws.Cells(ROW_DATA_START, todayCol).Top
        todayBottom = ws.Cells(lastRow, todayCol).Top + ws.Cells(lastRow, todayCol).Height
        
        Set shp = ws.Shapes.AddLine(todayLeft, todayTop, todayLeft, todayBottom)
        shp.Name = "Today_Line"
        shp.Line.ForeColor.RGB = COLOR_TODAY
        shp.Line.Weight = TODAY_LINE_WEIGHT
    End If
    
    ' イナズマ線を描画（複数ポイントがある場合）
    If inazumaCount >= 2 Then
        Dim freeformBuilder As FreeformBuilder
        Set freeformBuilder = ws.Shapes.BuildFreeform(msoEditingAuto, inazumaPoints(1, 1), inazumaPoints(1, 2))
        
        Dim p As Long
        For p = 2 To inazumaCount
            freeformBuilder.AddNodes msoSegmentLine, msoEditingAuto, inazumaPoints(p, 1), inazumaPoints(p, 2)
        Next p
        
        Set shp = freeformBuilder.ConvertToShape
        shp.Name = "Inazuma_Line"
        shp.Line.ForeColor.RGB = COLOR_INAZUMA
        shp.Line.Weight = 2
        shp.Fill.Visible = msoFalse
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DrawGanttBars エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  日付から列番号を計算
' ==========================================
Private Function DateToColumn(ByVal ganttStartDate As Date, ByVal targetDate As Date, ByVal ganttStartCol As Long) As Long
    Dim daysDiff As Long
    daysDiff = targetDate - ganttStartDate
    DateToColumn = ganttStartCol + daysDiff
End Function

' ==========================================
'  全描画実行
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
    ApplyHolidayColors ws, lastRow
    
    Call DrawGanttBars
    
    MsgBox "イナズマガント更新完了！", vbInformation, "イナズマガント"
    Exit Sub
    
ErrorHandler:
    MsgBox "更新中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  祝日列の色塗り（祝日マスタ A列）
' ==========================================
Private Sub ApplyHolidayColors(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim wsHoliday As Worksheet
    On Error Resume Next
    Set wsHoliday = ThisWorkbook.Worksheets(HOLIDAY_SHEET_NAME)
    On Error GoTo 0
    
    If wsHoliday Is Nothing Then Exit Sub
    
    Dim ganttStartCol As Long
    ganttStartCol = ws.Columns(COL_GANTT_START).Column
    
    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        Exit Sub
    End If
    
    Dim lastHolidayRow As Long
    lastHolidayRow = wsHoliday.Cells(wsHoliday.Rows.Count, "A").End(xlUp).Row
    If lastHolidayRow < 2 Then Exit Sub
    
    Dim r As Long
    Dim holidayDate As Date
    Dim colIndex As Long
    
    For r = 2 To lastHolidayRow
        If IsDate(wsHoliday.Cells(r, "A").Value) Then
            holidayDate = CDate(wsHoliday.Cells(r, "A").Value)
            colIndex = DateToColumn(ganttStartDate, holidayDate, ganttStartCol)
            If colIndex >= ganttStartCol And colIndex <= ganttStartCol + GANTT_DAYS - 1 Then
                ws.Range(ws.Cells(ROW_DATE_HEADER, colIndex), ws.Cells(lastRow, colIndex)).Interior.Color = COLOR_HOLIDAY
            End If
        End If
    Next r
End Sub

' ==========================================
'  タスク列の開始位置を取得（階層レベルから）
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
'  タスク入力列から階層を自動判定
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

' ==========================================
'  コントロールボタンの作成
' ==========================================
Private Sub CreateControlButtons(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' 既存ボタンを削除
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "Btn_" Then shp.Delete
    Next shp
    On Error GoTo 0
    
    Dim btnLeft As Double, btnTop As Double, btnWidth As Double, btnHeight As Double
    btnTop = ws.Cells(2, 1).Top
    btnWidth = 80
    btnHeight = 22
    
    ' ガント更新ボタン
    btnLeft = ws.Cells(2, 1).Left
    Dim btnRefresh As Shape
    Set btnRefresh = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btnRefresh
        .Name = "Btn_Refresh"
        .Fill.ForeColor.RGB = RGB(48, 84, 150)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = "ガント更新"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "RefreshInazumaGantt"
    End With
    
    ' 土日切替ボタン
    btnLeft = btnLeft + btnWidth + 10
    Dim btnToggle As Shape
    Set btnToggle = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btnToggle
        .Name = "Btn_ToggleWeekend"
        .Fill.ForeColor.RGB = RGB(68, 114, 196)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = "土日切替"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "ToggleWeekends"
    End With
    
    ' 書式リセットボタン
    btnLeft = btnLeft + btnWidth + 10
    Dim btnReset As Shape
    Set btnReset = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btnReset
        .Name = "Btn_Reset"
        .Fill.ForeColor.RGB = RGB(112, 48, 160)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = "書式リセット"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "ResetFormatting"
    End With
End Sub

' ==========================================
'  土日列の表示/非表示切替
' ==========================================
Sub ToggleWeekends()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        MsgBox "開始日が設定されていません。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim i As Long, colIndex As Long, currentDate As Date
    Dim isHidden As Boolean
    
    ' 最初の土日列の状態を確認
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        If Weekday(currentDate, vbMonday) >= 6 Then
            isHidden = (ws.Columns(colIndex).ColumnWidth = 0)
            Exit For
        End If
    Next i
    
    ' 土日列の幅を切り替え
    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1
        If Weekday(currentDate, vbMonday) >= 6 Then
            If isHidden Then
                ws.Columns(colIndex).ColumnWidth = 3
            Else
                ws.Columns(colIndex).ColumnWidth = 0
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "土日切替エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  書式リセット
' ==========================================
Sub ResetFormatting()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    
    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        ganttStartDate = Date
    End If
    
    Dim ganttStartCol As Long
    ganttStartCol = Columns(COL_GANTT_START).Column
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 罫線を再適用
    ApplyGanttBorders ws, lastRow
    DrawWeekSeparators ws, lastRow
    ApplyWeekendColors ws, lastRow, ganttStartDate, ganttStartCol
    ApplyDataValidationAndFormats ws, lastRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "書式リセット完了！", vbInformation, "リセット"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "書式リセットエラー: " & Err.Description, vbCritical, "エラー"
End Sub
