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
Public Const CELL_PROJECT_START As String = "K3"
Public Const CELL_DISPLAY_WEEK As String = "K4"
Public Const CELL_TODAY As String = "M3"

' 色設定
Public Const COLOR_PLAN As Long = 16119285       ' RGB(245,245,245) 限りなく白に近い灰色
Public Const COLOR_PROGRESS As Long = 10921638   ' RGB(91,155,213) 青色
Public Const COLOR_HOLIDAY As Long = 15790320    ' RGB(240,240,240) 薄い灰色（休日祝日）
Public Const COLOR_ROW_BAND As Long = 16316664
Public Const COLOR_ACTUAL As Long = 5287936      ' RGB(0,176,80) 緑色
Public Const COLOR_TODAY As Long = 255           ' RGB(255,0,0) 赤
Public Const COLOR_WARN As Long = 13434879
Public Const COLOR_ERROR As Long = 13553151
Public Const COLOR_INAZUMA As Long = 42495       ' RGB(255,165,0) オレンジ
Public Const COLOR_HEADER_BG As Long = 12874308
Public Const COLOR_GANTT_HEADER As Long = 8421504
Public Const COLOR_WEEKEND As Long = 15790320    ' RGB(240,240,240) 薄い灰色
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
    
    ' ヘッダー設定 (ROW_HEADER = 8行目に統一)
    ws.Range(COL_HIERARCHY & ROW_HEADER).Value = "LV"
    ws.Range(COL_NO & ROW_HEADER).Value = "No."
    ws.Range("C" & ROW_HEADER).Value = "TASK(LV1)"
    ws.Range("D" & ROW_HEADER).Value = "TASK(LV2)"
    ws.Range("E" & ROW_HEADER).Value = "TASK(LV3)"
    ws.Range("F" & ROW_HEADER).Value = "TASK(LV4)"
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
    
    wsGuide.Cells(1, 1).Value = "InazumaGantt 説明"
    wsGuide.Cells(1, 1).Font.Bold = True
    wsGuide.Cells(3, 1).Value = "1) SetupInazumaGantt を実行して初期設定"
    wsGuide.Cells(4, 1).Value = "2) タスクを入力（C-F列）"
    wsGuide.Cells(5, 1).Value = "3) RefreshInazumaGantt を実行してガント更新"
    wsGuide.Columns(1).ColumnWidth = 50
End Sub

' ==========================================
'  ガント全体の罫線
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
                    barHeight = 10  ' 予定バーの高さ
                    
                    ' 予定バー（薄い灰色 + 黒枠線）
                    Set shp = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, cellWidth, barHeight)
                    shp.Name = "Bar_Plan_" & r
                    shp.Fill.ForeColor.RGB = COLOR_PLAN
                    shp.Line.Visible = msoTrue
                    shp.Line.ForeColor.RGB = RGB(0, 0, 0)  ' 黒枠線
                    shp.Line.Weight = 1
                    
                    ' 進捗バー（青色）
                    If progress > 0 Then
                        progressCol = startCol + CLng((endCol - startCol + 1) * progress) - 1
                        If progressCol >= startCol Then
                            Dim progressWidth As Double
                            progressWidth = ws.Cells(r, progressCol).Left + ws.Cells(r, progressCol).Width - cellLeft
                            If progress >= 1 Then progressWidth = cellWidth
                            
                            Set shp = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, progressWidth, barHeight)
                            shp.Name = "Bar_Progress_" & r
                            shp.Fill.ForeColor.RGB = COLOR_PROGRESS
                            shp.Line.Visible = msoFalse
                            
                            ' イナズマ線用のポイントを記録
                            inazumaCount = inazumaCount + 1
                            inazumaPoints(inazumaCount, 1) = cellLeft + progressWidth
                            inazumaPoints(inazumaCount, 2) = cellTop + barHeight / 2
                        End If
                    Else
                        ' 進捗0%の場合も開始位置を記録
                        inazumaCount = inazumaCount + 1
                        inazumaPoints(inazumaCount, 1) = cellLeft
                        inazumaPoints(inazumaCount, 2) = cellTop + barHeight / 2
                    End If
                End If
            End If
        End If
        
        ' 実績バー（緑色の塗りつぶしバー、予定の下に配置）
        If IsDate(startActual) Then
            Dim actualEndDate As Date
            If IsDate(endActual) Then
                actualEndDate = CDate(endActual)
            Else
                actualEndDate = Date
            End If
            
            startCol = DateToColumn(ganttStartDate, CDate(startActual), ganttStartCol)
            endCol = DateToColumn(ganttStartDate, actualEndDate, ganttStartCol)
            
            If startCol >= ganttStartCol And startCol <= ganttStartCol + GANTT_DAYS - 1 Then
                If endCol > ganttStartCol + GANTT_DAYS - 1 Then endCol = ganttStartCol + GANTT_DAYS - 1
                If endCol >= startCol Then
                    Dim actualBarHeight As Double
                    actualBarHeight = 6  ' 実績バーの高さ（予定より細め）
                    cellTop = ws.Cells(r, startCol).Top + 14  ' 予定バーの下に配置
                    cellLeft = ws.Cells(r, startCol).Left
                    cellWidth = ws.Cells(r, endCol).Left + ws.Cells(r, endCol).Width - cellLeft
                    
                    Set shp = ws.Shapes.AddShape(msoShapeRectangle, cellLeft, cellTop, cellWidth, actualBarHeight)
                    shp.Name = "Bar_Actual_" & r
                    shp.Fill.ForeColor.RGB = COLOR_ACTUAL
                    shp.Line.Visible = msoFalse
                End If
            End If
        End If
    Next r
    
    ' 今日線を描画
    Dim todayCol As Long
    todayCol = DateToColumn(ganttStartDate, Date, ganttStartCol)
    
    If todayCol >= ganttStartCol And todayCol <= ganttStartCol + GANTT_DAYS - 1 Then
        Dim todayLeft As Double, todayTop As Double, todayBottom As Double
        todayLeft = ws.Cells(ROW_DATE_HEADER, todayCol).Left + ws.Cells(ROW_DATE_HEADER, todayCol).Width / 2
        todayTop = ws.Cells(ROW_DATE_HEADER, todayCol).Top
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
    
    Call DrawGanttBars
    
    MsgBox "イナズマガント更新完了！", vbInformation, "イナズマガント"
    Exit Sub
    
ErrorHandler:
    MsgBox "更新中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
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
