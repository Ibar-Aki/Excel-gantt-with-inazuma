Attribute VB_Name = "WBSDashboardViews"
Option Explicit

Private Const DASHBOARD_SHEET_NAME As String = "WBSダッシュボード"
Private Const FILTER_VIEW_SHEET_NAME As String = "WBSフィルタビュー"
Private Const ALERT_WINDOW_DAYS As Long = 7
Private Const MAX_ALERT_ITEMS As Long = 12

Private Function RequireMainWorksheet(ByVal operationName As String) As Worksheet
    On Error Resume Next
    Set RequireMainWorksheet = ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME)
    On Error GoTo 0

    If RequireMainWorksheet Is Nothing Then
        MsgBox "メインシート '" & InazumaGantt_v3.MAIN_SHEET_NAME & "' が見つかりません。", vbExclamation, operationName
    End If
End Function

Private Function GetTaskNameFromRow(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    If Trim$(CStr(ws.Cells(targetRow, "F").Value)) <> "" Then
        GetTaskNameFromRow = Trim$(CStr(ws.Cells(targetRow, "F").Value))
    ElseIf Trim$(CStr(ws.Cells(targetRow, "E").Value)) <> "" Then
        GetTaskNameFromRow = Trim$(CStr(ws.Cells(targetRow, "E").Value))
    ElseIf Trim$(CStr(ws.Cells(targetRow, "D").Value)) <> "" Then
        GetTaskNameFromRow = Trim$(CStr(ws.Cells(targetRow, "D").Value))
    Else
        GetTaskNameFromRow = Trim$(CStr(ws.Cells(targetRow, "C").Value))
    End If
End Function

Private Function GetHierarchyLevel(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    If IsNumeric(ws.Cells(targetRow, InazumaGantt_v3.COL_HIERARCHY).Value) Then
        GetHierarchyLevel = CLng(ws.Cells(targetRow, InazumaGantt_v3.COL_HIERARCHY).Value)
    End If
End Function

Private Function IsTaskRow(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    IsTaskRow = (GetTaskNameFromRow(ws, targetRow) <> "")
End Function

Private Function HasChildTaskRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal currentLevel As Long) As Boolean
    Dim r As Long
    Dim nextLevel As Variant

    For r = startRow + 1 To endRow
        nextLevel = ws.Cells(r, InazumaGantt_v3.COL_HIERARCHY).Value
        If IsNumeric(nextLevel) Then
            If CLng(nextLevel) <= currentLevel Then Exit Function
            HasChildTaskRows = True
            Exit Function
        End If
    Next r
End Function

Private Function IsLeafTaskRow(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal lastRow As Long) As Boolean
    If Not IsTaskRow(ws, targetRow) Then Exit Function
    IsLeafTaskRow = Not HasChildTaskRows(ws, targetRow, lastRow, GetHierarchyLevel(ws, targetRow))
End Function

Private Function GetPhaseNameForRow(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    Dim r As Long

    For r = targetRow To InazumaGantt_v3.ROW_DATA_START Step -1
        If GetHierarchyLevel(ws, r) = 1 And IsTaskRow(ws, r) Then
            GetPhaseNameForRow = GetTaskNameFromRow(ws, r)
            Exit Function
        End If
    Next r
End Function

Private Function TryParseDevelopmentHours(ByVal hoursValue As Variant, ByRef normalizedHours As Double) As Boolean
    Dim textValue As String

    If IsEmpty(hoursValue) Then Exit Function

    textValue = Trim$(CStr(hoursValue))
    If textValue = "" Then Exit Function

    textValue = Replace$(textValue, " ", "")
    textValue = Replace$(textValue, "H", "")
    textValue = Replace$(textValue, "h", "")
    textValue = Replace$(textValue, "ｈ", "")

    If Not IsNumeric(textValue) Then Exit Function

    normalizedHours = CDbl(textValue)
    If normalizedHours < 0 Then Exit Function

    TryParseDevelopmentHours = True
End Function

Private Function FormatHoursOrBlank(ByVal hourValue As Variant) As String
    If IsNumeric(hourValue) Then
        FormatHoursOrBlank = InazumaGantt_v3.FormatDevelopmentHours(CDbl(hourValue))
    Else
        FormatHoursOrBlank = ""
    End If
End Function

Private Function FormatDateOrBlank(ByVal candidateValue As Variant) As String
    If IsDate(candidateValue) Then
        FormatDateOrBlank = Format$(CDate(candidateValue), "yy/mm/dd")
    Else
        FormatDateOrBlank = ""
    End If
End Function

Private Function NormalizeOwner(ByVal ownerValue As Variant) As String
    NormalizeOwner = Trim$(CStr(ownerValue))
    If NormalizeOwner = "" Then NormalizeOwner = "(未設定)"
End Function

Private Function CalculateDelayDays(ByVal planEnd As Variant, ByVal progressValue As Double, ByVal actualEnd As Variant) As Long
    If Not IsDate(planEnd) Then Exit Function

    If progressValue >= 1 And IsDate(actualEnd) Then
        If CDate(actualEnd) > CDate(planEnd) Then
            CalculateDelayDays = CLng(CDate(actualEnd) - CDate(planEnd))
        End If
    ElseIf progressValue < 1 And Date > CDate(planEnd) Then
        CalculateDelayDays = CLng(Date - CDate(planEnd))
    End If
End Function

Private Function IsDueSoon(ByVal planEnd As Variant, ByVal progressValue As Double) As Boolean
    If Not IsDate(planEnd) Then Exit Function
    If progressValue >= 1 Then Exit Function

    IsDueSoon = (CDate(planEnd) >= Date And CDate(planEnd) <= Date + ALERT_WINDOW_DAYS)
End Function

Private Function ResolveStatusText(ByVal statusValue As Variant, ByVal progressValue As Double) As String
    ResolveStatusText = Trim$(CStr(statusValue))

    If ResolveStatusText = "" Then
        If progressValue >= 1 Then
            ResolveStatusText = "完了"
        ElseIf progressValue > 0 Then
            ResolveStatusText = "進行中"
        Else
            ResolveStatusText = "未着手"
        End If
    End If
End Function

Private Function DetermineViewCategory(ByVal statusText As String, ByVal planEnd As Variant, ByVal progressValue As Double, ByVal delayDays As Long) As String
    If delayDays > 0 Then
        DetermineViewCategory = "遅延"
    ElseIf IsDueSoon(planEnd, progressValue) Then
        DetermineViewCategory = "期限接近"
    ElseIf statusText = "保留" Then
        DetermineViewCategory = "保留"
    ElseIf progressValue >= 1 Or Left$(statusText, 2) = "完了" Then
        DetermineViewCategory = "完了"
    ElseIf progressValue > 0 Or statusText = "進行中" Then
        DetermineViewCategory = "進行中"
    Else
        DetermineViewCategory = "未着手"
    End If
End Function

Private Function GetRowTypeLabel(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal lastRow As Long) As String
    Dim rowLevel As Long
    rowLevel = GetHierarchyLevel(ws, targetRow)

    If rowLevel = 1 Then
        GetRowTypeLabel = "フェーズ"
    ElseIf HasChildTaskRows(ws, targetRow, lastRow, rowLevel) Then
        GetRowTypeLabel = "親タスク"
    Else
        GetRowTypeLabel = "末端タスク"
    End If
End Function

Private Function PrepareReportSheet(ByVal sheetName As String, ByVal reportTitle As String, ByVal reportSubtitle As String) As Worksheet
    Dim ws As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    ws.Range("A1").Value = reportTitle
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A2").Value = reportSubtitle
    ws.Range("A3").Value = "生成日時: " & Format$(Now, "yy/mm/dd hh:mm")
    ws.Range("A4").Value = "元データ: " & InazumaGantt_v3.MAIN_SHEET_NAME

    Set PrepareReportSheet = ws
End Function

Private Sub ApplyCategoryStyle(ByVal targetCell As Range, ByVal categoryText As String)
    targetCell.Font.Bold = True
    targetCell.HorizontalAlignment = xlCenter

    Select Case categoryText
        Case "遅延"
            targetCell.Interior.Color = RGB(255, 199, 206)
            targetCell.Font.Color = RGB(156, 0, 6)
        Case "期限接近"
            targetCell.Interior.Color = RGB(255, 235, 156)
            targetCell.Font.Color = RGB(156, 87, 0)
        Case "進行中"
            targetCell.Interior.Color = RGB(221, 235, 247)
            targetCell.Font.Color = RGB(31, 78, 121)
        Case "完了"
            targetCell.Interior.Color = RGB(226, 239, 218)
            targetCell.Font.Color = RGB(0, 97, 0)
        Case "保留"
            targetCell.Interior.Color = RGB(217, 217, 217)
            targetCell.Font.Color = RGB(89, 89, 89)
        Case Else
            targetCell.Interior.Color = RGB(242, 242, 242)
            targetCell.Font.Color = RGB(68, 68, 68)
    End Select
End Sub

Private Sub WriteMetricBlock(ByVal ws As Worksheet, ByVal labelCell As String, ByVal valueCell As String, _
                             ByVal labelText As String, ByVal valueText As Variant)
    ws.Range(labelCell).Value = labelText
    ws.Range(valueCell).Value = valueText
    ws.Range(labelCell).Font.Bold = True
    ws.Range(labelCell & ":" & valueCell).Interior.Color = RGB(248, 249, 250)
End Sub

Public Sub CreateWBSDashboardSheet()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Set wsMain = RequireMainWorksheet("ダッシュボード作成")
    If wsMain Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = InazumaGantt_v3.GetLastDataRow(wsMain)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then
        MsgBox "タスクデータが見つかりません。", vbInformation, "ダッシュボード"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wsDash As Worksheet
    Set wsDash = PrepareReportSheet(DASHBOARD_SHEET_NAME, "WBSダッシュボード", "全体状況、担当負荷、要注意タスクの一覧")

    Dim assigneeStats As Object
    Dim assigneeOrder As Collection
    Dim alertItems As Collection
    Set assigneeStats = CreateObject("Scripting.Dictionary")
    Set assigneeOrder = New Collection
    Set alertItems = New Collection

    Dim totalTasks As Long
    Dim completedTasks As Long
    Dim inProgressTasks As Long
    Dim notStartedTasks As Long
    Dim holdTasks As Long
    Dim delayedTasks As Long
    Dim dueSoonTasks As Long
    Dim totalHours As Double
    Dim remainingHours As Double
    Dim totalProgress As Double
    Dim progressCount As Long
    Dim weightedProgress As Double
    Dim weightedHours As Double

    Dim r As Long
    Dim taskName As String
    Dim phaseName As String
    Dim ownerName As String
    Dim statusText As String
    Dim progressValue As Double
    Dim delayDays As Long
    Dim categoryText As String
    Dim hoursValue As Double
    Dim statValues As Variant

    For r = InazumaGantt_v3.ROW_DATA_START To lastRow
        If Not IsLeafTaskRow(wsMain, r, lastRow) Then GoTo ContinueRow

        taskName = GetTaskNameFromRow(wsMain, r)
        If taskName = "" Then GoTo ContinueRow

        totalTasks = totalTasks + 1
        phaseName = GetPhaseNameForRow(wsMain, r)
        ownerName = NormalizeOwner(wsMain.Cells(r, InazumaGantt_v3.COL_ASSIGNEE).Value)
        progressValue = InazumaGantt_v3.NormalizeProgressValue(wsMain.Cells(r, InazumaGantt_v3.COL_PROGRESS).Value, 0)
        statusText = ResolveStatusText(wsMain.Cells(r, InazumaGantt_v3.COL_STATUS).Value, progressValue)
        delayDays = CalculateDelayDays(wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value, progressValue, wsMain.Cells(r, InazumaGantt_v3.COL_END_ACTUAL).Value)
        categoryText = DetermineViewCategory(statusText, wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value, progressValue, delayDays)

        Select Case categoryText
            Case "完了"
                completedTasks = completedTasks + 1
            Case "進行中"
                inProgressTasks = inProgressTasks + 1
            Case "保留"
                holdTasks = holdTasks + 1
            Case Else
                notStartedTasks = notStartedTasks + 1
        End Select

        If delayDays > 0 Then delayedTasks = delayedTasks + 1
        If categoryText = "期限接近" Then dueSoonTasks = dueSoonTasks + 1

        totalProgress = totalProgress + progressValue
        progressCount = progressCount + 1

        If TryParseDevelopmentHours(wsMain.Cells(r, InazumaGantt_v3.COL_DEV_LT).Value, hoursValue) Then
            totalHours = totalHours + hoursValue
            remainingHours = remainingHours + (hoursValue * (1 - progressValue))
            weightedProgress = weightedProgress + (progressValue * hoursValue)
            weightedHours = weightedHours + hoursValue
        Else
            hoursValue = 0
        End If

        If Not assigneeStats.Exists(ownerName) Then
            assigneeStats.Add ownerName, Array(0, 0, 0, 0, 0#, 0#, 0#, 0#)
            assigneeOrder.Add ownerName
        End If

        statValues = assigneeStats(ownerName)
        statValues(0) = CLng(statValues(0)) + 1
        If categoryText = "完了" Then statValues(1) = CLng(statValues(1)) + 1
        If categoryText = "進行中" Or categoryText = "期限接近" Then statValues(2) = CLng(statValues(2)) + 1
        If delayDays > 0 Then statValues(3) = CLng(statValues(3)) + 1
        statValues(4) = CDbl(statValues(4)) + hoursValue
        statValues(5) = CDbl(statValues(5)) + (hoursValue * (1 - progressValue))
        statValues(6) = CDbl(statValues(6)) + progressValue
        statValues(7) = CDbl(statValues(7)) + IIf(hoursValue > 0, hoursValue, 1)
        assigneeStats(ownerName) = statValues

        If (delayDays > 0 Or categoryText = "期限接近") And alertItems.Count < MAX_ALERT_ITEMS Then
            alertItems.Add Array(phaseName, taskName, ownerName, categoryText, _
                                 FormatDateOrBlank(wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value), _
                                 delayDays, progressValue, statusText)
        End If

ContinueRow:
    Next r

    If totalTasks = 0 Then
        Application.ScreenUpdating = True
        MsgBox "末端タスクが見つかりません。", vbInformation, "ダッシュボード"
        Exit Sub
    End If

    WriteMetricBlock wsDash, "A6", "B6", "総タスク", totalTasks
    WriteMetricBlock wsDash, "D6", "E6", "完了", completedTasks
    WriteMetricBlock wsDash, "G6", "H6", "進行中", inProgressTasks
    WriteMetricBlock wsDash, "J6", "K6", "遅延", delayedTasks
    WriteMetricBlock wsDash, "M6", "N6", "期限接近", dueSoonTasks

    WriteMetricBlock wsDash, "A8", "B8", "未着手", notStartedTasks
    WriteMetricBlock wsDash, "D8", "E8", "保留", holdTasks
    WriteMetricBlock wsDash, "G8", "H8", "総開発LT", FormatHoursOrBlank(totalHours)
    WriteMetricBlock wsDash, "J8", "K8", "残開発LT", FormatHoursOrBlank(remainingHours)
    WriteMetricBlock wsDash, "M8", "N8", "全体進捗", IIf(weightedHours > 0, weightedProgress / weightedHours, totalProgress / progressCount)
    wsDash.Range("N8").NumberFormat = "0%"

    wsDash.Range("A11:I11").Value = Array("No.", "担当", "担当タスク", "完了", "進行中", "遅延", "総開発LT", "残開発LT", "平均進捗")
    wsDash.Range("A11:I11").Font.Bold = True

    Dim rowIndex As Long
    Dim i As Long
    Dim assigneeName As String
    rowIndex = 12

    For i = 1 To assigneeOrder.Count
        assigneeName = CStr(assigneeOrder(i))
        statValues = assigneeStats(assigneeName)

        wsDash.Cells(rowIndex, "A").Value = i
        wsDash.Cells(rowIndex, "B").Value = assigneeName
        wsDash.Cells(rowIndex, "C").Value = CLng(statValues(0))
        wsDash.Cells(rowIndex, "D").Value = CLng(statValues(1))
        wsDash.Cells(rowIndex, "E").Value = CLng(statValues(2))
        wsDash.Cells(rowIndex, "F").Value = CLng(statValues(3))
        wsDash.Cells(rowIndex, "G").Value = FormatHoursOrBlank(statValues(4))
        wsDash.Cells(rowIndex, "H").Value = FormatHoursOrBlank(statValues(5))
        wsDash.Cells(rowIndex, "I").Value = CDbl(statValues(6)) / CDbl(statValues(7))
        wsDash.Cells(rowIndex, "I").NumberFormat = "0%"
        rowIndex = rowIndex + 1
    Next i

    wsDash.Range("K11:R11").Value = Array("フェーズ", "タスク", "担当", "分類", "完了予定", "遅延日数", "進捗率", "状況")
    wsDash.Range("K11:R11").Font.Bold = True

    rowIndex = 12
    Dim alertInfo As Variant
    For i = 1 To alertItems.Count
        alertInfo = alertItems(i)
        wsDash.Cells(rowIndex, "K").Value = alertInfo(0)
        wsDash.Cells(rowIndex, "L").Value = alertInfo(1)
        wsDash.Cells(rowIndex, "M").Value = alertInfo(2)
        wsDash.Cells(rowIndex, "N").Value = alertInfo(3)
        wsDash.Cells(rowIndex, "O").Value = alertInfo(4)
        wsDash.Cells(rowIndex, "P").Value = IIf(CLng(alertInfo(5)) = 0, "", alertInfo(5))
        wsDash.Cells(rowIndex, "Q").Value = CDbl(alertInfo(6))
        wsDash.Cells(rowIndex, "Q").NumberFormat = "0%"
        wsDash.Cells(rowIndex, "R").Value = alertInfo(7)
        ApplyCategoryStyle wsDash.Cells(rowIndex, "N"), CStr(alertInfo(3))
        rowIndex = rowIndex + 1
    Next i

    If alertItems.Count = 0 Then
        wsDash.Range("K12").Value = "要注意タスクはありません。"
    End If

    With wsDash
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 14
        .Columns("C:F").ColumnWidth = 10
        .Columns("G:H").ColumnWidth = 12
        .Columns("I").ColumnWidth = 10
        .Columns("K").ColumnWidth = 18
        .Columns("L").ColumnWidth = 24
        .Columns("M").ColumnWidth = 14
        .Columns("N").ColumnWidth = 12
        .Columns("O").ColumnWidth = 11
        .Columns("P").ColumnWidth = 9
        .Columns("Q").ColumnWidth = 9
        .Columns("R").ColumnWidth = 12
    End With

    If Application.Visible Then
        wsDash.Activate
        ActiveWindow.DisplayGridlines = False
        wsDash.Range("A12").Select
        ActiveWindow.FreezePanes = True
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = "WBSダッシュボードを更新しました。"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Application.DisplayAlerts Then
        MsgBox "ダッシュボード作成エラー: " & Err.Description, vbCritical, "ダッシュボード"
    Else
        Err.Raise Err.Number, "CreateWBSDashboardSheet", Err.Description
    End If
End Sub

Public Sub CreateWBSFilterViewSheet()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Set wsMain = RequireMainWorksheet("フィルタビュー作成")
    If wsMain Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = InazumaGantt_v3.GetLastDataRow(wsMain)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then
        MsgBox "タスクデータが見つかりません。", vbInformation, "フィルタビュー"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wsView As Worksheet
    Set wsView = PrepareReportSheet(FILTER_VIEW_SHEET_NAME, "WBSフィルタビュー", "担当、状態、期限観点でフィルタしやすい一覧ビュー")

    wsView.Range("A5").Value = "使い方: 5行目のフィルタから、担当・分類・状況・期限接近を絞り込めます。"
    wsView.Range("A6:R6").Value = Array("元行", "No.", "LV", "行種別", "フェーズ", "タスク名", "担当", "状況", "進捗率", "開発LT", "開始予定", "完了予定", "開始実績", "完了実績", "遅延日数", "期限接近", "分類", "タスク詳細")
    wsView.Range("A6:R6").Font.Bold = True

    Dim currentPhase As String
    Dim rowIndex As Long
    Dim r As Long
    Dim rowType As String
    Dim taskName As String
    Dim statusText As String
    Dim progressValue As Double
    Dim delayDays As Long
    Dim categoryText As String
    Dim dueSoonText As String

    rowIndex = 7
    For r = InazumaGantt_v3.ROW_DATA_START To lastRow
        If Not IsTaskRow(wsMain, r) Then GoTo ContinueFilterRow

        taskName = GetTaskNameFromRow(wsMain, r)
        If taskName = "" Then GoTo ContinueFilterRow

        If GetHierarchyLevel(wsMain, r) = 1 Then currentPhase = taskName

        progressValue = InazumaGantt_v3.NormalizeProgressValue(wsMain.Cells(r, InazumaGantt_v3.COL_PROGRESS).Value, 0)
        statusText = ResolveStatusText(wsMain.Cells(r, InazumaGantt_v3.COL_STATUS).Value, progressValue)
        delayDays = CalculateDelayDays(wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value, progressValue, wsMain.Cells(r, InazumaGantt_v3.COL_END_ACTUAL).Value)
        categoryText = DetermineViewCategory(statusText, wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value, progressValue, delayDays)
        dueSoonText = IIf(IsDueSoon(wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value, progressValue), "対象", "")
        rowType = GetRowTypeLabel(wsMain, r, lastRow)

        wsView.Cells(rowIndex, "A").Value = r
        wsView.Cells(rowIndex, "B").Value = wsMain.Cells(r, "B").Value
        wsView.Cells(rowIndex, "C").Value = GetHierarchyLevel(wsMain, r)
        wsView.Cells(rowIndex, "D").Value = rowType
        wsView.Cells(rowIndex, "E").Value = currentPhase
        wsView.Cells(rowIndex, "F").Value = taskName
        wsView.Cells(rowIndex, "G").Value = NormalizeOwner(wsMain.Cells(r, InazumaGantt_v3.COL_ASSIGNEE).Value)
        wsView.Cells(rowIndex, "H").Value = statusText
        wsView.Cells(rowIndex, "I").Value = progressValue
        wsView.Cells(rowIndex, "I").NumberFormat = "0%"
        wsView.Cells(rowIndex, "J").Value = Trim$(CStr(wsMain.Cells(r, InazumaGantt_v3.COL_DEV_LT).Value))
        wsView.Cells(rowIndex, "K").Value = FormatDateOrBlank(wsMain.Cells(r, InazumaGantt_v3.COL_START_PLAN).Value)
        wsView.Cells(rowIndex, "L").Value = FormatDateOrBlank(wsMain.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value)
        wsView.Cells(rowIndex, "M").Value = FormatDateOrBlank(wsMain.Cells(r, InazumaGantt_v3.COL_START_ACTUAL).Value)
        wsView.Cells(rowIndex, "N").Value = FormatDateOrBlank(wsMain.Cells(r, InazumaGantt_v3.COL_END_ACTUAL).Value)
        wsView.Cells(rowIndex, "O").Value = IIf(delayDays = 0, "", delayDays)
        wsView.Cells(rowIndex, "P").Value = dueSoonText
        wsView.Cells(rowIndex, "Q").Value = categoryText
        wsView.Cells(rowIndex, "R").Value = Trim$(CStr(wsMain.Cells(r, InazumaGantt_v3.COL_TASK_DETAIL).Value))
        ApplyCategoryStyle wsView.Cells(rowIndex, "Q"), categoryText

        rowIndex = rowIndex + 1

ContinueFilterRow:
    Next r

    If rowIndex = 7 Then
        Application.ScreenUpdating = True
        MsgBox "表示できるタスクが見つかりません。", vbInformation, "フィルタビュー"
        Exit Sub
    End If

    With wsView
        .Range("A6:R" & rowIndex - 1).AutoFilter
        .Columns("A").ColumnWidth = 7
        .Columns("B:C").ColumnWidth = 7
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 18
        .Columns("F").ColumnWidth = 24
        .Columns("G").ColumnWidth = 14
        .Columns("H").ColumnWidth = 12
        .Columns("I").ColumnWidth = 9
        .Columns("J").ColumnWidth = 10
        .Columns("K:N").ColumnWidth = 11
        .Columns("O").ColumnWidth = 9
        .Columns("P").ColumnWidth = 10
        .Columns("Q").ColumnWidth = 12
        .Columns("R").ColumnWidth = 28
    End With

    If Application.Visible Then
        wsView.Activate
        ActiveWindow.DisplayGridlines = False
        wsView.Range("A7").Select
        ActiveWindow.FreezePanes = True
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = "WBSフィルタビューを更新しました。"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    If Application.DisplayAlerts Then
        MsgBox "フィルタビュー作成エラー: " & Err.Description, vbCritical, "フィルタビュー"
    Else
        Err.Raise Err.Number, "CreateWBSFilterViewSheet", Err.Description
    End If
End Sub
