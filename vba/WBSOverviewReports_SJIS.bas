Attribute VB_Name = "WBSOverviewReports"
Option Explicit

Private Const SUMMARY_SHEET_NAME As String = "WBS全体サマリ"
Private Const ROADMAP_SHEET_NAME As String = "WBSロードマップ"

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

Private Function FindPhaseEndRow(ByVal ws As Worksheet, ByVal phaseRow As Long, ByVal lastRow As Long) As Long
    Dim r As Long

    FindPhaseEndRow = phaseRow
    For r = phaseRow + 1 To lastRow
        If GetHierarchyLevel(ws, r) = 1 And IsTaskRow(ws, r) Then
            FindPhaseEndRow = r - 1
            Exit Function
        End If
    Next r

    FindPhaseEndRow = lastRow
End Function

Private Sub UpdateMinDate(ByRef currentValue As Variant, ByVal candidateValue As Variant)
    If Not IsDate(candidateValue) Then Exit Sub

    If Not IsDate(currentValue) Then
        currentValue = CDate(candidateValue)
    ElseIf CDate(candidateValue) < CDate(currentValue) Then
        currentValue = CDate(candidateValue)
    End If
End Sub

Private Sub UpdateMaxDate(ByRef currentValue As Variant, ByVal candidateValue As Variant)
    If Not IsDate(candidateValue) Then Exit Sub

    If Not IsDate(currentValue) Then
        currentValue = CDate(candidateValue)
    ElseIf CDate(candidateValue) > CDate(currentValue) Then
        currentValue = CDate(candidateValue)
    End If
End Sub

Private Function FormatDateOrBlank(ByVal candidateValue As Variant) As String
    If IsDate(candidateValue) Then
        FormatDateOrBlank = Format$(CDate(candidateValue), "yy/mm/dd")
    Else
        FormatDateOrBlank = ""
    End If
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

Private Function EvaluatePhaseStatus(ByVal completedCount As Long, ByVal inProgressCount As Long, _
                                     ByVal pendingCount As Long, ByVal holdCount As Long, _
                                     ByVal phaseProgress As Double, ByVal planStart As Variant, _
                                     ByVal planEnd As Variant, ByVal actualEnd As Variant) As String
    If completedCount > 0 And completedCount = completedCount + inProgressCount + pendingCount Then
        If IsDate(planEnd) And IsDate(actualEnd) And CDate(actualEnd) > CDate(planEnd) Then
            EvaluatePhaseStatus = "完了(遅延)"
        Else
            EvaluatePhaseStatus = "完了"
        End If
        Exit Function
    End If

    If holdCount > 0 And holdCount = completedCount + inProgressCount + pendingCount Then
        EvaluatePhaseStatus = "保留"
        Exit Function
    End If

    If IsDate(planEnd) And CDate(planEnd) < Date And phaseProgress < 1 Then
        EvaluatePhaseStatus = "遅延"
    ElseIf phaseProgress > 0 Or inProgressCount > 0 Or completedCount > 0 Then
        EvaluatePhaseStatus = "進行中"
    ElseIf IsDate(planStart) And CDate(planStart) > Date Then
        EvaluatePhaseStatus = "予定前"
    Else
        EvaluatePhaseStatus = "未着手"
    End If
End Function

Private Function CollectPhaseMetrics(ByVal ws As Worksheet, ByVal phaseRow As Long, ByVal phaseEndRow As Long) As Variant
    ' 1:フェーズ名 2:担当 3:配下タスク数 4:完了 5:進行中 6:未着手+保留
    ' 7:進捗率 8:総開発LT 9:残開発LT 10:開始予定 11:完了予定
    ' 12:開始実績 13:完了実績 14:判定 15:遅延日数
    Dim metrics(1 To 15) As Variant
    Dim r As Long
    Dim leafCount As Long
    Dim completedCount As Long
    Dim inProgressCount As Long
    Dim pendingCount As Long
    Dim holdCount As Long
    Dim progressSum As Double
    Dim progressCount As Long
    Dim totalHours As Double
    Dim weightedProgress As Double
    Dim hasHours As Boolean
    Dim phaseProgress As Double
    Dim hoursValue As Double
    Dim rowStatus As String
    Dim rowProgress As Double
    Dim rowOwner As String
    Dim planStart As Variant
    Dim planEnd As Variant
    Dim actualStart As Variant
    Dim actualEnd As Variant
    Dim hasLeafRows As Boolean

    metrics(1) = GetTaskNameFromRow(ws, phaseRow)

    For r = phaseRow To phaseEndRow
        If IsTaskRow(ws, r) Then
            UpdateMinDate planStart, ws.Cells(r, InazumaGantt_v3.COL_START_PLAN).Value
            UpdateMaxDate planEnd, ws.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value
            UpdateMinDate actualStart, ws.Cells(r, InazumaGantt_v3.COL_START_ACTUAL).Value
            UpdateMaxDate actualEnd, ws.Cells(r, InazumaGantt_v3.COL_END_ACTUAL).Value

            rowOwner = Trim$(CStr(ws.Cells(r, InazumaGantt_v3.COL_ASSIGNEE).Value))
            If metrics(2) = "" And rowOwner <> "" Then metrics(2) = rowOwner
        End If
    Next r

    For r = phaseRow To phaseEndRow
        If IsTaskRow(ws, r) Then
            If Not HasChildTaskRows(ws, r, phaseEndRow, GetHierarchyLevel(ws, r)) Then
                hasLeafRows = True
                leafCount = leafCount + 1

                rowStatus = Trim$(CStr(ws.Cells(r, InazumaGantt_v3.COL_STATUS).Value))
                Select Case rowStatus
                    Case "完了"
                        completedCount = completedCount + 1
                    Case "進行中"
                        inProgressCount = inProgressCount + 1
                    Case "保留"
                        pendingCount = pendingCount + 1
                        holdCount = holdCount + 1
                    Case Else
                        pendingCount = pendingCount + 1
                End Select

                rowProgress = InazumaGantt_v3.NormalizeProgressValue(ws.Cells(r, InazumaGantt_v3.COL_PROGRESS).Value, 0)
                progressSum = progressSum + rowProgress
                progressCount = progressCount + 1

                If TryParseDevelopmentHours(ws.Cells(r, InazumaGantt_v3.COL_DEV_LT).Value, hoursValue) Then
                    totalHours = totalHours + hoursValue
                    weightedProgress = weightedProgress + (rowProgress * hoursValue)
                    hasHours = True
                End If
            End If
        End If
    Next r

    If Not hasLeafRows And IsTaskRow(ws, phaseRow) Then
        leafCount = 1
        rowStatus = Trim$(CStr(ws.Cells(phaseRow, InazumaGantt_v3.COL_STATUS).Value))
        Select Case rowStatus
            Case "完了"
                completedCount = 1
            Case "進行中"
                inProgressCount = 1
            Case "保留"
                pendingCount = 1
                holdCount = 1
            Case Else
                pendingCount = 1
        End Select

        rowProgress = InazumaGantt_v3.NormalizeProgressValue(ws.Cells(phaseRow, InazumaGantt_v3.COL_PROGRESS).Value, 0)
        progressSum = rowProgress
        progressCount = 1

        If TryParseDevelopmentHours(ws.Cells(phaseRow, InazumaGantt_v3.COL_DEV_LT).Value, hoursValue) Then
            totalHours = hoursValue
            weightedProgress = rowProgress * hoursValue
            hasHours = True
        End If
    End If

    If hasHours And totalHours > 0 Then
        phaseProgress = weightedProgress / totalHours
        metrics(8) = totalHours
        metrics(9) = totalHours * (1 - phaseProgress)
    ElseIf progressCount > 0 Then
        phaseProgress = progressSum / progressCount
    End If

    metrics(3) = leafCount
    metrics(4) = completedCount
    metrics(5) = inProgressCount
    metrics(6) = pendingCount
    metrics(7) = phaseProgress
    metrics(10) = planStart
    metrics(11) = planEnd
    metrics(12) = actualStart
    metrics(13) = actualEnd
    metrics(14) = EvaluatePhaseStatus(completedCount, inProgressCount, pendingCount, holdCount, phaseProgress, planStart, planEnd, actualEnd)

    If IsDate(planEnd) Then
        If completedCount > 0 And completedCount = leafCount And IsDate(actualEnd) Then
            If CDate(actualEnd) > CDate(planEnd) Then
                metrics(15) = CLng(CDate(actualEnd) - CDate(planEnd))
            Else
                metrics(15) = 0
            End If
        ElseIf phaseProgress < 1 And Date > CDate(planEnd) Then
            metrics(15) = CLng(Date - CDate(planEnd))
        Else
            metrics(15) = 0
        End If
    Else
        metrics(15) = ""
    End If

    CollectPhaseMetrics = metrics
End Function

Private Function BuildPhaseMetricsCollection(ByVal ws As Worksheet) As Collection
    Dim phases As Collection
    Set phases = New Collection

    Dim lastRow As Long
    Dim phaseRow As Long
    Dim phaseEndRow As Long

    lastRow = InazumaGantt_v3.GetLastDataRow(ws)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then
        Set BuildPhaseMetricsCollection = phases
        Exit Function
    End If

    phaseRow = InazumaGantt_v3.ROW_DATA_START
    Do While phaseRow <= lastRow
        If GetHierarchyLevel(ws, phaseRow) = 1 And IsTaskRow(ws, phaseRow) Then
            phaseEndRow = FindPhaseEndRow(ws, phaseRow, lastRow)
            phases.Add CollectPhaseMetrics(ws, phaseRow, phaseEndRow)
            phaseRow = phaseEndRow + 1
        Else
            phaseRow = phaseRow + 1
        End If
    Loop

    Set BuildPhaseMetricsCollection = phases
End Function

Private Function PrepareReportSheet(ByVal sheetName As String, ByVal reportTitle As String) As Worksheet
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
    ws.Range("A2").Value = "生成日時: " & Format$(Now, "yy/mm/dd hh:mm")
    ws.Range("A3").Value = "元データ: " & InazumaGantt_v3.MAIN_SHEET_NAME

    Set PrepareReportSheet = ws
End Function

Private Function MonthStartDate(ByVal targetDate As Date) As Date
    MonthStartDate = DateSerial(Year(targetDate), Month(targetDate), 1)
End Function

Private Function ResolveRoadmapStartDate(ByVal metrics As Variant) As Variant
    If IsDate(metrics(10)) Then
        ResolveRoadmapStartDate = CDate(metrics(10))
    ElseIf IsDate(metrics(12)) Then
        ResolveRoadmapStartDate = CDate(metrics(12))
    End If
End Function

Private Function ResolveRoadmapEndDate(ByVal metrics As Variant) As Variant
    If IsDate(metrics(11)) Then
        ResolveRoadmapEndDate = CDate(metrics(11))
    ElseIf IsDate(metrics(13)) Then
        ResolveRoadmapEndDate = CDate(metrics(13))
    End If
End Function

Private Function ResolveRoadmapProgressEndDate(ByVal metrics As Variant) As Variant
    Dim planStart As Variant
    Dim planEnd As Variant
    Dim actualStart As Variant
    Dim actualEnd As Variant
    Dim progressValue As Double
    Dim spanDays As Long

    progressValue = 0
    If IsNumeric(metrics(7)) Then progressValue = CDbl(metrics(7))
    If progressValue <= 0 Then Exit Function

    planStart = metrics(10)
    planEnd = metrics(11)
    actualStart = metrics(12)
    actualEnd = metrics(13)

    If progressValue >= 1 And IsDate(actualEnd) Then
        ResolveRoadmapProgressEndDate = CDate(actualEnd)
        Exit Function
    End If

    If IsDate(planStart) And IsDate(planEnd) Then
        spanDays = CLng(CDate(planEnd) - CDate(planStart))
        ResolveRoadmapProgressEndDate = CDate(planStart) + CLng(spanDays * progressValue)
        If IsDate(actualStart) Then
            If CDate(ResolveRoadmapProgressEndDate) < CDate(actualStart) Then
                ResolveRoadmapProgressEndDate = CDate(actualStart)
            End If
        End If
    ElseIf IsDate(actualStart) Then
        ResolveRoadmapProgressEndDate = CDate(actualStart)
    ElseIf IsDate(planStart) Then
        ResolveRoadmapProgressEndDate = CDate(planStart)
    End If
End Function

Public Sub CreatePhaseSummarySheet()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Set wsMain = RequireMainWorksheet("全体サマリ作成")
    If wsMain Is Nothing Then Exit Sub

    Dim phases As Collection
    Set phases = BuildPhaseMetricsCollection(wsMain)
    If phases.Count = 0 Then
        MsgBox "LV1 フェーズが見つかりません。メインシートにタスクを入力してください。", vbInformation, "全体サマリ"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wsSummary As Worksheet
    Set wsSummary = PrepareReportSheet(SUMMARY_SHEET_NAME, SUMMARY_SHEET_NAME)

    Dim totalProgress As Double
    Dim totalHours As Double
    Dim remainingHours As Double
    Dim completedPhases As Long
    Dim delayedPhases As Long
    Dim rowIndex As Long
    Dim i As Long
    Dim metrics As Variant

    For i = 1 To phases.Count
        metrics = phases(i)
        If IsNumeric(metrics(8)) Then totalHours = totalHours + CDbl(metrics(8))
        If IsNumeric(metrics(9)) Then remainingHours = remainingHours + CDbl(metrics(9))
        If Left$(CStr(metrics(14)), 2) = "完了" Then completedPhases = completedPhases + 1
        If CStr(metrics(14)) = "遅延" Or CStr(metrics(14)) = "完了(遅延)" Then delayedPhases = delayedPhases + 1
        totalProgress = totalProgress + CDbl(metrics(7))
    Next i

    wsSummary.Range("A5").Value = "フェーズ数"
    wsSummary.Range("B5").Value = phases.Count
    wsSummary.Range("D5").Value = "完了フェーズ"
    wsSummary.Range("E5").Value = completedPhases
    wsSummary.Range("G5").Value = "遅延フェーズ"
    wsSummary.Range("H5").Value = delayedPhases
    wsSummary.Range("J5").Value = "全体進捗"
    wsSummary.Range("K5").Value = totalProgress / phases.Count
    wsSummary.Range("K5").NumberFormat = "0%"
    wsSummary.Range("M5").Value = "総開発LT"
    wsSummary.Range("N5").Value = FormatHoursOrBlank(totalHours)
    wsSummary.Range("P5").Value = "残開発LT"
    wsSummary.Range("Q5").Value = FormatHoursOrBlank(remainingHours)
    wsSummary.Range("A5:Q5").Font.Bold = True

    rowIndex = 8
    wsSummary.Range("A8:P8").Value = Array("No.", "フェーズ", "担当", "配下タスク", "完了", "進行中", "未着手/保留", "進捗率", "総開発LT", "残開発LT", "開始予定", "完了予定", "開始実績", "完了実績", "遅延日数", "判定")
    wsSummary.Range("A8:P8").Font.Bold = True

    rowIndex = 9
    For i = 1 To phases.Count
        metrics = phases(i)
        wsSummary.Cells(rowIndex, "A").Value = i
        wsSummary.Cells(rowIndex, "B").Value = metrics(1)
        wsSummary.Cells(rowIndex, "C").Value = metrics(2)
        wsSummary.Cells(rowIndex, "D").Value = metrics(3)
        wsSummary.Cells(rowIndex, "E").Value = metrics(4)
        wsSummary.Cells(rowIndex, "F").Value = metrics(5)
        wsSummary.Cells(rowIndex, "G").Value = metrics(6)
        wsSummary.Cells(rowIndex, "H").Value = CDbl(metrics(7))
        wsSummary.Cells(rowIndex, "H").NumberFormat = "0%"
        wsSummary.Cells(rowIndex, "I").Value = FormatHoursOrBlank(metrics(8))
        wsSummary.Cells(rowIndex, "J").Value = FormatHoursOrBlank(metrics(9))
        wsSummary.Cells(rowIndex, "K").Value = FormatDateOrBlank(metrics(10))
        wsSummary.Cells(rowIndex, "L").Value = FormatDateOrBlank(metrics(11))
        wsSummary.Cells(rowIndex, "M").Value = FormatDateOrBlank(metrics(12))
        wsSummary.Cells(rowIndex, "N").Value = FormatDateOrBlank(metrics(13))
        wsSummary.Cells(rowIndex, "O").Value = metrics(15)
        If wsSummary.Cells(rowIndex, "O").Value = 0 Then wsSummary.Cells(rowIndex, "O").Value = ""
        wsSummary.Cells(rowIndex, "P").Value = metrics(14)
        rowIndex = rowIndex + 1
    Next i

    With wsSummary
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 24
        .Columns("C").ColumnWidth = 12
        .Columns("D:G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 9
        .Columns("I:J").ColumnWidth = 11
        .Columns("K:N").ColumnWidth = 11
        .Columns("O").ColumnWidth = 9
        .Columns("P").ColumnWidth = 12
    End With

    If Application.Visible Then
        wsSummary.Activate
        ActiveWindow.DisplayGridlines = False
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = "WBS全体サマリを更新しました。"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "全体サマリ作成エラー: " & Err.Description, vbCritical, "全体サマリ"
End Sub

Public Sub CreateRoadmapOverviewSheet()
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Set wsMain = RequireMainWorksheet("ロードマップ作成")
    If wsMain Is Nothing Then Exit Sub

    Dim phases As Collection
    Set phases = BuildPhaseMetricsCollection(wsMain)
    If phases.Count = 0 Then
        MsgBox "LV1 フェーズが見つかりません。メインシートにタスクを入力してください。", vbInformation, "ロードマップ"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wsRoadmap As Worksheet
    Set wsRoadmap = PrepareReportSheet(ROADMAP_SHEET_NAME, ROADMAP_SHEET_NAME)

    Dim timelineStart As Variant
    Dim timelineEnd As Variant
    Dim i As Long
    Dim metrics As Variant
    Dim currentMonth As Date
    Dim monthCount As Long
    Dim monthCol As Long
    Dim rowIndex As Long
    Dim roadmapStart As Variant
    Dim roadmapEnd As Variant
    Dim progressEnd As Variant
    Dim planStartMonth As Long
    Dim planEndMonth As Long
    Dim progressMonth As Long
    Dim overrunStartMonth As Long
    Dim overrunEndMonth As Long

    For i = 1 To phases.Count
        metrics = phases(i)
        UpdateMinDate timelineStart, ResolveRoadmapStartDate(metrics)
        UpdateMaxDate timelineEnd, ResolveRoadmapEndDate(metrics)
        UpdateMaxDate timelineEnd, metrics(13)
    Next i

    If Not IsDate(timelineStart) Then
        If IsDate(wsMain.Range(InazumaGantt_v3.CELL_PROJECT_START).Value) Then
            timelineStart = CDate(wsMain.Range(InazumaGantt_v3.CELL_PROJECT_START).Value)
        Else
            timelineStart = Date
        End If
    End If

    If Not IsDate(timelineEnd) Then timelineEnd = DateAdd("m", 5, CDate(timelineStart))

    timelineStart = MonthStartDate(CDate(timelineStart))
    timelineEnd = MonthStartDate(CDate(timelineEnd))
    monthCount = DateDiff("m", CDate(timelineStart), CDate(timelineEnd)) + 1
    If monthCount < 6 Then monthCount = 6

    wsRoadmap.Range("A5").Value = "凡例:"
    wsRoadmap.Range("B5").Value = "■ 進捗済み"
    wsRoadmap.Range("D5").Value = "□ 残予定"
    wsRoadmap.Range("F5").Value = "■ 完了超過"

    wsRoadmap.Range("A7:F7").Value = Array("No.", "フェーズ", "進捗率", "開始予定", "完了予定", "判定")
    wsRoadmap.Range("A7:F7").Font.Bold = True

    currentMonth = MonthStartDate(Date)
    For i = 0 To monthCount - 1
        monthCol = 7 + i
        wsRoadmap.Cells(7, monthCol).Value = Format$(DateAdd("m", i, CDate(timelineStart)), "yy/mm")
        wsRoadmap.Cells(7, monthCol).HorizontalAlignment = xlCenter
        wsRoadmap.Cells(7, monthCol).Font.Bold = True
        wsRoadmap.Columns(monthCol).ColumnWidth = 6
        If DateAdd("m", i, CDate(timelineStart)) = currentMonth Then
            wsRoadmap.Cells(7, monthCol).Interior.Color = RGB(255, 242, 204)
        End If
    Next i

    rowIndex = 8
    For i = 1 To phases.Count
        metrics = phases(i)
        roadmapStart = ResolveRoadmapStartDate(metrics)
        roadmapEnd = ResolveRoadmapEndDate(metrics)
        progressEnd = ResolveRoadmapProgressEndDate(metrics)

        wsRoadmap.Cells(rowIndex, "A").Value = i
        wsRoadmap.Cells(rowIndex, "B").Value = metrics(1)
        wsRoadmap.Cells(rowIndex, "C").Value = CDbl(metrics(7))
        wsRoadmap.Cells(rowIndex, "C").NumberFormat = "0%"
        wsRoadmap.Cells(rowIndex, "D").Value = FormatDateOrBlank(metrics(10))
        wsRoadmap.Cells(rowIndex, "E").Value = FormatDateOrBlank(metrics(11))
        wsRoadmap.Cells(rowIndex, "F").Value = metrics(14)

        If IsDate(roadmapStart) Then
            planStartMonth = DateDiff("m", CDate(timelineStart), MonthStartDate(CDate(roadmapStart)))
        End If

        If IsDate(roadmapStart) And IsDate(roadmapEnd) Then
            planEndMonth = DateDiff("m", CDate(timelineStart), MonthStartDate(CDate(roadmapEnd)))
            For monthCol = 7 + planStartMonth To 7 + planEndMonth
                wsRoadmap.Cells(rowIndex, monthCol).Value = "□"
                wsRoadmap.Cells(rowIndex, monthCol).HorizontalAlignment = xlCenter
                wsRoadmap.Cells(rowIndex, monthCol).Font.Color = RGB(166, 166, 166)
            Next monthCol
        End If

        If IsDate(progressEnd) And IsDate(roadmapStart) Then
            progressMonth = DateDiff("m", CDate(timelineStart), MonthStartDate(CDate(progressEnd)))
            For monthCol = 7 + planStartMonth To 7 + progressMonth
                wsRoadmap.Cells(rowIndex, monthCol).Value = "■"
                wsRoadmap.Cells(rowIndex, monthCol).HorizontalAlignment = xlCenter
                If Left$(CStr(metrics(14)), 2) = "完了" Then
                    wsRoadmap.Cells(rowIndex, monthCol).Font.Color = RGB(0, 176, 80)
                Else
                    wsRoadmap.Cells(rowIndex, monthCol).Font.Color = InazumaGantt_v3.COLOR_PROGRESS
                End If
            Next monthCol
        End If

        If Left$(CStr(metrics(14)), 2) = "完了" And IsDate(metrics(11)) And IsDate(metrics(13)) Then
            If MonthStartDate(CDate(metrics(13))) > MonthStartDate(CDate(metrics(11))) Then
                overrunStartMonth = DateDiff("m", CDate(timelineStart), DateAdd("m", 1, MonthStartDate(CDate(metrics(11)))))
                overrunEndMonth = DateDiff("m", CDate(timelineStart), MonthStartDate(CDate(metrics(13))))
                For monthCol = 7 + overrunStartMonth To 7 + overrunEndMonth
                    wsRoadmap.Cells(rowIndex, monthCol).Value = "■"
                    wsRoadmap.Cells(rowIndex, monthCol).HorizontalAlignment = xlCenter
                    wsRoadmap.Cells(rowIndex, monthCol).Font.Color = RGB(237, 125, 49)
                Next monthCol
            End If
        End If

        rowIndex = rowIndex + 1
    Next i

    With wsRoadmap
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 24
        .Columns("C").ColumnWidth = 9
        .Columns("D:E").ColumnWidth = 11
        .Columns("F").ColumnWidth = 12
    End With

    If Application.Visible Then
        wsRoadmap.Activate
        ActiveWindow.DisplayGridlines = False
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = "WBSロードマップを更新しました。"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "ロードマップ作成エラー: " & Err.Description, vbCritical, "ロードマップ"
End Sub
