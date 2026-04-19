Attribute VB_Name = "WBSRoadmapReport"
Option Explicit

Private Const ROADMAP_SHEET_NAME As String = "WBSサマリ"
Private Const LEGACY_ROADMAP_SHEET_NAME As String = "WBSロードマップ"
Private Const ROADMAP_HEADER_FILL As Long = 11854022
Private Const ROADMAP_PROGRESS_COLOR As Long = 9851952
Private Const ROADMAP_COMPLETE_COLOR As Long = 5287936
Private Const ROADMAP_PLAN_COLOR As Long = 10921638
Private Const ROADMAP_OVERRUN_COLOR As Long = 3243501
Private Const ROADMAP_CELL_FONT As String = "MS Gothic"

Private Function RequireMainWorksheet(ByVal operationName As String) As Worksheet
    On Error Resume Next
    Set RequireMainWorksheet = ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME)
    On Error GoTo 0

    If RequireMainWorksheet Is Nothing Then
        MsgBox "メインシート '" & InazumaGantt_v3.MAIN_SHEET_NAME & "' が見つかりません。", vbExclamation, operationName
    End If
End Function

Private Function GetTaskNameFromRow(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    GetTaskNameFromRow = InazumaGantt_v3.GetVisibleTaskLabelForRow(ws, targetRow)
End Function

Private Function GetHierarchyLevel(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    If IsNumeric(ws.Cells(targetRow, InazumaGantt_v3.COL_HIERARCHY).Value) Then
        GetHierarchyLevel = CLng(ws.Cells(targetRow, InazumaGantt_v3.COL_HIERARCHY).Value)
    End If
End Function

Private Function IsTaskRow(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    IsTaskRow = InazumaGantt_v3.HasTaskContentInRow(ws, targetRow)
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

Private Function EvaluatePhaseStatus(ByVal completedCount As Long, ByVal inProgressCount As Long, _
                                     ByVal pendingCount As Long, ByVal holdCount As Long, _
                                     ByVal phaseProgress As Double, ByVal planStart As Variant, _
                                     ByVal planEnd As Variant, ByVal actualEnd As Variant, _
                                     ByVal referenceDate As Date) As String
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

    If IsDate(planEnd) And CDate(planEnd) < referenceDate And phaseProgress < 1 Then
        EvaluatePhaseStatus = "遅延"
    ElseIf phaseProgress > 0 Or inProgressCount > 0 Or completedCount > 0 Then
        EvaluatePhaseStatus = "進行中"
    ElseIf IsDate(planStart) And CDate(planStart) > referenceDate Then
        EvaluatePhaseStatus = "予定前"
    Else
        EvaluatePhaseStatus = "未着手"
    End If
End Function

Private Function CollectPhaseMetrics(ByVal ws As Worksheet, ByVal phaseRow As Long, ByVal phaseEndRow As Long, _
                                     ByVal referenceDate As Date) As Variant
    ' 1:フェーズ名 2:配下タスク数 3:進捗率 4:総開発LT 5:残開発LT
    ' 6:開始予定 7:完了予定 8:開始実績 9:完了実績 10:判定
    Dim metrics(1 To 10) As Variant
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
        metrics(4) = totalHours
        metrics(5) = totalHours * (1 - phaseProgress)
    ElseIf progressCount > 0 Then
        phaseProgress = progressSum / progressCount
    End If

    metrics(2) = leafCount
    metrics(3) = phaseProgress
    metrics(6) = planStart
    metrics(7) = planEnd
    metrics(8) = actualStart
    metrics(9) = actualEnd
    metrics(10) = EvaluatePhaseStatus(completedCount, inProgressCount, pendingCount, holdCount, phaseProgress, planStart, planEnd, actualEnd, referenceDate)

    CollectPhaseMetrics = metrics
End Function

Private Function BuildPhaseMetricsCollection(ByVal ws As Worksheet, ByVal referenceDate As Date) As Collection
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
            phases.Add CollectPhaseMetrics(ws, phaseRow, phaseEndRow, referenceDate)
            phaseRow = phaseEndRow + 1
        Else
            phaseRow = phaseRow + 1
        End If
    Loop

    Set BuildPhaseMetricsCollection = phases
End Function

Private Function PrepareRoadmapSheet() As Worksheet
    Dim ws As Worksheet
    Dim legacyWs As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ROADMAP_SHEET_NAME)
    Set legacyWs = ThisWorkbook.Worksheets(LEGACY_ROADMAP_SHEET_NAME)
    On Error GoTo 0

    If Not legacyWs Is Nothing Then
        If ws Is Nothing Or legacyWs.Name <> ws.Name Then
            Application.DisplayAlerts = False
            legacyWs.Delete
            Application.DisplayAlerts = True
        End If
    End If

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ROADMAP_SHEET_NAME
    Else
        ws.Cells.Clear
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
    End If

    Set PrepareRoadmapSheet = ws
End Function

Private Function MonthStartDate(ByVal targetDate As Date) As Date
    MonthStartDate = DateSerial(Year(targetDate), Month(targetDate), 1)
End Function

Private Function GetRoadmapEmptySlot() As String
    GetRoadmapEmptySlot = ChrW$(&H3000)
End Function

Private Function ResolveRoadmapStartDate(ByVal metrics As Variant) As Variant
    If IsDate(metrics(6)) Then
        ResolveRoadmapStartDate = CDate(metrics(6))
    ElseIf IsDate(metrics(8)) Then
        ResolveRoadmapStartDate = CDate(metrics(8))
    End If
End Function

Private Function ResolveRoadmapEndDate(ByVal metrics As Variant) As Variant
    If IsDate(metrics(7)) Then
        ResolveRoadmapEndDate = CDate(metrics(7))
    ElseIf IsDate(metrics(9)) Then
        ResolveRoadmapEndDate = CDate(metrics(9))
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
    If IsNumeric(metrics(3)) Then progressValue = CDbl(metrics(3))
    If progressValue <= 0 Then Exit Function

    planStart = metrics(6)
    planEnd = metrics(7)
    actualStart = metrics(8)
    actualEnd = metrics(9)

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

Private Function FormatHoursText(ByVal hoursValue As Variant) As String
    If IsNumeric(hoursValue) Then
        FormatHoursText = InazumaGantt_v3.FormatDevelopmentHours(CDbl(hoursValue))
    Else
        FormatHoursText = ""
    End If
End Function

Private Function FormatRoadmapDate(ByVal dateValue As Variant) As String
    If Not IsDate(dateValue) Then Exit Function
    FormatRoadmapDate = Format$(CDate(dateValue), "mm/dd") & "(" & GetWeekdayNameJa(CDate(dateValue)) & ")"
End Function

Private Function GetWeekdayNameJa(ByVal targetDate As Date) As String
    Select Case Weekday(targetDate, vbSunday)
        Case 1: GetWeekdayNameJa = "日"
        Case 2: GetWeekdayNameJa = "月"
        Case 3: GetWeekdayNameJa = "火"
        Case 4: GetWeekdayNameJa = "水"
        Case 5: GetWeekdayNameJa = "木"
        Case 6: GetWeekdayNameJa = "金"
        Case Else: GetWeekdayNameJa = "土"
    End Select
End Function

Private Function SlotOverlaps(ByVal slotStart As Date, ByVal slotEnd As Date, ByVal rangeStart As Date, ByVal rangeEnd As Date) As Boolean
    SlotOverlaps = (rangeStart <= slotEnd And rangeEnd >= slotStart)
End Function

Private Sub SetRoadmapMonthCell(ByVal targetCell As Range, ByVal planStart As Variant, ByVal planEnd As Variant, _
                                ByVal progressEnd As Variant, ByVal actualEnd As Variant, _
                                ByVal monthStart As Date, ByVal phaseStatus As String)
    Dim monthEnd As Date
    Dim firstHalfEnd As Date
    Dim secondHalfStart As Date
    Dim slotText(1 To 2) As String
    Dim slotColor(1 To 2) As Long
    Dim slotStart(1 To 2) As Date
    Dim slotEnd(1 To 2) As Date
    Dim i As Long
    Dim progressColor As Long
    Dim overrunStart As Date

    monthEnd = DateSerial(Year(monthStart), Month(monthStart) + 1, 0)
    firstHalfEnd = DateSerial(Year(monthStart), Month(monthStart), 15)
    secondHalfStart = DateSerial(Year(monthStart), Month(monthStart), 16)

    slotStart(1) = monthStart
    slotEnd(1) = firstHalfEnd
    slotStart(2) = secondHalfStart
    slotEnd(2) = monthEnd

    If Left$(phaseStatus, 2) = "完了" Then
        progressColor = ROADMAP_COMPLETE_COLOR
    Else
        progressColor = ROADMAP_PROGRESS_COLOR
    End If

    For i = 1 To 2
        slotText(i) = GetRoadmapEmptySlot()
        slotColor(i) = 0

        If IsDate(planStart) And IsDate(planEnd) Then
            If SlotOverlaps(slotStart(i), slotEnd(i), CDate(planStart), CDate(planEnd)) Then
                slotText(i) = "□"
                slotColor(i) = ROADMAP_PLAN_COLOR
            End If
        End If

        If IsDate(progressEnd) And IsDate(planStart) Then
            If SlotOverlaps(slotStart(i), slotEnd(i), CDate(planStart), CDate(progressEnd)) Then
                slotText(i) = "■"
                slotColor(i) = progressColor
            End If
        End If

        If Left$(phaseStatus, 2) = "完了" And IsDate(planEnd) And IsDate(actualEnd) Then
            If CDate(actualEnd) > CDate(planEnd) Then
                overrunStart = CDate(planEnd) + 1
                If SlotOverlaps(slotStart(i), slotEnd(i), overrunStart, CDate(actualEnd)) Then
                    slotText(i) = "■"
                    slotColor(i) = ROADMAP_OVERRUN_COLOR
                End If
            End If
        End If
    Next i

    targetCell.Value = slotText(1) & slotText(2)
    targetCell.NumberFormat = "@"
    targetCell.HorizontalAlignment = xlCenter
    targetCell.Font.Name = ROADMAP_CELL_FONT
    targetCell.Font.Color = RGB(0, 0, 0)
    targetCell.Characters(1, 1).Font.Color = slotColor(1)
    targetCell.Characters(2, 1).Font.Color = slotColor(2)
End Sub

Private Sub ApplyRoadmapTableStyle(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal lastRow As Long)
    With ws.Range(ws.Cells(5, 1), ws.Cells(5, lastCol))
        .Interior.Color = ROADMAP_HEADER_FILL
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range(ws.Cells(5, 1), ws.Cells(lastRow, lastCol))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End With

    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").HorizontalAlignment = xlLeft
    ws.Range("A2:A3").HorizontalAlignment = xlLeft
    ws.Range("A6:A" & lastRow).HorizontalAlignment = xlCenter
    ws.Range("C6:G" & lastRow).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(6, 8), ws.Cells(lastRow, lastCol)).HorizontalAlignment = xlCenter

    ws.Rows(1).RowHeight = 25.5
    ws.Rows(2).RowHeight = 18.75
    ws.Rows(3).RowHeight = 18.75
    ws.Rows(5).RowHeight = 19.5
    ws.Range("6:" & lastRow).RowHeight = 18.75
End Sub

Public Sub CreateRoadmapOverviewSheet(Optional ByVal referenceDate As Variant)
    On Error GoTo ErrorHandler

    Dim wsMain As Worksheet
    Set wsMain = RequireMainWorksheet("WBSサマリ作成")
    If wsMain Is Nothing Then Exit Sub

    Dim reportDate As Date
    If IsDate(referenceDate) Then
        reportDate = CDate(referenceDate)
    Else
        reportDate = Date
    End If

    Dim phases As Collection
    Set phases = BuildPhaseMetricsCollection(wsMain, reportDate)
    If phases.Count = 0 Then
        MsgBox "LV1 フェーズが見つかりません。メインシートにタスクを入力してください。", vbInformation, "WBSサマリ"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Dim wsRoadmap As Worksheet
    Set wsRoadmap = PrepareRoadmapSheet()

    Dim timelineStart As Variant
    Dim timelineEnd As Variant
    Dim i As Long
    Dim metrics As Variant
    Dim monthCount As Long
    Dim monthCol As Long
    Dim rowIndex As Long
    Dim currentMonth As Date
    Dim roadmapStart As Variant
    Dim roadmapEnd As Variant
    Dim progressEnd As Variant
    Dim planEndDate As Variant
    Dim totalHours As Double
    Dim remainingHours As Double
    Dim totalWeightedProgress As Double
    Dim lastCol As Long
    Dim totalRow As Long

    For i = 1 To phases.Count
        metrics = phases(i)
        UpdateMinDate timelineStart, ResolveRoadmapStartDate(metrics)
        UpdateMaxDate timelineEnd, ResolveRoadmapEndDate(metrics)
        UpdateMaxDate timelineEnd, metrics(9)
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

    wsRoadmap.Range("A1").Value = ROADMAP_SHEET_NAME
    wsRoadmap.Range("A2").Value = "生成日時: " & Format$(Now, "yy/mm/dd hh:mm")
    wsRoadmap.Range("A3").Value = "凡例:"
    wsRoadmap.Range("B3").Value = "■　進捗済み, □　残予定, 　■ 後半予定"
    wsRoadmap.Range("A5:G5").Value = Array("No.", "フェーズ", "進捗率", "総LT", "残LT", "完了日", "判定")

    currentMonth = MonthStartDate(reportDate)
    For i = 0 To monthCount - 1
        monthCol = 8 + i
        wsRoadmap.Cells(5, monthCol).Value = CStr(Month(DateAdd("m", i, CDate(timelineStart)))) & "月"
        wsRoadmap.Columns(monthCol).ColumnWidth = 7
        If MonthStartDate(DateAdd("m", i, CDate(timelineStart))) = currentMonth Then
            wsRoadmap.Cells(5, monthCol).Interior.Color = ROADMAP_HEADER_FILL
        End If
    Next i

    rowIndex = 6
    For i = 1 To phases.Count
        metrics = phases(i)
        roadmapStart = ResolveRoadmapStartDate(metrics)
        roadmapEnd = ResolveRoadmapEndDate(metrics)
        progressEnd = ResolveRoadmapProgressEndDate(metrics)
        planEndDate = metrics(7)

        wsRoadmap.Cells(rowIndex, "A").Value = i
        wsRoadmap.Cells(rowIndex, "B").Value = metrics(1)
        wsRoadmap.Cells(rowIndex, "C").Value = CDbl(metrics(3))
        wsRoadmap.Cells(rowIndex, "C").NumberFormat = "0%"
        wsRoadmap.Cells(rowIndex, "D").Value = FormatHoursText(metrics(4))
        wsRoadmap.Cells(rowIndex, "E").Value = FormatHoursText(metrics(5))
        wsRoadmap.Cells(rowIndex, "F").Value = FormatRoadmapDate(planEndDate)
        wsRoadmap.Cells(rowIndex, "G").Value = metrics(10)

        If IsNumeric(metrics(4)) Then
            totalHours = totalHours + CDbl(metrics(4))
            totalWeightedProgress = totalWeightedProgress + (CDbl(metrics(3)) * CDbl(metrics(4)))
        End If
        If IsNumeric(metrics(5)) Then remainingHours = remainingHours + CDbl(metrics(5))

        For monthCol = 8 To 7 + monthCount
            SetRoadmapMonthCell wsRoadmap.Cells(rowIndex, monthCol), roadmapStart, roadmapEnd, progressEnd, metrics(9), _
                DateAdd("m", monthCol - 8, CDate(timelineStart)), CStr(metrics(10))
        Next monthCol

        rowIndex = rowIndex + 1
    Next i

    totalRow = rowIndex
    wsRoadmap.Cells(totalRow, "A").Value = "合計"
    If totalHours > 0 Then
        wsRoadmap.Cells(totalRow, "C").Value = totalWeightedProgress / totalHours
        wsRoadmap.Cells(totalRow, "C").NumberFormat = "0%"
    End If
    wsRoadmap.Cells(totalRow, "D").Value = FormatHoursText(totalHours)
    wsRoadmap.Cells(totalRow, "E").Value = FormatHoursText(remainingHours)

    lastCol = 7 + monthCount

    With wsRoadmap
        .Columns("A").ColumnWidth = 3.5
        .Columns("B").ColumnWidth = 24
        .Columns("C").ColumnWidth = 6.8
        .Columns("D").ColumnWidth = 7.9
        .Columns("E").ColumnWidth = 9
        .Columns("F").ColumnWidth = 9.1
        .Columns("G").ColumnWidth = 6.5
    End With

    ApplyRoadmapTableStyle wsRoadmap, lastCol, totalRow

    If ThisWorkbook.Windows.Count > 0 Then
        wsRoadmap.Activate
        ActiveWindow.DisplayGridlines = False
    End If

    Application.ScreenUpdating = True
    Application.StatusBar = "WBSサマリを更新しました。"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "WBSサマリ作成エラー: " & Err.Description, vbCritical, "WBSサマリ"
End Sub
