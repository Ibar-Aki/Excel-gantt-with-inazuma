Attribute VB_Name = "WBSParentRollup"
Option Explicit

Public Const ALERT_MARK_TODAY As String = "!"
Public Const ALERT_MARK_DELAY As String = "!!"
Public Const ALERT_COLOR_RED As Long = 255

Private Function GetRollupWorksheet(ByVal ws As Worksheet, ByVal operationName As String) As Worksheet
    If Not ws Is Nothing Then
        Set GetRollupWorksheet = ws
        Exit Function
    End If

    On Error Resume Next
    Set GetRollupWorksheet = ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME)
    On Error GoTo 0

    If GetRollupWorksheet Is Nothing Then
        MsgBox "メインシート '" & InazumaGantt_v3.MAIN_SHEET_NAME & "' が見つかりません。", vbExclamation, operationName
    End If
End Function

Private Function IsMarkerOnlyText(ByVal textValue As String) As Boolean
    textValue = Trim$(textValue)
    IsMarkerOnlyText = (textValue = ALERT_MARK_TODAY Or textValue = ALERT_MARK_DELAY)
End Function

Private Function HasTaskName(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    HasTaskName = InazumaGantt_v3.HasTaskContentInRow(ws, targetRow)
End Function

Private Function GetHierarchyLevel(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    If IsNumeric(ws.Cells(targetRow, InazumaGantt_v3.COL_HIERARCHY).Value) Then
        GetHierarchyLevel = CLng(ws.Cells(targetRow, InazumaGantt_v3.COL_HIERARCHY).Value)
    End If
End Function

Private Function HasChildTaskRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal currentLevel As Long) As Boolean
    Dim r As Long
    Dim rowLevel As Long

    For r = startRow + 1 To endRow
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > 0 Then
            If rowLevel <= currentLevel Then Exit Function
            If HasTaskName(ws, r) Then
                HasChildTaskRows = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function FindSubtreeEndRow(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal lastRow As Long) As Long
    Dim targetLevel As Long
    Dim r As Long
    Dim rowLevel As Long

    targetLevel = GetHierarchyLevel(ws, targetRow)
    FindSubtreeEndRow = targetRow

    If targetLevel <= 0 Then Exit Function

    For r = targetRow + 1 To lastRow
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > 0 Then
            If rowLevel <= targetLevel Then
                FindSubtreeEndRow = r - 1
                Exit Function
            End If
        End If
    Next r

    FindSubtreeEndRow = lastRow
End Function

Private Function FindParentTaskRow(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    Dim currentLevel As Long
    Dim r As Long
    Dim rowLevel As Long

    currentLevel = GetHierarchyLevel(ws, targetRow)
    If currentLevel <= 1 Then Exit Function

    For r = targetRow - 1 To InazumaGantt_v3.ROW_DATA_START Step -1
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > 0 Then
            If rowLevel < currentLevel And HasTaskName(ws, r) Then
                FindParentTaskRow = r
                Exit Function
            End If
        End If
    Next r
End Function

Private Function TryParseDevelopmentHoursLocal(ByVal hoursValue As Variant, ByRef normalizedHours As Double) As Boolean
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

    TryParseDevelopmentHoursLocal = True
End Function

Private Function SubtreeHasLeafHours(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long) As Boolean
    Dim r As Long
    Dim rowLevel As Long
    Dim hoursValue As Double

    For r = startRow + 1 To endRow
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > GetHierarchyLevel(ws, startRow) And HasTaskName(ws, r) Then
            If Not HasChildTaskRows(ws, r, endRow, rowLevel) Then
                If TryParseDevelopmentHoursLocal(ws.Cells(r, InazumaGantt_v3.COL_DEV_LT).Value, hoursValue) Then
                    SubtreeHasLeafHours = True
                    Exit Function
                End If
            End If
        End If
    Next r
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

Private Sub SetDateCell(ByVal targetCell As Range, ByVal dateValue As Variant)
    If IsDate(dateValue) Then
        targetCell.Value = CDate(dateValue)
    Else
        targetCell.ClearContents
    End If
End Sub

Private Function DetermineParentStatus(ByVal allComplete As Boolean, ByVal anyInProgress As Boolean, _
                                       ByVal anyOverdueIncomplete As Boolean, ByVal allFutureOnly As Boolean) As String
    If anyOverdueIncomplete Then
        DetermineParentStatus = "遅延"
    ElseIf allComplete Then
        DetermineParentStatus = "完了"
    ElseIf anyInProgress Then
        DetermineParentStatus = "進行中"
    ElseIf allFutureOnly Then
        DetermineParentStatus = "予定前"
    Else
        DetermineParentStatus = "未着手"
    End If
End Function

Private Sub RecalculateParentRow(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date)
    Dim targetLevel As Long
    Dim lastRow As Long
    Dim endRow As Long
    Dim r As Long
    Dim rowLevel As Long
    Dim leafCount As Long
    Dim progressSum As Double
    Dim weightedProgress As Double
    Dim totalHours As Double
    Dim hasHours As Boolean
    Dim rowProgress As Double
    Dim hoursValue As Double
    Dim rowStatus As String
    Dim isComplete As Boolean
    Dim allComplete As Boolean
    Dim anyInProgress As Boolean
    Dim anyOverdueIncomplete As Boolean
    Dim allFutureOnly As Boolean
    Dim planStart As Variant
    Dim planEnd As Variant
    Dim progressValue As Double
    Dim childEndRow As Long
    Dim targetManualHours As Double
    Dim subtreeHasHours As Boolean

    targetLevel = GetHierarchyLevel(ws, targetRow)
    If targetLevel <= 0 Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(ws)
    endRow = FindSubtreeEndRow(ws, targetRow, lastRow)
    If endRow <= targetRow Then Exit Sub
    If Not HasChildTaskRows(ws, targetRow, endRow, targetLevel) Then Exit Sub

    allComplete = True
    allFutureOnly = True

    For r = targetRow + 1 To endRow
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > targetLevel And HasTaskName(ws, r) Then
            If Not HasChildTaskRows(ws, r, endRow, rowLevel) Then
                leafCount = leafCount + 1

                rowProgress = InazumaGantt_v3.NormalizeProgressValue(ws.Cells(r, InazumaGantt_v3.COL_PROGRESS).Value, 0)
                progressSum = progressSum + rowProgress

                If TryParseDevelopmentHoursLocal(ws.Cells(r, InazumaGantt_v3.COL_DEV_LT).Value, hoursValue) Then
                    totalHours = totalHours + hoursValue
                    weightedProgress = weightedProgress + (rowProgress * hoursValue)
                    hasHours = True
                End If

                UpdateMinDate planStart, ws.Cells(r, InazumaGantt_v3.COL_START_PLAN).Value
                UpdateMaxDate planEnd, ws.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value
                rowStatus = Trim$(CStr(ws.Cells(r, InazumaGantt_v3.COL_STATUS).Value))
                isComplete = (rowStatus = "完了" Or rowProgress >= 1)

                If Not isComplete Then
                    allComplete = False

                    If rowStatus = "進行中" Or rowProgress > 0 Then
                        anyInProgress = True
                    End If

                    If IsDate(ws.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value) Then
                        If CDate(ws.Cells(r, InazumaGantt_v3.COL_END_PLAN).Value) < referenceDate Then
                            anyOverdueIncomplete = True
                        End If
                    End If

                    If Not (IsDate(ws.Cells(r, InazumaGantt_v3.COL_START_PLAN).Value) And _
                            CDate(ws.Cells(r, InazumaGantt_v3.COL_START_PLAN).Value) > referenceDate) Then
                        allFutureOnly = False
                    End If
                End If
            ElseIf targetLevel = 1 And rowLevel = 2 Then
                childEndRow = FindSubtreeEndRow(ws, r, lastRow)
                If Not SubtreeHasLeafHours(ws, r, childEndRow) Then
                    If TryParseDevelopmentHoursLocal(ws.Cells(r, InazumaGantt_v3.COL_DEV_LT).Value, hoursValue) Then
                        rowProgress = InazumaGantt_v3.NormalizeProgressValue(ws.Cells(r, InazumaGantt_v3.COL_PROGRESS).Value, 0)
                        totalHours = totalHours + hoursValue
                        weightedProgress = weightedProgress + (rowProgress * hoursValue)
                        hasHours = True
                    End If
                End If
            End If
        End If
    Next r

    If leafCount = 0 Then Exit Sub

    If hasHours And totalHours > 0 Then
        progressValue = weightedProgress / totalHours
    Else
        progressValue = progressSum / leafCount
    End If

    If allComplete Then progressValue = 1
    If progressValue < 0 Then progressValue = 0
    If progressValue > 1 Then progressValue = 1

    If targetLevel = 2 Then
        subtreeHasHours = SubtreeHasLeafHours(ws, targetRow, endRow)
        If Not subtreeHasHours Then
            If TryParseDevelopmentHoursLocal(ws.Cells(targetRow, InazumaGantt_v3.COL_DEV_LT).Value, targetManualHours) Then
                totalHours = targetManualHours
            End If
        End If
    End If

    ws.Cells(targetRow, InazumaGantt_v3.COL_STATUS).Value = DetermineParentStatus(allComplete, anyInProgress, anyOverdueIncomplete, allFutureOnly)
    ws.Cells(targetRow, InazumaGantt_v3.COL_PROGRESS).Value = progressValue
    ws.Cells(targetRow, InazumaGantt_v3.COL_DEV_LT).Value = totalHours
    ws.Cells(targetRow, InazumaGantt_v3.COL_DEV_LT).NumberFormat = InazumaGantt_v3.DEV_HOURS_NUMBER_FORMAT

    SetDateCell ws.Cells(targetRow, InazumaGantt_v3.COL_START_PLAN), planStart
    SetDateCell ws.Cells(targetRow, InazumaGantt_v3.COL_END_PLAN), planEnd
End Sub

Public Sub RecalculateTaskRowAndAncestors(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim workingWs As Worksheet
    Dim currentRow As Long
    Dim lastRow As Long
    Dim currentLevel As Long
    Dim referenceDate As Date

    Set workingWs = GetRollupWorksheet(ws, "親タスク再計算")
    If workingWs Is Nothing Then Exit Sub
    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    referenceDate = Date
    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    currentRow = targetRow

    currentLevel = GetHierarchyLevel(workingWs, currentRow)
    If currentLevel > 0 Then
        If HasChildTaskRows(workingWs, currentRow, FindSubtreeEndRow(workingWs, currentRow, lastRow), currentLevel) Then
            RecalculateParentRow workingWs, currentRow, referenceDate
        End If
    End If

    currentRow = FindParentTaskRow(workingWs, targetRow)
    Do While currentRow >= InazumaGantt_v3.ROW_DATA_START
        RecalculateParentRow workingWs, currentRow, referenceDate
        currentRow = FindParentTaskRow(workingWs, currentRow)
    Loop
End Sub

Public Sub RefreshAllParentTasks(ByVal ws As Worksheet)
    Dim workingWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rowLevel As Long

    Set workingWs = GetRollupWorksheet(ws, "親タスク再計算")
    If workingWs Is Nothing Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    For r = lastRow To InazumaGantt_v3.ROW_DATA_START Step -1
        rowLevel = GetHierarchyLevel(workingWs, r)
        If rowLevel > 0 And HasTaskName(workingWs, r) Then
            If HasChildTaskRows(workingWs, r, FindSubtreeEndRow(workingWs, r, lastRow), rowLevel) Then
                RecalculateParentRow workingWs, r, Date
            End If
        End If
    Next r
End Sub

Private Function IsThisWeek(ByVal targetDate As Date, ByVal referenceDate As Date) As Boolean
    Dim weekStart As Date
    Dim weekEnd As Date

    weekStart = referenceDate - (Weekday(referenceDate, vbMonday) - 1)
    weekEnd = weekStart + 6
    IsThisWeek = (targetDate >= weekStart And targetDate <= weekEnd)
End Function

Private Function DetermineAlertMarker(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date) As String
    Dim rowLevel As Long
    Dim statusText As String
    Dim progressValue As Double
    Dim startPlan As Variant
    Dim endPlan As Variant

    rowLevel = GetHierarchyLevel(ws, targetRow)
    If rowLevel <= 1 Then Exit Function
    If Not InazumaGantt_v3.HasTaskContentInRow(ws, targetRow) Then Exit Function

    statusText = Trim$(CStr(ws.Cells(targetRow, InazumaGantt_v3.COL_STATUS).Value))
    progressValue = InazumaGantt_v3.NormalizeProgressValue(ws.Cells(targetRow, InazumaGantt_v3.COL_PROGRESS).Value, 0)
    If statusText = "完了" Or progressValue >= 1 Then Exit Function

    startPlan = ws.Cells(targetRow, InazumaGantt_v3.COL_START_PLAN).Value
    endPlan = ws.Cells(targetRow, InazumaGantt_v3.COL_END_PLAN).Value

    If IsDate(endPlan) Then
        If CDate(endPlan) < referenceDate Then
            DetermineAlertMarker = ALERT_MARK_DELAY
            Exit Function
        End If
        If CDate(endPlan) = referenceDate Then
            DetermineAlertMarker = ALERT_MARK_TODAY
            Exit Function
        End If
    End If

    If IsDate(startPlan) Then
        If CDate(startPlan) = referenceDate Then
            DetermineAlertMarker = ALERT_MARK_TODAY
        End If
    End If
End Function

Private Sub ApplyAlertMarkerState(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date)
    Dim rowLevel As Long
    Dim markerText As String
    Dim markerCell As Range
    Dim currentValue As String
    Dim nextValue As String

    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    rowLevel = GetHierarchyLevel(ws, targetRow)
    Set markerCell = ws.Cells(targetRow, "C")
    currentValue = Trim$(CStr(markerCell.Value))

    If rowLevel <= 1 Then
        If IsMarkerOnlyText(currentValue) Then
            markerCell.ClearContents
        End If
        InazumaGantt_v3.RefreshTaskLabelPresentation ws, targetRow
        Exit Sub
    End If

    markerText = DetermineAlertMarker(ws, targetRow, referenceDate)
    If InazumaGantt_v3.HasPrimaryTaskContentInRow(ws, targetRow) Then
        nextValue = markerText
    ElseIf InazumaGantt_v3.HasTaskContentInRow(ws, targetRow) Then
        nextValue = InazumaGantt_v3.BuildAuxiliaryDisplayText(markerText)
    End If

    If currentValue <> nextValue Then
        markerCell.Value = nextValue
    End If

    InazumaGantt_v3.RefreshTaskLabelPresentation ws, targetRow
End Sub

Public Sub RefreshTaskAlertMarkers(ByVal ws As Worksheet)
    Dim workingWs As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim referenceDate As Date

    Set workingWs = GetRollupWorksheet(ws, "タスク強調表示")
    If workingWs Is Nothing Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    referenceDate = Date

    For r = InazumaGantt_v3.ROW_DATA_START To lastRow
        ApplyAlertMarkerState workingWs, r, referenceDate
    Next r
End Sub

Public Sub RefreshTaskAlertMarkersForRowAndAncestors(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim workingWs As Worksheet
    Dim currentRow As Long
    Dim referenceDate As Date

    Set workingWs = GetRollupWorksheet(ws, "タスク強調表示")
    If workingWs Is Nothing Then Exit Sub
    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    referenceDate = Date
    currentRow = targetRow
    ApplyAlertMarkerState workingWs, currentRow, referenceDate

    currentRow = FindParentTaskRow(workingWs, targetRow)
    Do While currentRow >= InazumaGantt_v3.ROW_DATA_START
        ApplyAlertMarkerState workingWs, currentRow, referenceDate
        currentRow = FindParentTaskRow(workingWs, currentRow)
    Loop
End Sub

Public Sub RefreshAllDerivedTaskData(ByVal ws As Worksheet)
    RefreshAllParentTasks ws
    RefreshTaskAlertMarkers ws
End Sub

Public Sub RecalculateAllParentTasks()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetRollupWorksheet(Nothing, "親タスク再計算")
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    RefreshAllDerivedTaskData ws
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "親タスク再計算エラー: " & Err.Description, vbCritical, "親タスク再計算"
End Sub
