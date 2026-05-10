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
    Dim actualStart As Variant
    Dim actualEnd As Variant
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
                UpdateMinDate actualStart, ws.Cells(r, InazumaGantt_v3.COL_START_ACTUAL).Value
                UpdateMaxDate actualEnd, ws.Cells(r, InazumaGantt_v3.COL_END_ACTUAL).Value

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
    SetDateCell ws.Cells(targetRow, InazumaGantt_v3.COL_START_ACTUAL), actualStart
    SetDateCell ws.Cells(targetRow, InazumaGantt_v3.COL_END_ACTUAL), actualEnd
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

Private Function GetMaxAffectedRow(ByVal affectedRows As Object, ByVal defaultMaxRow As Long) As Long
    Dim rowKey As Variant
    Dim targetRow As Long

    GetMaxAffectedRow = defaultMaxRow
    If affectedRows Is Nothing Then Exit Function

    For Each rowKey In affectedRows.Keys
        targetRow = CLng(rowKey)
        If targetRow > GetMaxAffectedRow Then GetMaxAffectedRow = targetRow
    Next rowKey
End Function

Private Function BuildParentRowMap(ByVal ws As Worksheet, ByVal maxRow As Long) As Object
    Const MAX_TRACKED_LEVEL As Long = 32

    Dim parentRows As Object
    Dim levelStack(1 To MAX_TRACKED_LEVEL) As Long
    Dim r As Long
    Dim rowLevel As Long
    Dim levelIndex As Long

    Set parentRows = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set BuildParentRowMap = parentRows
        Exit Function
    End If

    For r = InazumaGantt_v3.ROW_DATA_START To maxRow
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > 0 Then
            If rowLevel <= MAX_TRACKED_LEVEL Then
                For levelIndex = rowLevel - 1 To 1 Step -1
                    If levelStack(levelIndex) > 0 Then
                        parentRows(CStr(r)) = levelStack(levelIndex)
                        Exit For
                    End If
                Next levelIndex

                If HasTaskName(ws, r) Then
                    levelStack(rowLevel) = r
                Else
                    levelStack(rowLevel) = 0
                End If

                For levelIndex = rowLevel + 1 To MAX_TRACKED_LEVEL
                    levelStack(levelIndex) = 0
                Next levelIndex
            Else
                parentRows(CStr(r)) = FindParentTaskRow(ws, r)
            End If
        End If
    Next r

    Set BuildParentRowMap = parentRows
End Function

Private Sub AddRowAndAncestorKeys(ByVal targetRow As Long, ByVal parentRows As Object, ByVal rowKeys As Object)
    Dim currentRow As Long
    Dim currentKey As String

    If parentRows Is Nothing Then Exit Sub
    If rowKeys Is Nothing Then Exit Sub
    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    rowKeys(CStr(targetRow)) = True
    currentRow = targetRow

    Do
        currentKey = CStr(currentRow)
        If Not parentRows.Exists(currentKey) Then Exit Do
        currentRow = CLng(parentRows(currentKey))
        If currentRow < InazumaGantt_v3.ROW_DATA_START Then Exit Do
        rowKeys(CStr(currentRow)) = True
    Loop
End Sub

Public Sub RecalculateTaskRowsAndAncestors(ByVal ws As Worksheet, ByVal affectedRows As Object)
    Dim workingWs As Worksheet
    Dim rowKeys As Object
    Dim parentRows As Object
    Dim rowKey As Variant
    Dim r As Long
    Dim lastRow As Long
    Dim maxRow As Long
    Dim currentLevel As Long
    Dim referenceDate As Date

    Set workingWs = GetRollupWorksheet(ws, "親タスク再計算")
    If workingWs Is Nothing Then Exit Sub
    If affectedRows Is Nothing Then Exit Sub
    If affectedRows.Count = 0 Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    maxRow = GetMaxAffectedRow(affectedRows, lastRow)
    referenceDate = Date
    Set rowKeys = CreateObject("Scripting.Dictionary")
    Set parentRows = BuildParentRowMap(workingWs, maxRow)

    For Each rowKey In affectedRows.Keys
        AddRowAndAncestorKeys CLng(rowKey), parentRows, rowKeys
    Next rowKey

    For r = lastRow To InazumaGantt_v3.ROW_DATA_START Step -1
        If rowKeys.Exists(CStr(r)) Then
            currentLevel = GetHierarchyLevel(workingWs, r)
            If currentLevel > 0 Then
                If HasChildTaskRows(workingWs, r, FindSubtreeEndRow(workingWs, r, lastRow), currentLevel) Then
                    RecalculateParentRow workingWs, r, referenceDate
                End If
            End If
        End If
    Next r
End Sub

Private Function DataText(ByVal cellValue As Variant) As String
    If IsError(cellValue) Or IsEmpty(cellValue) Then Exit Function
    DataText = Trim$(CStr(cellValue))
End Function

Private Function HasTaskContentInData(ByRef data As Variant, ByVal rowIndex As Long) As Boolean
    Dim colIndex As Long
    Dim textValue As String

    If DataText(data(rowIndex, 6)) <> "" Then
        HasTaskContentInData = True
        Exit Function
    End If
    If DataText(data(rowIndex, 5)) <> "" Then
        HasTaskContentInData = True
        Exit Function
    End If
    If DataText(data(rowIndex, 4)) <> "" Then
        HasTaskContentInData = True
        Exit Function
    End If

    textValue = DataText(data(rowIndex, 3))
    If textValue <> "" And Not InazumaGantt_v3.IsAlertMarkerText(textValue) And _
       Not InazumaGantt_v3.IsAuxiliaryPlaceholderText(textValue) Then
        HasTaskContentInData = True
        Exit Function
    End If

    For colIndex = 7 To UBound(data, 2)
        If DataText(data(rowIndex, colIndex)) <> "" Then
            HasTaskContentInData = True
            Exit Function
        End If
    Next colIndex
End Function

Public Sub RefreshAllParentTasks(ByVal ws As Worksheet)
    Const MAX_TRACKED_LEVEL As Long = 32

    Dim workingWs As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim data As Variant
    Dim i As Long
    Dim levelIndex As Long
    Dim parentIndex As Long
    Dim referenceDate As Date
    Dim levels() As Long
    Dim parentRows() As Long
    Dim levelStack(1 To MAX_TRACKED_LEVEL) As Long
    Dim hasTask() As Boolean
    Dim hasChild() As Boolean
    Dim leafCount() As Long
    Dim progressSum() As Double
    Dim weightedProgress() As Double
    Dim totalHours() As Double
    Dim hasHours() As Boolean
    Dim leafHoursExists() As Boolean
    Dim allComplete() As Boolean
    Dim anyInProgress() As Boolean
    Dim anyOverdueIncomplete() As Boolean
    Dim allFutureOnly() As Boolean
    Dim planStart() As Variant
    Dim planEnd() As Variant
    Dim actualStart() As Variant
    Dim actualEnd() As Variant
    Dim rowProgress As Double
    Dim progressValue As Double
    Dim hoursValue As Double
    Dim rowStatus As String
    Dim isComplete As Boolean

    Set workingWs = GetRollupWorksheet(ws, "親タスク再計算")
    If workingWs Is Nothing Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    data = workingWs.Range("A" & InazumaGantt_v3.ROW_DATA_START & ":" & InazumaGantt_v3.COL_DATA_END & lastRow).Value2
    rowCount = UBound(data, 1)
    colCount = UBound(data, 2)
    referenceDate = Date

    ReDim levels(1 To rowCount)
    ReDim parentRows(1 To rowCount)
    ReDim hasTask(1 To rowCount)
    ReDim hasChild(1 To rowCount)
    ReDim leafCount(1 To rowCount)
    ReDim progressSum(1 To rowCount)
    ReDim weightedProgress(1 To rowCount)
    ReDim totalHours(1 To rowCount)
    ReDim hasHours(1 To rowCount)
    ReDim leafHoursExists(1 To rowCount)
    ReDim allComplete(1 To rowCount)
    ReDim anyInProgress(1 To rowCount)
    ReDim anyOverdueIncomplete(1 To rowCount)
    ReDim allFutureOnly(1 To rowCount)
    ReDim planStart(1 To rowCount)
    ReDim planEnd(1 To rowCount)
    ReDim actualStart(1 To rowCount)
    ReDim actualEnd(1 To rowCount)

    For i = 1 To rowCount
        allComplete(i) = True
        allFutureOnly(i) = True
        If IsNumeric(data(i, 1)) Then levels(i) = CLng(data(i, 1))
        hasTask(i) = HasTaskContentInData(data, i)

        If levels(i) > 0 Then
            If levels(i) <= MAX_TRACKED_LEVEL Then
                For levelIndex = levels(i) - 1 To 1 Step -1
                    If levelStack(levelIndex) > 0 Then
                        parentRows(i) = levelStack(levelIndex)
                        If hasTask(i) Then hasChild(parentRows(i)) = True
                        Exit For
                    End If
                Next levelIndex

                levelStack(levels(i)) = i
                For levelIndex = levels(i) + 1 To MAX_TRACKED_LEVEL
                    levelStack(levelIndex) = 0
                Next levelIndex
            End If
        End If
    Next i

    For i = rowCount To 1 Step -1
        If Not hasTask(i) Or levels(i) <= 0 Then GoTo ContinueRow

        If hasChild(i) Then
            If leafCount(i) = 0 Then GoTo ContinueRow

            If hasHours(i) And totalHours(i) > 0 Then
                progressValue = weightedProgress(i) / totalHours(i)
            Else
                progressValue = progressSum(i) / leafCount(i)
            End If

            If allComplete(i) Then progressValue = 1
            If progressValue < 0 Then progressValue = 0
            If progressValue > 1 Then progressValue = 1

            If levels(i) = 2 And Not leafHoursExists(i) Then
                If TryParseDevelopmentHoursLocal(data(i, 11), hoursValue) Then
                    totalHours(i) = hoursValue
                    weightedProgress(i) = progressValue * hoursValue
                    hasHours(i) = True
                End If
            End If

            workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, InazumaGantt_v3.COL_STATUS).Value = _
                DetermineParentStatus(allComplete(i), anyInProgress(i), anyOverdueIncomplete(i), allFutureOnly(i))
            workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, InazumaGantt_v3.COL_PROGRESS).Value = progressValue
            workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, InazumaGantt_v3.COL_DEV_LT).Value = totalHours(i)
            workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, InazumaGantt_v3.COL_DEV_LT).NumberFormat = InazumaGantt_v3.DEV_HOURS_NUMBER_FORMAT
            SetDateCell workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, InazumaGantt_v3.COL_START_PLAN), planStart(i)
            SetDateCell workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, InazumaGantt_v3.COL_END_PLAN), planEnd(i)
            If colCount >= 15 Then
                SetDateCell workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, 14), actualStart(i)
                SetDateCell workingWs.Cells(i + InazumaGantt_v3.ROW_DATA_START - 1, 15), actualEnd(i)
            End If
        Else
            leafCount(i) = 1
            rowProgress = InazumaGantt_v3.NormalizeProgressValue(data(i, 9), 0)
            progressSum(i) = rowProgress

            If TryParseDevelopmentHoursLocal(data(i, 11), hoursValue) Then
                totalHours(i) = hoursValue
                weightedProgress(i) = rowProgress * hoursValue
                hasHours(i) = True
                leafHoursExists(i) = True
            End If

            If IsDate(data(i, 12)) Then planStart(i) = CDate(data(i, 12))
            If IsDate(data(i, 13)) Then planEnd(i) = CDate(data(i, 13))
            If colCount >= 15 Then
                If IsDate(data(i, 14)) Then actualStart(i) = CDate(data(i, 14))
                If IsDate(data(i, 15)) Then actualEnd(i) = CDate(data(i, 15))
            End If

            rowStatus = DataText(data(i, 8))
            isComplete = (rowStatus = "完了" Or rowProgress >= 1)
            allComplete(i) = isComplete
            If Not isComplete Then
                If rowStatus = "進行中" Or rowProgress > 0 Then anyInProgress(i) = True
                If IsDate(data(i, 13)) Then
                    If CDate(data(i, 13)) < referenceDate Then anyOverdueIncomplete(i) = True
                End If
                allFutureOnly(i) = (IsDate(data(i, 12)) And CDate(data(i, 12)) > referenceDate)
            End If
        End If

        parentIndex = parentRows(i)
        If parentIndex > 0 And leafCount(i) > 0 Then
            leafCount(parentIndex) = leafCount(parentIndex) + leafCount(i)
            progressSum(parentIndex) = progressSum(parentIndex) + progressSum(i)
            If hasHours(i) Then
                totalHours(parentIndex) = totalHours(parentIndex) + totalHours(i)
                weightedProgress(parentIndex) = weightedProgress(parentIndex) + weightedProgress(i)
                hasHours(parentIndex) = True
            End If
            leafHoursExists(parentIndex) = leafHoursExists(parentIndex) Or leafHoursExists(i)
            allComplete(parentIndex) = allComplete(parentIndex) And allComplete(i)
            anyInProgress(parentIndex) = anyInProgress(parentIndex) Or anyInProgress(i)
            anyOverdueIncomplete(parentIndex) = anyOverdueIncomplete(parentIndex) Or anyOverdueIncomplete(i)
            allFutureOnly(parentIndex) = allFutureOnly(parentIndex) And allFutureOnly(i)
            UpdateMinDate planStart(parentIndex), planStart(i)
            UpdateMaxDate planEnd(parentIndex), planEnd(i)
            UpdateMinDate actualStart(parentIndex), actualStart(i)
            UpdateMaxDate actualEnd(parentIndex), actualEnd(i)
        End If

ContinueRow:
    Next i
End Sub

Private Function IsThisWeek(ByVal targetDate As Date, ByVal referenceDate As Date) As Boolean
    Dim weekStart As Date
    Dim weekEnd As Date

    weekStart = referenceDate - (Weekday(referenceDate, vbMonday) - 1)
    weekEnd = weekStart + 6
    IsThisWeek = (targetDate >= weekStart And targetDate <= weekEnd)
End Function

Private Function PlanRangeIntersectsThisWeek(ByVal startPlan As Variant, ByVal endPlan As Variant, ByVal referenceDate As Date) As Boolean
    Dim weekStart As Date
    Dim weekEnd As Date
    Dim planStartDate As Date
    Dim planEndDate As Date
    Dim hasStart As Boolean
    Dim hasEnd As Boolean

    hasStart = InazumaGantt_v3.TryGetDateValue(startPlan, planStartDate)
    hasEnd = InazumaGantt_v3.TryGetDateValue(endPlan, planEndDate)
    If Not hasStart And Not hasEnd Then Exit Function

    If hasStart And hasEnd Then
        weekStart = referenceDate - (Weekday(referenceDate, vbMonday) - 1)
        weekEnd = weekStart + 6
        PlanRangeIntersectsThisWeek = (planStartDate <= weekEnd And planEndDate >= weekStart)
    ElseIf hasStart Then
        PlanRangeIntersectsThisWeek = IsThisWeek(planStartDate, referenceDate)
    Else
        PlanRangeIntersectsThisWeek = IsThisWeek(planEndDate, referenceDate)
    End If
End Function

Private Function GetAlertSeverity(ByVal markerText As String) As Long
    markerText = Trim$(markerText)
    Select Case markerText
        Case ALERT_MARK_DELAY
            GetAlertSeverity = 2
        Case ALERT_MARK_TODAY
            GetAlertSeverity = 1
        Case Else
            GetAlertSeverity = 0
    End Select
End Function

Private Function StrongerAlertMarker(ByVal currentMarker As String, ByVal candidateMarker As String) As String
    If GetAlertSeverity(candidateMarker) > GetAlertSeverity(currentMarker) Then
        StrongerAlertMarker = candidateMarker
    Else
        StrongerAlertMarker = currentMarker
    End If
End Function

Private Function RemoveAlertMarkerPrefix(ByVal textValue As String) As String
    textValue = Trim$(textValue)

    If textValue = ALERT_MARK_DELAY Then
        RemoveAlertMarkerPrefix = ""
    ElseIf Left$(textValue, Len(ALERT_MARK_DELAY & " ")) = ALERT_MARK_DELAY & " " Then
        RemoveAlertMarkerPrefix = Trim$(Mid$(textValue, Len(ALERT_MARK_DELAY & " ") + 1))
    ElseIf textValue = ALERT_MARK_TODAY Then
        RemoveAlertMarkerPrefix = ""
    ElseIf Left$(textValue, Len(ALERT_MARK_TODAY & " ")) = ALERT_MARK_TODAY & " " Then
        RemoveAlertMarkerPrefix = Trim$(Mid$(textValue, Len(ALERT_MARK_TODAY & " ") + 1))
    Else
        RemoveAlertMarkerPrefix = textValue
    End If
End Function

Private Function BuildMarkedPrimaryTaskText(ByVal currentValue As String, ByVal markerText As String) As String
    Dim baseText As String

    baseText = RemoveAlertMarkerPrefix(currentValue)
    markerText = Trim$(markerText)

    If markerText <> "" And baseText <> "" Then
        BuildMarkedPrimaryTaskText = markerText & " " & baseText
    Else
        BuildMarkedPrimaryTaskText = baseText
    End If
End Function

Private Function HasPrimaryTaskTextInMarkerCell(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    Dim textValue As String

    textValue = Trim$(CStr(ws.Cells(targetRow, "C").Value))
    HasPrimaryTaskTextInMarkerCell = (textValue <> "" And Not IsMarkerOnlyText(textValue) And _
                                      Not InazumaGantt_v3.IsAuxiliaryPlaceholderText(textValue))
End Function

Private Function DetermineSelfAlertMarker(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date) As String
    Dim rowLevel As Long
    Dim statusText As String
    Dim progressValue As Double
    Dim startPlan As Variant
    Dim endPlan As Variant
    Dim endPlanDate As Date

    rowLevel = GetHierarchyLevel(ws, targetRow)
    If rowLevel <= 0 Then Exit Function
    If Not InazumaGantt_v3.HasTaskContentInRow(ws, targetRow) Then Exit Function

    statusText = Trim$(CStr(ws.Cells(targetRow, InazumaGantt_v3.COL_STATUS).Value))
    progressValue = InazumaGantt_v3.NormalizeProgressValue(ws.Cells(targetRow, InazumaGantt_v3.COL_PROGRESS).Value, 0)
    If statusText = "完了" Or progressValue >= 1 Then Exit Function

    startPlan = ws.Cells(targetRow, InazumaGantt_v3.COL_START_PLAN).Value
    endPlan = ws.Cells(targetRow, InazumaGantt_v3.COL_END_PLAN).Value

    If InazumaGantt_v3.TryGetDateValue(endPlan, endPlanDate) Then
        If endPlanDate < referenceDate Then
            DetermineSelfAlertMarker = ALERT_MARK_DELAY
            Exit Function
        End If
    End If

    If PlanRangeIntersectsThisWeek(startPlan, endPlan, referenceDate) Then
        DetermineSelfAlertMarker = ALERT_MARK_TODAY
    End If
End Function

Private Function DetermineDescendantAlertMarker(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date) As String
    Dim targetLevel As Long
    Dim lastRow As Long
    Dim r As Long
    Dim rowLevel As Long
    Dim markerText As String

    targetLevel = GetHierarchyLevel(ws, targetRow)
    If targetLevel <= 0 Then Exit Function

    lastRow = InazumaGantt_v3.GetLastDataRow(ws)
    For r = targetRow + 1 To lastRow
        rowLevel = GetHierarchyLevel(ws, r)
        If rowLevel > 0 Then
            If rowLevel <= targetLevel Then Exit For
            markerText = DetermineSelfAlertMarker(ws, r, referenceDate)
            DetermineDescendantAlertMarker = StrongerAlertMarker(DetermineDescendantAlertMarker, markerText)
            If DetermineDescendantAlertMarker = ALERT_MARK_DELAY Then Exit For
        End If
    Next r
End Function

Private Function DetermineAlertMarker(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date) As String
    DetermineAlertMarker = StrongerAlertMarker(DetermineSelfAlertMarker(ws, targetRow, referenceDate), _
                                               DetermineDescendantAlertMarker(ws, targetRow, referenceDate))
End Function

Private Sub ApplyResolvedAlertMarkerState(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal markerText As String)
    Dim rowLevel As Long
    Dim markerCell As Range
    Dim currentValue As String
    Dim nextValue As String

    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    rowLevel = GetHierarchyLevel(ws, targetRow)
    Set markerCell = ws.Cells(targetRow, "C")
    currentValue = Trim$(CStr(markerCell.Value))

    If rowLevel <= 0 Then
        If IsMarkerOnlyText(currentValue) Then
            markerCell.ClearContents
        End If
        InazumaGantt_v3.RefreshTaskLabelPresentation ws, targetRow
        Exit Sub
    End If

    If HasPrimaryTaskTextInMarkerCell(ws, targetRow) Then
        nextValue = BuildMarkedPrimaryTaskText(currentValue, markerText)
    ElseIf InazumaGantt_v3.HasPrimaryTaskContentInRow(ws, targetRow) Then
        nextValue = markerText
    ElseIf InazumaGantt_v3.HasTaskContentInRow(ws, targetRow) Then
        nextValue = InazumaGantt_v3.BuildAuxiliaryDisplayText(markerText)
    End If

    If currentValue <> nextValue Then
        markerCell.Value = nextValue
    End If

    InazumaGantt_v3.RefreshTaskLabelPresentation ws, targetRow
End Sub

Private Sub ApplyAlertMarkerState(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal referenceDate As Date)
    ApplyResolvedAlertMarkerState ws, targetRow, DetermineAlertMarker(ws, targetRow, referenceDate)
End Sub

Public Sub RefreshTaskAlertMarkers(ByVal ws As Worksheet)
    Const MAX_TRACKED_LEVEL As Long = 32

    Dim workingWs As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim r As Long
    Dim rowIndex As Long
    Dim levelIndex As Long
    Dim parentIndex As Long
    Dim referenceDate As Date
    Dim levels() As Long
    Dim parentRows() As Long
    Dim resolvedMarkers() As String
    Dim levelStack(1 To MAX_TRACKED_LEVEL) As Long

    Set workingWs = GetRollupWorksheet(ws, "タスク強調表示")
    If workingWs Is Nothing Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    rowCount = lastRow - InazumaGantt_v3.ROW_DATA_START + 1
    ReDim levels(1 To rowCount)
    ReDim parentRows(1 To rowCount)
    ReDim resolvedMarkers(1 To rowCount)
    referenceDate = Date

    For r = InazumaGantt_v3.ROW_DATA_START To lastRow
        rowIndex = r - InazumaGantt_v3.ROW_DATA_START + 1
        If IsNumeric(workingWs.Cells(r, InazumaGantt_v3.COL_HIERARCHY).Value) Then
            levels(rowIndex) = CLng(workingWs.Cells(r, InazumaGantt_v3.COL_HIERARCHY).Value)
        End If

        If levels(rowIndex) > 0 And levels(rowIndex) <= MAX_TRACKED_LEVEL Then
            For levelIndex = levels(rowIndex) - 1 To 1 Step -1
                If levelStack(levelIndex) > 0 Then
                    parentRows(rowIndex) = levelStack(levelIndex)
                    Exit For
                End If
            Next levelIndex

            levelStack(levels(rowIndex)) = rowIndex
            For levelIndex = levels(rowIndex) + 1 To MAX_TRACKED_LEVEL
                levelStack(levelIndex) = 0
            Next levelIndex
        End If

        resolvedMarkers(rowIndex) = DetermineSelfAlertMarker(workingWs, r, referenceDate)
    Next r

    For rowIndex = rowCount To 1 Step -1
        parentIndex = parentRows(rowIndex)
        If parentIndex > 0 Then
            resolvedMarkers(parentIndex) = StrongerAlertMarker(resolvedMarkers(parentIndex), resolvedMarkers(rowIndex))
        End If
    Next rowIndex

    For r = InazumaGantt_v3.ROW_DATA_START To lastRow
        rowIndex = r - InazumaGantt_v3.ROW_DATA_START + 1
        ApplyResolvedAlertMarkerState workingWs, r, resolvedMarkers(rowIndex)
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

Public Sub RefreshTaskAlertMarkersForRowsAndAncestors(ByVal ws As Worksheet, ByVal affectedRows As Object)
    Dim workingWs As Worksheet
    Dim rowKeys As Object
    Dim parentRows As Object
    Dim rowKey As Variant
    Dim r As Long
    Dim lastRow As Long
    Dim maxRow As Long
    Dim referenceDate As Date

    Set workingWs = GetRollupWorksheet(ws, "タスク強調表示")
    If workingWs Is Nothing Then Exit Sub
    If affectedRows Is Nothing Then Exit Sub
    If affectedRows.Count = 0 Then Exit Sub

    lastRow = InazumaGantt_v3.GetLastDataRow(workingWs)
    maxRow = GetMaxAffectedRow(affectedRows, lastRow)
    referenceDate = Date
    Set rowKeys = CreateObject("Scripting.Dictionary")
    Set parentRows = BuildParentRowMap(workingWs, maxRow)

    For Each rowKey In affectedRows.Keys
        AddRowAndAncestorKeys CLng(rowKey), parentRows, rowKeys
    Next rowKey

    For r = InazumaGantt_v3.ROW_DATA_START To lastRow
        If rowKeys.Exists(CStr(r)) Then
            ApplyAlertMarkerState workingWs, r, referenceDate
        End If
    Next r
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
