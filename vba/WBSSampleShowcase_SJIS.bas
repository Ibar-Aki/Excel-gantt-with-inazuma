Attribute VB_Name = "WBSSampleShowcase"
Option Explicit

Private Const SHOWCASE_ROW_COUNT As Long = 156

Public Sub CreateShowcaseSampleWBS(Optional ByVal baseDate As Date = 0, _
                                   Optional ByVal skipExistingCheck As Boolean = False, _
                                   Optional ByVal refreshAfterCreate As Boolean = True)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("Showcase Sample WBS")
    If ws Is Nothing Then Exit Sub

    If baseDate = 0 Then
        If IsDate(ws.Range(InazumaGantt_v3.CELL_PROJECT_START).Value) Then
            baseDate = CDate(ws.Range(InazumaGantt_v3.CELL_PROJECT_START).Value)
        Else
            baseDate = Date
        End If
    End If

    If Not skipExistingCheck And MainSheetHasTaskData(ws) Then
        MsgBox "Existing tasks were found. Showcase sample creation was skipped.", vbInformation, "Showcase Sample"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ClearTaskArea ws
    BuildShowcaseData ws, baseDate

    ws.Range("A4").Value = "Memo: Aether Pulse showcase sample (" & CStr(SHOWCASE_ROW_COUNT) & " rows)"

    InazumaGantt_v3.AutoDetectTaskLevel
    InazumaGantt_v3.RenumberRows

    If refreshAfterCreate Then
        HierarchyColor.SetupHierarchyColors
        InazumaGantt_v3.RefreshInazumaGantt
    End If

    Application.ScreenUpdating = True

    If Application.DisplayAlerts Then
        MsgBox "Showcase sample WBS was created." & vbCrLf & _
               "Rows: " & CStr(SHOWCASE_ROW_COUNT), vbInformation, "Showcase Sample"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Showcase sample creation failed: " & Err.Description, vbCritical, "Showcase Sample"
End Sub

Private Function RequireMainWorksheet(ByVal operationName As String) As Worksheet
    On Error Resume Next
    Set RequireMainWorksheet = ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME)
    On Error GoTo 0

    If RequireMainWorksheet Is Nothing Then
        MsgBox "Main sheet '" & InazumaGantt_v3.MAIN_SHEET_NAME & "' was not found.", vbExclamation, operationName
    End If
End Function

Private Function MainSheetHasTaskData(ByVal ws As Worksheet) As Boolean
    Dim lastRow As Long

    lastRow = InazumaGantt_v3.GetLastDataRow(ws)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then Exit Function

    MainSheetHasTaskData = (Application.WorksheetFunction.CountA(ws.Range("C" & InazumaGantt_v3.ROW_DATA_START & ":F" & lastRow)) > 0)
End Function

Private Sub ClearTaskArea(ByVal ws As Worksheet)
    Dim endRow As Long
    endRow = InazumaGantt_v3.ROW_DATA_START + InazumaGantt_v3.DATA_ROWS_DEFAULT - 1

    ws.Range("A" & InazumaGantt_v3.ROW_DATA_START & ":O" & endRow).ClearContents
End Sub

Private Sub BuildShowcaseData(ByVal ws As Worksheet, ByVal baseDate As Date)
    Dim phaseNames As Variant
    Dim owners As Variant
    Dim rowIndex As Long
    Dim phaseIndex As Long

    phaseNames = Array( _
        "North Star Strategy", _
        "Brand Storyworld", _
        "Experience Blueprint", _
        "Platform Foundation", _
        "Commerce Engine", _
        "Data Intelligence", _
        "Creator Studio", _
        "Launch Command", _
        "Partner Orbit", _
        "Trust and Security", _
        "Scale Automation", _
        "Growth Catalyst")

    owners = Array("Aki", "Mika", "Ren", "Sora", "Kai", "Yui", "Jin", "Riku")

    rowIndex = InazumaGantt_v3.ROW_DATA_START

    For phaseIndex = LBound(phaseNames) To UBound(phaseNames)
        AddShowcasePhase ws, rowIndex, phaseIndex + 1, CStr(phaseNames(phaseIndex)), owners, baseDate
    Next phaseIndex
End Sub

Private Sub AddShowcasePhase(ByVal ws As Worksheet, ByRef rowIndex As Long, ByVal phaseIndex As Long, _
                             ByVal phaseName As String, ByVal owners As Variant, ByVal baseDate As Date)
    Dim phaseMode As String
    Dim phaseOwner As String
    Dim phaseStatus As String
    Dim phaseProgress As Double
    Dim phaseStart As Date
    Dim phaseEnd As Date
    Dim phaseActualStart As Variant
    Dim phaseActualEnd As Variant
    Dim streamIndex As Long

    phaseMode = GetPhaseMode(phaseIndex)
    phaseOwner = CStr(owners((phaseIndex - 1) Mod (UBound(owners) + 1)))
    phaseStatus = GetParentStatus(phaseMode)
    phaseProgress = GetParentProgress(phaseMode)
    phaseStart = AddBusinessDays(baseDate, GetPhaseStartOffset(phaseIndex))
    phaseEnd = AddBusinessDays(phaseStart, GetPhaseDuration(phaseMode))

    If phaseProgress > 0 Then phaseActualStart = phaseStart
    If phaseStatus = "完了" Then phaseActualEnd = AddBusinessDays(phaseEnd, IIf(phaseIndex Mod 2 = 0, -1, 0))

    SetTaskRow ws, rowIndex, "C", phaseName, "Signature phase for the Aether Pulse showcase program.", _
               phaseStatus, phaseProgress, phaseOwner, "", phaseStart, phaseEnd, phaseActualStart, phaseActualEnd
    rowIndex = rowIndex + 1

    For streamIndex = 1 To 3
        AddShowcaseStream ws, rowIndex, phaseIndex, streamIndex, phaseName, phaseMode, owners, phaseStart, phaseEnd
    Next streamIndex
End Sub

Private Sub AddShowcaseStream(ByVal ws As Worksheet, ByRef rowIndex As Long, ByVal phaseIndex As Long, _
                              ByVal streamIndex As Long, ByVal phaseName As String, ByVal phaseMode As String, _
                              ByVal owners As Variant, ByVal phaseStart As Date, ByVal phaseEnd As Date)
    Dim streamOwner As String
    Dim streamName As String
    Dim streamStart As Date
    Dim streamEnd As Date
    Dim streamStatus As String
    Dim streamProgress As Double
    Dim streamActualStart As Variant
    Dim streamActualEnd As Variant
    Dim leafIndex As Long
    Dim taskStatus As String
    Dim taskProgress As Double
    Dim actualStart As Variant
    Dim actualEnd As Variant
    Dim lv3ParentName As String
    Dim lv4TaskName As String
    Dim lv3LeafName As String
    Dim taskHours As Double

    streamOwner = CStr(owners((phaseIndex + streamIndex - 1) Mod (UBound(owners) + 1)))
    streamName = GetStreamName(streamIndex, phaseName)
    streamStatus = GetParentStatus(phaseMode)
    streamProgress = GetParentProgress(phaseMode)
    streamStart = AddBusinessDays(phaseStart, (streamIndex - 1) * 2)
    streamEnd = AddBusinessDays(phaseEnd, -((3 - streamIndex) * 2))

    If streamProgress > 0 Then streamActualStart = streamStart
    If streamStatus = "完了" Then streamActualEnd = AddBusinessDays(streamEnd, IIf(streamIndex = 2, -1, 0))

    SetTaskRow ws, rowIndex, "D", streamName, GetStreamDetail(streamIndex, phaseName), _
               streamStatus, streamProgress, streamOwner, "", streamStart, streamEnd, streamActualStart, streamActualEnd
    rowIndex = rowIndex + 1

    lv3ParentName = GetTaskName(streamIndex, 1, phaseName)
    SetTaskRow ws, rowIndex, "E", lv3ParentName, "Parent cluster that groups the hero deliverables.", _
               streamStatus, streamProgress, streamOwner, "", _
               AddBusinessDays(streamStart, 0), AddBusinessDays(streamStart, 3), streamActualStart, Empty
    rowIndex = rowIndex + 1

    leafIndex = ((streamIndex - 1) * 2) + 1
    GetLeafExecutionState phaseMode, leafIndex, taskStatus, taskProgress
    ResolveActualDates phaseMode, leafIndex, AddBusinessDays(streamStart, 1), AddBusinessDays(streamStart, 4), taskStatus, actualStart, actualEnd
    lv4TaskName = GetTaskName(streamIndex, 2, phaseName)
    taskHours = GetLeafHours(phaseIndex, leafIndex)
    SetTaskRow ws, rowIndex, "F", lv4TaskName, GetLeafDetail(streamIndex, 1, phaseName), _
               taskStatus, taskProgress, streamOwner, InazumaGantt_v3.FormatDevelopmentHours(taskHours), _
               AddBusinessDays(streamStart, 1), AddBusinessDays(streamStart, 4), actualStart, actualEnd
    rowIndex = rowIndex + 1

    leafIndex = leafIndex + 1
    GetLeafExecutionState phaseMode, leafIndex, taskStatus, taskProgress
    ResolveActualDates phaseMode, leafIndex, AddBusinessDays(streamStart, 3), AddBusinessDays(streamEnd, 0), taskStatus, actualStart, actualEnd
    lv3LeafName = GetTaskName(streamIndex, 3, phaseName)
    taskHours = GetLeafHours(phaseIndex, leafIndex)
    SetTaskRow ws, rowIndex, "E", lv3LeafName, GetLeafDetail(streamIndex, 2, phaseName), _
               taskStatus, taskProgress, streamOwner, InazumaGantt_v3.FormatDevelopmentHours(taskHours), _
               AddBusinessDays(streamStart, 3), AddBusinessDays(streamEnd, 0), actualStart, actualEnd
    rowIndex = rowIndex + 1
End Sub

Private Sub SetTaskRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal taskColumn As String, _
                       ByVal taskName As String, ByVal detailText As String, ByVal statusText As String, _
                       ByVal progressValue As Double, ByVal ownerName As String, ByVal hoursText As String, _
                       ByVal planStart As Variant, ByVal planEnd As Variant, ByVal actualStart As Variant, _
                       ByVal actualEnd As Variant)
    ws.Cells(rowIndex, taskColumn).Value = taskName
    ws.Cells(rowIndex, InazumaGantt_v3.COL_TASK_DETAIL).Value = detailText
    ws.Cells(rowIndex, InazumaGantt_v3.COL_STATUS).Value = statusText
    ws.Cells(rowIndex, InazumaGantt_v3.COL_PROGRESS).Value = progressValue
    ws.Cells(rowIndex, InazumaGantt_v3.COL_ASSIGNEE).Value = ownerName

    If hoursText <> "" Then ws.Cells(rowIndex, InazumaGantt_v3.COL_DEV_LT).Value = hoursText
    If IsDate(planStart) Then ws.Cells(rowIndex, InazumaGantt_v3.COL_START_PLAN).Value = CDate(planStart)
    If IsDate(planEnd) Then ws.Cells(rowIndex, InazumaGantt_v3.COL_END_PLAN).Value = CDate(planEnd)
    If IsDate(actualStart) Then ws.Cells(rowIndex, InazumaGantt_v3.COL_START_ACTUAL).Value = CDate(actualStart)
    If IsDate(actualEnd) Then ws.Cells(rowIndex, InazumaGantt_v3.COL_END_ACTUAL).Value = CDate(actualEnd)
End Sub

Private Function GetPhaseMode(ByVal phaseIndex As Long) As String
    Select Case phaseIndex
        Case 1 To 4
            GetPhaseMode = "complete"
        Case 5, 6, 9
            GetPhaseMode = "active"
        Case 7
            GetPhaseMode = "delayed"
        Case 8
            GetPhaseMode = "dueSoon"
        Case 10
            GetPhaseMode = "hold"
        Case Else
            GetPhaseMode = "planned"
    End Select
End Function

Private Function GetPhaseStartOffset(ByVal phaseIndex As Long) As Long
    Select Case phaseIndex
        Case 1: GetPhaseStartOffset = 0
        Case 2: GetPhaseStartOffset = 2
        Case 3: GetPhaseStartOffset = 4
        Case 4: GetPhaseStartOffset = 6
        Case 5: GetPhaseStartOffset = 10
        Case 6: GetPhaseStartOffset = 12
        Case 7: GetPhaseStartOffset = 8
        Case 8: GetPhaseStartOffset = 13
        Case 9: GetPhaseStartOffset = 16
        Case 10: GetPhaseStartOffset = 18
        Case 11: GetPhaseStartOffset = 24
        Case Else: GetPhaseStartOffset = 28
    End Select
End Function

Private Function GetPhaseDuration(ByVal phaseMode As String) As Long
    Select Case phaseMode
        Case "complete": GetPhaseDuration = 8
        Case "active": GetPhaseDuration = 12
        Case "delayed": GetPhaseDuration = 8
        Case "dueSoon": GetPhaseDuration = 8
        Case "hold": GetPhaseDuration = 10
        Case Else: GetPhaseDuration = 12
    End Select
End Function

Private Function GetParentStatus(ByVal phaseMode As String) As String
    Select Case phaseMode
        Case "complete": GetParentStatus = "完了"
        Case "hold": GetParentStatus = "保留"
        Case "planned": GetParentStatus = "未着手"
        Case Else: GetParentStatus = "進行中"
    End Select
End Function

Private Function GetParentProgress(ByVal phaseMode As String) As Double
    Select Case phaseMode
        Case "complete": GetParentProgress = 1
        Case "active": GetParentProgress = 0.62
        Case "delayed": GetParentProgress = 0.41
        Case "dueSoon": GetParentProgress = 0.76
        Case "hold": GetParentProgress = 0.18
        Case Else: GetParentProgress = 0
    End Select
End Function

Private Function GetStreamName(ByVal streamIndex As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1: GetStreamName = phaseName & " Signal Track"
        Case 2: GetStreamName = phaseName & " Build Track"
        Case Else: GetStreamName = phaseName & " Launch Track"
    End Select
End Function

Private Function GetStreamDetail(ByVal streamIndex As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1: GetStreamDetail = "Shapes the intent and success frame for " & phaseName & "."
        Case 2: GetStreamDetail = "Turns the concept into a visible and measurable experience."
        Case Else: GetStreamDetail = "Prepares enablement, operations, and launch confidence."
    End Select
End Function

Private Function GetTaskName(ByVal streamIndex As Long, ByVal slotIndex As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1
            Select Case slotIndex
                Case 1: GetTaskName = phaseName & " Signal Brief"
                Case 2: GetTaskName = phaseName & " Executive Cut"
                Case Else: GetTaskName = phaseName & " Success Scoreboard"
            End Select
        Case 2
            Select Case slotIndex
                Case 1: GetTaskName = phaseName & " Hero Flow"
                Case 2: GetTaskName = phaseName & " Prototype Sprint"
                Case Else: GetTaskName = phaseName & " Integration Map"
            End Select
        Case Else
            Select Case slotIndex
                Case 1: GetTaskName = phaseName & " Launch Rehearsal"
                Case 2: GetTaskName = phaseName & " War Room Drill"
                Case Else: GetTaskName = phaseName & " Enablement Pack"
            End Select
    End Select
End Function

Private Function GetLeafDetail(ByVal streamIndex As Long, ByVal leafSlot As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1
            If leafSlot = 1 Then
                GetLeafDetail = "Tight edit for the story spine of " & phaseName & "."
            Else
                GetLeafDetail = "Measures focus, outcomes, and stakeholder confidence."
            End If
        Case 2
            If leafSlot = 1 Then
                GetLeafDetail = "Fast prototype cycle to make the hero experience tangible."
            Else
                GetLeafDetail = "System and channel handshake design for launch quality."
            End If
        Case Else
            If leafSlot = 1 Then
                GetLeafDetail = "Operational drill that stress-tests the go-live moment."
            Else
                GetLeafDetail = "Field-ready assets that make the launch feel polished."
            End If
    End Select
End Function

Private Sub GetLeafExecutionState(ByVal phaseMode As String, ByVal leafIndex As Long, _
                                  ByRef statusText As String, ByRef progressValue As Double)
    Select Case phaseMode
        Case "complete"
            statusText = "完了"
            progressValue = 1
        Case "active"
            Select Case leafIndex
                Case 1, 2
                    statusText = "完了": progressValue = 1
                Case 3
                    statusText = "進行中": progressValue = 0.85
                Case 4
                    statusText = "進行中": progressValue = 0.65
                Case 5
                    statusText = "進行中": progressValue = 0.3
                Case Else
                    statusText = "未着手": progressValue = 0
            End Select
        Case "delayed"
            Select Case leafIndex
                Case 1
                    statusText = "完了": progressValue = 1
                Case 2
                    statusText = "進行中": progressValue = 0.9
                Case 3
                    statusText = "進行中": progressValue = 0.55
                Case 4
                    statusText = "進行中": progressValue = 0.2
                Case 5
                    statusText = "未着手": progressValue = 0
                Case Else
                    statusText = "保留": progressValue = 0
            End Select
        Case "dueSoon"
            Select Case leafIndex
                Case 1
                    statusText = "完了": progressValue = 1
                Case 2
                    statusText = "進行中": progressValue = 0.95
                Case 3
                    statusText = "進行中": progressValue = 0.75
                Case 4
                    statusText = "進行中": progressValue = 0.6
                Case 5
                    statusText = "進行中": progressValue = 0.2
                Case Else
                    statusText = "未着手": progressValue = 0
            End Select
        Case "hold"
            Select Case leafIndex
                Case 1
                    statusText = "完了": progressValue = 1
                Case 2
                    statusText = "保留": progressValue = 0.35
                Case 3
                    statusText = "保留": progressValue = 0.1
                Case 4
                    statusText = "進行中": progressValue = 0.25
                Case 5
                    statusText = "未着手": progressValue = 0
                Case Else
                    statusText = "保留": progressValue = 0
            End Select
        Case Else
            statusText = "未着手"
            progressValue = 0
    End Select
End Sub

Private Sub ResolveActualDates(ByVal phaseMode As String, ByVal leafIndex As Long, ByVal planStart As Date, _
                               ByVal planEnd As Date, ByVal statusText As String, _
                               ByRef actualStart As Variant, ByRef actualEnd As Variant)
    If statusText = "完了" Then
        actualStart = planStart
        Select Case phaseMode
            Case "complete"
                actualEnd = AddBusinessDays(planEnd, IIf(leafIndex Mod 2 = 0, -1, 0))
            Case "delayed"
                actualEnd = AddBusinessDays(planEnd, 1)
            Case Else
                actualEnd = planEnd
        End Select
    ElseIf statusText = "進行中" Or statusText = "保留" Then
        actualStart = planStart
    End If
End Sub

Private Function GetLeafHours(ByVal phaseIndex As Long, ByVal leafIndex As Long) As Double
    Dim baseHours As Variant

    baseHours = Array(8, 10, 12, 14, 9, 11)
    GetLeafHours = CDbl(baseHours((leafIndex - 1) Mod 6)) + ((phaseIndex - 1) Mod 3)
End Function

Private Function AddBusinessDays(ByVal originDate As Date, ByVal offsetDays As Long) As Date
    Dim resultDate As Date
    Dim stepValue As Long
    Dim remaining As Long

    resultDate = originDate
    remaining = Abs(offsetDays)

    If offsetDays < 0 Then
        stepValue = -1
    Else
        stepValue = 1
    End If

    Do While remaining > 0
        resultDate = resultDate + stepValue
        If Weekday(resultDate, vbMonday) <= 5 Then remaining = remaining - 1
    Loop

    If offsetDays = 0 And Weekday(resultDate, vbMonday) > 5 Then
        Do While Weekday(resultDate, vbMonday) > 5
            resultDate = resultDate + 1
        Loop
    End If

    AddBusinessDays = resultDate
End Function
