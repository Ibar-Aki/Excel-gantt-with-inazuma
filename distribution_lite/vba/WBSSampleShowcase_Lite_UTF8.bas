Attribute VB_Name = "WBSSampleShowcase"
Option Explicit

Private Const SHOWCASE_ROW_COUNT As Long = 156
Private Const SHOWCASE_APP_NAME As String = "会話力向上アプリ「ハナセル」"
Private Const SHOWCASE_REFERENCE_DATE_OFFSET As Long = 19

Public Sub CreateShowcaseSampleWBS(Optional ByVal baseDate As Date = 0, _
                                   Optional ByVal skipExistingCheck As Boolean = False, _
                                   Optional ByVal refreshAfterCreate As Boolean = True, _
                                   Optional ByVal generateRoadmapView As Boolean = True)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim previousEnableEvents As Boolean
    Set ws = RequireMainWorksheet("サンプルWBS作成")
    If ws Is Nothing Then Exit Sub

    If baseDate = 0 Then
        If IsDate(ws.Range(InazumaGantt_v3.CELL_PROJECT_START).Value) Then
            baseDate = CDate(ws.Range(InazumaGantt_v3.CELL_PROJECT_START).Value)
        Else
            baseDate = Date
        End If
    End If

    If Not skipExistingCheck And MainSheetHasTaskData(ws) Then
        MsgBox "既存タスクがあるため、サンプルWBSの生成をスキップしました。", vbInformation, "サンプルWBS"
        Exit Sub
    End If

    previousEnableEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ClearTaskArea ws
    BuildShowcaseData ws, baseDate

    ws.Range("A4").Value = "メモ：" & SHOWCASE_APP_NAME & " 開発のデモWBS（" & CStr(SHOWCASE_ROW_COUNT) & "行、WBSサマリ付き）"

    InazumaGantt_v3.AutoDetectTaskLevel
    InazumaGantt_v3.RenumberRows

    If refreshAfterCreate Then
        HierarchyColor.SetupHierarchyColors
        InazumaGantt_v3.RefreshInazumaGantt
    End If

    If generateRoadmapView Then
        FinalizeShowcasePresentation True, GetShowcaseReferenceDate(baseDate)
    End If

    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = True

    If Application.DisplayAlerts Then
        MsgBox "サンプルWBSを生成しました。" & vbCrLf & _
               "行数: " & CStr(SHOWCASE_ROW_COUNT), vbInformation, "サンプルWBS"
    End If
    Exit Sub

ErrorHandler:
    Application.EnableEvents = previousEnableEvents
    Application.ScreenUpdating = True
    If Application.DisplayAlerts Then
        MsgBox "サンプルWBSの生成に失敗しました: " & Err.Description, vbCritical, "サンプルWBS"
    Else
        Err.Raise Err.Number, "CreateShowcaseSampleWBS", Err.Description
    End If
End Sub

Public Sub FinalizeShowcasePresentation(Optional ByVal generateRoadmapView As Boolean = True, _
                                        Optional ByVal roadmapReferenceDate As Variant)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("サンプルWBS仕上げ")
    If ws Is Nothing Then Exit Sub

    If generateRoadmapView Then
        WBSRoadmapReport.CreateRoadmapOverviewSheet roadmapReferenceDate
    End If
    Exit Sub

ErrorHandler:
    If Application.DisplayAlerts Then
        MsgBox "サンプルWBSの仕上げに失敗しました: " & Err.Description, vbCritical, "サンプルWBS"
    Else
        Err.Raise Err.Number, "FinalizeShowcasePresentation", Err.Description
    End If
End Sub

Public Function GetShowcaseReferenceDate(ByVal baseDate As Date) As Date
    GetShowcaseReferenceDate = DateAdd("d", SHOWCASE_REFERENCE_DATE_OFFSET, baseDate)
End Function

Private Function RequireMainWorksheet(ByVal operationName As String) As Worksheet
    On Error Resume Next
    Set RequireMainWorksheet = ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME)
    On Error GoTo 0

    If RequireMainWorksheet Is Nothing Then
        MsgBox "メインシート '" & InazumaGantt_v3.MAIN_SHEET_NAME & "' が見つかりません。", vbExclamation, operationName
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

    ws.Range("A" & InazumaGantt_v3.ROW_DATA_START & ":M" & endRow).ClearContents
End Sub

Private Sub BuildShowcaseData(ByVal ws As Worksheet, ByVal baseDate As Date)
    Dim phaseNames As Variant
    Dim owners As Variant
    Dim rowIndex As Long
    Dim phaseIndex As Long

    phaseNames = Array( _
        "事業構想と市場分析", _
        "学習体験コンセプト設計", _
        "会話診断ロジック設計", _
        "レッスンコンテンツ制作", _
        "UX/UIデザイン", _
        "AI会話エンジン開発", _
        "音声認識・発話評価", _
        "モバイルアプリ実装", _
        "バックエンド・分析基盤", _
        "ベータ運用と改善", _
        "セキュリティ・法務対応", _
        "正式リリースとグロース")

    owners = Array("秋山", "美咲", "蓮", "蒼", "海斗", "結衣", "仁", "陸")

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
    Dim streamIndex As Long

    phaseMode = GetPhaseMode(phaseIndex)
    phaseOwner = CStr(owners((phaseIndex - 1) Mod (UBound(owners) + 1)))
    phaseStatus = GetParentStatus(phaseMode)
    phaseProgress = GetParentProgress(phaseMode)
    phaseStart = AddBusinessDays(baseDate, GetPhaseStartOffset(phaseIndex))
    phaseEnd = AddBusinessDays(phaseStart, GetPhaseDuration(phaseMode))

    SetTaskRow ws, rowIndex, "C", phaseName, SHOWCASE_APP_NAME & " の中核フェーズ。", _
               phaseStatus, phaseProgress, phaseOwner, "", phaseStart, phaseEnd
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
    Dim leafIndex As Long
    Dim taskStatus As String
    Dim taskProgress As Double
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

    SetTaskRow ws, rowIndex, "D", streamName, GetStreamDetail(streamIndex, phaseName), _
               streamStatus, streamProgress, streamOwner, "", streamStart, streamEnd
    rowIndex = rowIndex + 1

    lv3ParentName = GetTaskName(streamIndex, 1, phaseName)
    SetTaskRow ws, rowIndex, "E", lv3ParentName, "主要成果物を束ねる親タスク。", _
               streamStatus, streamProgress, streamOwner, "", _
               AddBusinessDays(streamStart, 0), AddBusinessDays(streamStart, 3)
    rowIndex = rowIndex + 1

    leafIndex = ((streamIndex - 1) * 2) + 1
    GetLeafExecutionState phaseMode, leafIndex, taskStatus, taskProgress
    lv4TaskName = GetTaskName(streamIndex, 2, phaseName)
    taskHours = GetLeafHours(phaseIndex, leafIndex)
    SetTaskRow ws, rowIndex, "F", lv4TaskName, GetLeafDetail(streamIndex, 1, phaseName), _
               taskStatus, taskProgress, streamOwner, InazumaGantt_v3.FormatDevelopmentHours(taskHours), _
               AddBusinessDays(streamStart, 1), AddBusinessDays(streamStart, 4)
    rowIndex = rowIndex + 1

    leafIndex = leafIndex + 1
    GetLeafExecutionState phaseMode, leafIndex, taskStatus, taskProgress
    lv3LeafName = GetTaskName(streamIndex, 3, phaseName)
    taskHours = GetLeafHours(phaseIndex, leafIndex)
    SetTaskRow ws, rowIndex, "E", lv3LeafName, GetLeafDetail(streamIndex, 2, phaseName), _
               taskStatus, taskProgress, streamOwner, InazumaGantt_v3.FormatDevelopmentHours(taskHours), _
               AddBusinessDays(streamStart, 3), AddBusinessDays(streamEnd, 0)
    rowIndex = rowIndex + 1
End Sub

Private Sub SetTaskRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal taskColumn As String, _
                       ByVal taskName As String, ByVal detailText As String, ByVal statusText As String, _
                       ByVal progressValue As Double, ByVal ownerName As String, ByVal hoursText As String, _
                       ByVal planStart As Variant, ByVal planEnd As Variant)
    ws.Cells(rowIndex, taskColumn).Value = taskName
    ws.Cells(rowIndex, InazumaGantt_v3.COL_TASK_DETAIL).Value = detailText
    ws.Cells(rowIndex, InazumaGantt_v3.COL_STATUS).Value = statusText
    ws.Cells(rowIndex, InazumaGantt_v3.COL_PROGRESS).Value = progressValue
    ws.Cells(rowIndex, InazumaGantt_v3.COL_ASSIGNEE).Value = ownerName

    If hoursText <> "" Then ws.Cells(rowIndex, InazumaGantt_v3.COL_DEV_LT).Value = hoursText
    If IsDate(planStart) Then ws.Cells(rowIndex, InazumaGantt_v3.COL_START_PLAN).Value = CDate(planStart)
    If IsDate(planEnd) Then ws.Cells(rowIndex, InazumaGantt_v3.COL_END_PLAN).Value = CDate(planEnd)
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
        Case 1: GetStreamName = phaseName & " 構想トラック"
        Case 2: GetStreamName = phaseName & " 実装トラック"
        Case Else: GetStreamName = phaseName & " 展開トラック"
    End Select
End Function

Private Function GetStreamDetail(ByVal streamIndex As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1: GetStreamDetail = SHOWCASE_APP_NAME & " の " & phaseName & " に必要な要件と成功条件を定義する。"
        Case 2: GetStreamDetail = phaseName & " を実装と体験品質へ落とし込む。"
        Case Else: GetStreamDetail = phaseName & " を安定運用と展開へつなげる。"
    End Select
End Function

Private Function GetTaskName(ByVal streamIndex As Long, ByVal slotIndex As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1
            Select Case slotIndex
                Case 1: GetTaskName = phaseName & " 戦略整理"
                Case 2: GetTaskName = phaseName & " 経営レビュー"
                Case Else: GetTaskName = phaseName & " 成果指標設計"
            End Select
        Case 2
            Select Case slotIndex
                Case 1: GetTaskName = phaseName & " 体験導線設計"
                Case 2: GetTaskName = phaseName & " プロトタイプ反復"
                Case Else: GetTaskName = phaseName & " 連携設計"
            End Select
        Case Else
            Select Case slotIndex
                Case 1: GetTaskName = phaseName & " 展開リハーサル"
                Case 2: GetTaskName = phaseName & " 当日運営訓練"
                Case Else: GetTaskName = phaseName & " 利用定着パック"
            End Select
    End Select
End Function

Private Function GetLeafDetail(ByVal streamIndex As Long, ByVal leafSlot As Long, ByVal phaseName As String) As String
    Select Case streamIndex
        Case 1
            If leafSlot = 1 Then
                GetLeafDetail = SHOWCASE_APP_NAME & " の " & phaseName & " で核になる仮説や設計を磨き込む。"
            Else
                GetLeafDetail = "会話力向上の成果指標、期待値、判断基準を揃える。"
            End If
        Case 2
            If leafSlot = 1 Then
                GetLeafDetail = "主要な学習体験を素早く試作し、魅力と違和感を見える化する。"
            Else
                GetLeafDetail = "アプリ、API、分析基盤の接続品質を設計する。"
            End If
        Case Else
            If leafSlot = 1 Then
                GetLeafDetail = "本番やベータに向けて運営訓練を行い、立ち上がりを安定させる。"
            Else
                GetLeafDetail = "CS、マーケ、運営が迷わない展開資料と支援物を整える。"
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
