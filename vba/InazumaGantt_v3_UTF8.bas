Attribute VB_Name = "InazumaGantt_v3"
Option Explicit

' ==========================================
'  イナズマガントチャート - 設定エリア
' ==========================================
' レイアウト:
' A: LV(階層) | B: No. | C: TASK | D-F: (タスク用スペース)
' G: タスクの詳細 | H: 状況 | I: 進捗率 | J: 担当
' K: 開発LT | L: 開始予定 | M: 完了予定 | N: 開始実績 | O: 完了実績
' P以降: ガントチャート領域 (日付)

Public Const COL_HIERARCHY As String = "A"   ' LV(階層)
Public Const COL_NO As String = "B"          ' No.
Public Const COL_TASK As String = "C"        ' TASK
' D-F列はタスク用のスペース（幅広め）
Public Const COL_TASK_DETAIL As String = "G" ' タスクの詳細
Public Const COL_STATUS As String = "H"      ' 状況
Public Const COL_PROGRESS As String = "I"    ' 進捗率
Public Const COL_ASSIGNEE As String = "J"    ' 担当
Public Const COL_DEV_LT As String = "K"      ' 開発LT
Public Const COL_START_PLAN As String = "L"  ' 開始予定
Public Const COL_END_PLAN As String = "M"    ' 完了予定
Public Const COL_START_ACTUAL As String = "N" ' 開始実績
Public Const COL_END_ACTUAL As String = "O"  ' 完了実績

Public Const COL_GANTT_START As String = "P"  ' ガントチャートの開始列
Public Const ROW_TITLE As Long = 1            ' タイトル行
Public Const ROW_WEEK_HEADER As Long = 6      ' 週ヘッダー行
Public Const ROW_DATE_HEADER As Long = 7      ' 日付行（ガント）
Public Const ROW_HEADER As Long = 8           ' 曜日行（ガント）/ 項目ヘッダー行（A-O列）
Public Const ROW_DATA_START As Long = 9       ' データ開始行
Public Const GANTT_DAYS As Long = 120         ' ガントチャートの日数
Public Const DATA_ROWS_DEFAULT As Long = 1000  ' 初期入力範囲の行数

Public Const GUIDE_SHEET_NAME As String = "InazumaGantt_説明"
Public Const MAIN_SHEET_NAME As String = "InazumaGantt_v3"
Public Const SETTINGS_SHEET_NAME As String = "設定マスタ"  ' v3
Public Const HOLIDAY_DATA_START_ROW As Long = 16  ' 設定マスタ内の祝日データ開始行
Public Const GUIDE_LEGEND_START_CELL As String = "E1"
Public Const CELL_PROJECT_START As String = "L2"
Public Const CELL_DISPLAY_WEEK As String = "L3"
Public Const CELL_TODAY As String = "L4"
Public Const COL_DATA_END As String = "O"
Private Const BACKUP_SHEET_NAME As String = "WBS_Backup_v3"
Private Const BACKUP_TEMP_SHEET_NAME As String = "_WBS_Backup_v3_tmp"
Private Const BACKUP_OLD_SHEET_NAME As String = "_WBS_Backup_v3_old"
Private Const RESTORE_ROLLBACK_SHEET_NAME As String = "_WBS_Restore_v3_rollback"
Private Const BACKUP_META_LABEL_START As String = "Q1"
Private Const BACKUP_META_VALUE_START As String = "R1"

' 色設定
Public Const COLOR_PLAN As Long = 16119285       ' RGB(245,245,245) 限りなく白に近い灰色
Public Const COLOR_PROGRESS As Long = 9851952    ' RGB(48,84,150) 紺色
Public Const COLOR_HOLIDAY As Long = 5263430     ' RGB(70,70,80) 濃い灰色（休日祝日）
Public Const COLOR_ROW_BAND As Long = 16316664
Public Const COLOR_ACTUAL As Long = 5287936      ' RGB(0,176,80) 緑色
Public Const COLOR_ACTUAL_OVERRUN As Long = 52377  ' RGB(137,204,0) 黄緑色（超過分）
Public Const COLOR_TODAY As Long = 255           ' RGB(255,0,0) 赤
Public Const COLOR_WARN As Long = 13434879
Public Const COLOR_ERROR As Long = 13553151
Public Const COLOR_INAZUMA As Long = 42495       ' RGB(255,165,0) オレンジ
Public Const COLOR_HEADER_BG As Long = 12874308
Public Const COLOR_GANTT_HEADER As Long = 8421504
Public Const COLOR_WEEKEND As Long = 5263430     ' RGB(70,70,80) 濃い灰色
Public Const TODAY_LINE_WEIGHT As Double = 2
Public Const INAZUMA_LINE_WEIGHT As Double = 2
Public Const ACTUAL_LINE_WEIGHT As Double = 4
Public Const STATUS_NOT_STARTED As String = "未着手"
Public Const STATUS_IN_PROGRESS As String = "進行中"
Public Const STATUS_COMPLETED As String = "完了"
Public Const STATUS_ON_HOLD As String = "保留"
Public Const DEV_HOURS_NUMBER_FORMAT As String = "0.0""h"""
Public Const AUXILIARY_TASK_PLACEHOLDER As String = "（補助情報のみ）"
Private Const HIDDEN_VALUE_NUMBER_FORMAT As String = ";;;"
Public Const SETTINGS_ROW_BULK_EDIT_MODE As Long = 9
Public Const SETTINGS_ROW_WBS_SUMMARY_DEPTH As Long = 11
Public Const SETTINGS_ROW_AUTOMATION_MODE As Long = 12
Private Const BULK_EDIT_STATUS_RANGE As String = "A3:J3"
Private Const LOG_SHEET_NAME As String = "_InazumaGantt_Log"
Private Const GANTT_DOEVENTS_INTERVAL As Long = 25
Private mIsRefreshingGantt As Boolean
Private Const LOG_MAX_ROWS As Long = 2000
Private Const BULK_EDIT_STATE_SHEET_NAME As String = "_InazumaBulkEditState"
Private Const BULK_EDIT_RECONCILE_FULL_THRESHOLD As Long = 80
Private Const BULK_EDIT_SIGNATURE_FIRST_COLUMN As Long = 1
Private Const BULK_EDIT_SIGNATURE_LAST_COLUMN As Long = 15
Private mIsTogglingBulkEdit As Boolean

Private Function GetMainWorksheet() As Worksheet
    On Error Resume Next
    Set GetMainWorksheet = ThisWorkbook.Worksheets(MAIN_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function BuildBulkEditIndicatorText(Optional ByVal note As String = "") As String
    Dim baseText As String

    If IsBulkEditModeEnabled() Then
        baseText = "高速入力 ON: Ctrl+Z 優先のため自動更新を停止中です。作業後に OFF へ戻すか「ガント更新」で再整合します。"
    Else
        baseText = "高速入力 OFF: 通常モードです。入力に合わせて計算と見た目をその場で更新します。"
    End If

    If Trim$(note) <> "" Then
        BuildBulkEditIndicatorText = note & " " & baseText
    Else
        BuildBulkEditIndicatorText = baseText
    End If
End Function

Private Sub UpdateBulkEditModeIndicator(ByVal ws As Worksheet, Optional ByVal note As String = "")
    Dim statusRange As Range

    If ws Is Nothing Then Exit Sub

    Set statusRange = ws.Range(BULK_EDIT_STATUS_RANGE)
    On Error Resume Next
    If statusRange.MergeCells Then statusRange.UnMerge
    On Error GoTo 0
    statusRange.Merge
    statusRange.Value = BuildBulkEditIndicatorText(note)
    statusRange.WrapText = False
    statusRange.HorizontalAlignment = xlLeft
    statusRange.VerticalAlignment = xlCenter
    statusRange.Font.Size = 9
    statusRange.Font.Bold = True

    If IsBulkEditModeEnabled() Then
        statusRange.Interior.Color = RGB(252, 228, 214)
        statusRange.Font.Color = RGB(156, 0, 6)
    Else
        statusRange.Interior.Color = RGB(226, 239, 218)
        statusRange.Font.Color = RGB(0, 97, 0)
    End If
End Sub

Private Function RepairBulkEditRuntimeState(Optional ByVal ws As Worksheet = Nothing) As String
    Dim storedEnabled As Boolean

    EnsureSettingsSheet
    storedEnabled = IsBulkEditModeEnabled()
    If ws Is Nothing Then Set ws = GetMainWorksheet()

    If storedEnabled And Application.EnableEvents Then
        SetBulkEditMode False
        RepairBulkEditRuntimeState = "高速入力状態を自動修復し、通常モードに戻しました。"
    ElseIf (Not storedEnabled) And (Not Application.EnableEvents) Then
        Application.EnableEvents = True
        RepairBulkEditRuntimeState = "イベント停止状態を自動修復しました。"
    End If
End Function

Public Function IsBulkEditModeEnabledAfterRuntimeRepair() As Boolean
    Call RepairBulkEditRuntimeState(GetMainWorksheet())
    IsBulkEditModeEnabledAfterRuntimeRepair = IsBulkEditModeEnabled()
End Function

Private Function GetBackupWorksheet() As Worksheet
    On Error Resume Next
    Set GetBackupWorksheet = ThisWorkbook.Worksheets(BACKUP_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function GetWorksheetByName(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Sub CaptureApplicationState(ByRef prevCalc As XlCalculation, ByRef prevEvents As Boolean, _
                                    ByRef prevScreenUpdating As Boolean, ByRef stateCaptured As Boolean)
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating
    stateCaptured = True
End Sub

Private Sub RestoreApplicationState(ByVal prevCalc As XlCalculation, ByVal prevEvents As Boolean, _
                                    ByVal prevScreenUpdating As Boolean, ByVal stateCaptured As Boolean, _
                                    Optional ByVal forceEventsEnabled As Boolean = False)
    On Error Resume Next
    If stateCaptured Then
        If forceEventsEnabled Then
            Application.EnableEvents = True
        Else
            Application.EnableEvents = prevEvents
        End If
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreenUpdating
    Else
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End If
    On Error GoTo 0
End Sub

Private Sub DeleteWorksheetIfExists(ByVal sheetName As String, Optional ByVal fallbackSheetName As String = "")
    Dim ws As Worksheet
    Dim prevAlerts As Boolean

    Set ws = GetWorksheetByName(sheetName)
    If ws Is Nothing Then Exit Sub

    prevAlerts = Application.DisplayAlerts
    On Error GoTo CleanUp
    Application.DisplayAlerts = False
    If Len(fallbackSheetName) > 0 Then
        On Error Resume Next
        ThisWorkbook.Worksheets(fallbackSheetName).Activate
        Err.Clear
        On Error GoTo CleanUp
    End If
    ws.Delete

CleanUp:
    Application.DisplayAlerts = prevAlerts
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function CreateTemporaryWorksheet(ByVal tempSheetName As String) As Worksheet
    DeleteWorksheetIfExists tempSheetName, MAIN_SHEET_NAME
    Set CreateTemporaryWorksheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    CreateTemporaryWorksheet.Name = tempSheetName
End Function

Private Function GetOrCreateBackupWorksheet() As Worksheet
    Dim wsBackup As Worksheet

    Set wsBackup = GetBackupWorksheet()
    If wsBackup Is Nothing Then
        Set wsBackup = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsBackup.Name = BACKUP_SHEET_NAME
    End If

    wsBackup.Visible = xlSheetVisible
    Set GetOrCreateBackupWorksheet = wsBackup
End Function

Private Sub ReplaceBackupWorksheetWithStaged(ByVal stagedBackup As Worksheet)
    Dim oldBackup As Worksheet
    Dim oldRenamed As Boolean
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String

    If stagedBackup Is Nothing Then Err.Raise vbObjectError + 1200, "ReplaceBackupWorksheetWithStaged", "Backup staging sheet was not created."

    DeleteWorksheetIfExists BACKUP_OLD_SHEET_NAME, MAIN_SHEET_NAME
    Set oldBackup = GetBackupWorksheet()
    If Not oldBackup Is Nothing Then
        oldBackup.Name = BACKUP_OLD_SHEET_NAME
        oldRenamed = True
    End If

    On Error GoTo Rollback
    stagedBackup.Name = BACKUP_SHEET_NAME
    stagedBackup.Visible = xlSheetVisible
    On Error Resume Next
    DeleteWorksheetIfExists BACKUP_OLD_SHEET_NAME, MAIN_SHEET_NAME
    On Error GoTo 0
    Exit Sub

Rollback:
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    If oldRenamed Then
        On Error Resume Next
        oldBackup.Name = BACKUP_SHEET_NAME
        On Error GoTo 0
    End If
    Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub ClearAllShapes(ByVal ws As Worksheet)
    Dim shp As Shape

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0
End Sub

Private Sub CopyWbsSnapshot(ByVal sourceWs As Worksheet, ByVal targetWs As Worksheet, ByVal lastRow As Long, _
                            Optional ByVal clearWholeSheet As Boolean = False)
    Dim copyRange As Range
    Dim targetRange As Range
    Dim rowIndex As Long

    If sourceWs Is Nothing Or targetWs Is Nothing Then Exit Sub
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START

    Set copyRange = sourceWs.Range("A1:" & COL_DATA_END & lastRow)
    Set targetRange = targetWs.Range("A1:" & COL_DATA_END & lastRow)

    If clearWholeSheet Then
        targetWs.Cells.Clear
        ClearAllShapes targetWs
    End If

    copyRange.Copy
    targetWs.Range("A1").PasteSpecial xlPasteAll
    targetWs.Range("A1").PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False
    targetRange.Value = targetRange.Value

    For rowIndex = 1 To lastRow
        targetWs.Rows(rowIndex).RowHeight = sourceWs.Rows(rowIndex).RowHeight
    Next rowIndex
End Sub

Private Sub WriteBackupMetadata(ByVal wsBackup As Worksheet, ByVal sourceWs As Worksheet, ByVal lastRow As Long)
    If wsBackup Is Nothing Then Exit Sub

    With wsBackup
        .Range(BACKUP_META_LABEL_START).Value = "Backup updated"
        .Range(BACKUP_META_VALUE_START).Value = Now
        .Range(BACKUP_META_VALUE_START).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        .Range("Q2").Value = "Source sheet"
        .Range("R2").Value = sourceWs.Name
        .Range("Q3").Value = "Snapshot last row"
        .Range("R3").Value = lastRow
        .Range("Q1:R3").Font.Size = 9
        .Range("Q1:Q3").Font.Bold = True
    End With
End Sub

Private Function GetBackupTimestampLabel(ByVal wsBackup As Worksheet) As String
    Dim rawValue As Variant

    If wsBackup Is Nothing Then Exit Function

    rawValue = wsBackup.Range(BACKUP_META_VALUE_START).Value
    If IsDate(rawValue) Then
        GetBackupTimestampLabel = Format$(CDate(rawValue), "yyyy/mm/dd hh:nn:ss")
    Else
        GetBackupTimestampLabel = Trim$(CStr(rawValue))
    End If
End Function

Private Function BuildRestoreConfirmationMessage(ByVal wsBackup As Worksheet, ByVal backupLastRow As Long) As String
    Dim timestampLabel As String

    timestampLabel = GetBackupTimestampLabel(wsBackup)
    If Trim$(timestampLabel) = "" Then timestampLabel = "不明"

    BuildRestoreConfirmationMessage = "現在の WBS をバックアップで上書きします。" & vbCrLf & vbCrLf & _
                                      "バックアップ更新日時: " & timestampLabel & vbCrLf & _
                                      "復元対象シート: " & wsBackup.Name & vbCrLf & _
                                      "バックアップ最終行: " & backupLastRow & vbCrLf & vbCrLf & _
                                      "続行しますか？"
End Function

Private Sub DeleteManagedSheetShapes(ByVal ws As Worksheet)
    Dim shapeIndex As Long

    If ws Is Nothing Then Exit Sub

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        With ws.Shapes(shapeIndex)
            If Left(.Name, 4) = "Bar_" Or Left(.Name, 6) = "Today_" Or _
               Left(.Name, 8) = "Inazuma_" Or Left(.Name, 4) = "Btn_" Then
                .Delete
            End If
        End With
    Next shapeIndex
End Sub

Private Sub ClearMainWorksheetForRestore(ByVal ws As Worksheet, ByVal clearEndRow As Long)
    Dim ganttStartCol As Long
    Dim ganttEndCol As Long

    If ws Is Nothing Then Exit Sub
    If clearEndRow < ROW_DATA_START Then clearEndRow = ROW_DATA_START

    ws.Range("A1:" & COL_DATA_END & clearEndRow).Clear

    ganttStartCol = ws.Columns(COL_GANTT_START).Column
    ganttEndCol = ganttStartCol + GANTT_DAYS - 1
    ws.Range(ws.Cells(ROW_WEEK_HEADER, ganttStartCol), ws.Cells(clearEndRow, ganttEndCol)).Clear

    DeleteManagedSheetShapes ws
End Sub

Private Function GetSnapshotClearEndRow(ByVal ws As Worksheet, ByVal snapshotLastRow As Long) As Long
    Dim currentLastRow As Long
    Dim defaultLastRow As Long

    currentLastRow = GetLastDataRow(ws)
    If currentLastRow < ROW_DATA_START Then currentLastRow = ROW_DATA_START

    defaultLastRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    GetSnapshotClearEndRow = MaxRow(defaultLastRow, currentLastRow)
    GetSnapshotClearEndRow = MaxRow(GetSnapshotClearEndRow, snapshotLastRow)
End Function

Private Sub ApplyHierarchyColorsSilently()
    On Error GoTo CleanUp

    Dim prevAlerts As Boolean

    Err.Clear
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    HierarchyColor.SetupHierarchyColors

CleanUp:
    Application.DisplayAlerts = prevAlerts
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub ReconcileDeferredTaskState(ByVal ws As Worksheet, Optional ByVal affectedRows As Object = Nothing)
    Dim lastRow As Long
    Dim ganttStartDate As Date
    Dim ganttStartCol As Long
    Dim startedAt As Double
    Dim phaseStartedAt As Double
    Dim changedRowCount As Long
    Dim reconcileMode As String

    If ws Is Nothing Then Exit Sub

    startedAt = Timer
    If affectedRows Is Nothing Then
        changedRowCount = -1
    Else
        changedRowCount = affectedRows.Count
    End If

    phaseStartedAt = Timer
    ClearLiveTaskLevelHintFormulas ws
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    LogGanttPhase ws, "prepare", phaseStartedAt, "rows=" & CStr(lastRow) & ", changedRows=" & CStr(changedRowCount)

    phaseStartedAt = Timer
    SetGanttRefreshStatus "ガント更新: 入力値を整えています..."
    NormalizeTaskStatusAndProgressRange ws, ROW_DATA_START, lastRow
    SetGanttRefreshStatus "ガント更新: LVとNo.を整えています..."
    AutoDetectTaskLevelsInRange ws, ROW_DATA_START, lastRow
    RenumberRowsForWorksheet ws
    ganttStartCol = ws.Columns(COL_GANTT_START).Column
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        ganttStartDate = Date
    End If
    LogGanttPhase ws, "input-normalize", phaseStartedAt, "rows=" & CStr(lastRow)

    phaseStartedAt = Timer
    SetGanttRefreshStatus "ガント更新: 日付と罫線を整えています..."
    RegenerateDateHeaders ws
    ClearGanttColors ws, lastRow, ganttStartCol
    ApplyGanttBorders ws, lastRow
    DrawWeekSeparators ws, lastRow
    ApplyWeekendColors ws, lastRow, ganttStartDate, ganttStartCol
    ApplyDataValidationAndFormats ws, lastRow
    ApplyHolidayColors ws, lastRow
    LogGanttPhase ws, "headers-formats", phaseStartedAt, "rows=" & CStr(lastRow) & ", days=" & CStr(GANTT_DAYS)

    SetGanttRefreshStatus "ガント更新: 親タスクを集計しています..."
    phaseStartedAt = Timer
    If affectedRows Is Nothing Then
        reconcileMode = "full"
        WBSParentRollup.RefreshAllParentTasks ws
        LogGanttPhase ws, "parent-rollup", phaseStartedAt, "mode=" & reconcileMode & ", rows=" & CStr(lastRow)
        phaseStartedAt = Timer
        WBSParentRollup.RefreshTaskAlertMarkers ws
        LogGanttPhase ws, "alert-markers", phaseStartedAt, "mode=" & reconcileMode & ", rows=" & CStr(lastRow)
    ElseIf affectedRows.Count = 0 Then
        reconcileMode = "nochange"
        LogGanttPhase ws, "parent-rollup", phaseStartedAt, "mode=" & reconcileMode & ", skipped=true"
        phaseStartedAt = Timer
        LogGanttPhase ws, "alert-markers", phaseStartedAt, "mode=" & reconcileMode & ", skipped=true"
    ElseIf affectedRows.Count <= BULK_EDIT_RECONCILE_FULL_THRESHOLD Then
        reconcileMode = "incremental"
        WBSParentRollup.RecalculateTaskRowsAndAncestors ws, affectedRows
        LogGanttPhase ws, "parent-rollup", phaseStartedAt, "mode=" & reconcileMode & ", affectedRows=" & CStr(affectedRows.Count)
        phaseStartedAt = Timer
        WBSParentRollup.RefreshTaskAlertMarkersForRowsAndAncestors ws, affectedRows
        LogGanttPhase ws, "alert-markers", phaseStartedAt, "mode=" & reconcileMode & ", affectedRows=" & CStr(affectedRows.Count)
    Else
        reconcileMode = "full-threshold"
        WBSParentRollup.RefreshAllParentTasks ws
        LogGanttPhase ws, "parent-rollup", phaseStartedAt, "mode=" & reconcileMode & ", affectedRows=" & CStr(affectedRows.Count)
        phaseStartedAt = Timer
        WBSParentRollup.RefreshTaskAlertMarkers ws
        LogGanttPhase ws, "alert-markers", phaseStartedAt, "mode=" & reconcileMode & ", affectedRows=" & CStr(affectedRows.Count)
    End If

    phaseStartedAt = Timer
    If reconcileMode = "incremental" Then
        SetGanttRefreshStatus "ガント更新: 変更行の表示を整えています..."
        RefreshAffectedRowsAfterBulkEdit ws, affectedRows
        LogGanttPhase ws, "row-display", phaseStartedAt, "mode=" & reconcileMode & ", affectedRows=" & CStr(affectedRows.Count)
    ElseIf reconcileMode <> "nochange" Then
        SetGanttRefreshStatus "ガント更新: 表示を整えています..."
        ApplyHierarchyColorsSilently
        LogGanttPhase ws, "row-display", phaseStartedAt, "mode=" & reconcileMode & ", rows=" & CStr(lastRow)
    Else
        LogGanttPhase ws, "row-display", phaseStartedAt, "mode=" & reconcileMode & ", skipped=true"
    End If

    phaseStartedAt = Timer
    If reconcileMode <> "nochange" Then
        SetGanttRefreshStatus "ガント更新: ガント線を描画しています..."
        DrawGanttBars True
        LogGanttPhase ws, "shape-redraw", phaseStartedAt, "mode=" & reconcileMode & ", rows=" & CStr(lastRow)
    Else
        LogGanttPhase ws, "shape-redraw", phaseStartedAt, "mode=" & reconcileMode & ", skipped=true"
    End If

    LogGanttReconcile ws, reconcileMode, changedRowCount, lastRow, startedAt
End Sub

Private Function BuildLiveTaskLevelFormulaR1C1() As String
    BuildLiveTaskLevelFormulaR1C1 = "=IF(LEN(TRIM(RC[5]))>0,4,IF(LEN(TRIM(RC[4]))>0,3,IF(LEN(TRIM(RC[3]))>0,2,IF(AND(LEN(TRIM(RC[2]))>0,TRIM(RC[2])<>""（補助情報のみ）"",TRIM(RC[2])<>""! （補助情報のみ）"",TRIM(RC[2])<>""!! （補助情報のみ）"",TRIM(RC[2])<>""!"",TRIM(RC[2])<>""!!""),1,""""))))"
End Function

Private Sub PrepareLiveTaskLevelHintsForBulkEdit(ByVal ws As Worksheet)
    Dim r As Long
    Dim lastSupportedRow As Long

    If ws Is Nothing Then Exit Sub

    lastSupportedRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    For r = ROW_DATA_START To lastSupportedRow
        If Trim$(CStr(ws.Cells(r, COL_HIERARCHY).Value)) = "" Then
            If Not HasPrimaryTaskContentInRow(ws, r) Then
                ws.Cells(r, COL_HIERARCHY).FormulaR1C1 = BuildLiveTaskLevelFormulaR1C1()
                ws.Cells(r, COL_HIERARCHY).NumberFormat = "General"
            End If
        End If
    Next r
End Sub

Private Sub ClearLiveTaskLevelHintFormulas(ByVal ws As Worksheet)
    Dim lastSupportedRow As Long
    Dim levelRange As Range

    If ws Is Nothing Then Exit Sub

    lastSupportedRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    Set levelRange = ws.Range(COL_HIERARCHY & ROW_DATA_START & ":" & COL_HIERARCHY & lastSupportedRow)
    levelRange.Value = levelRange.Value
End Sub

Private Function GetBulkEditStateWorksheet(Optional ByVal createIfMissing As Boolean = False) As Worksheet
    Dim wsState As Worksheet

    On Error Resume Next
    Set wsState = ThisWorkbook.Worksheets(BULK_EDIT_STATE_SHEET_NAME)
    On Error GoTo 0

    If wsState Is Nothing And createIfMissing Then
        Set wsState = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsState.Name = BULK_EDIT_STATE_SHEET_NAME
    End If

    If Not wsState Is Nothing Then
        On Error Resume Next
        wsState.Visible = xlSheetVeryHidden
        On Error GoTo 0
    End If

    Set GetBulkEditStateWorksheet = wsState
End Function

Private Function GetBulkEditSnapshotLastRow() As Long
    Dim wsState As Worksheet
    Dim storedValue As Variant

    Set wsState = GetBulkEditStateWorksheet(False)
    If wsState Is Nothing Then Exit Function

    storedValue = wsState.Range("B2").Value
    If IsNumeric(storedValue) Then GetBulkEditSnapshotLastRow = CLng(storedValue)
End Function

Private Function GetBulkEditScanEndRow(ByVal ws As Worksheet) As Long
    Dim snapshotLastRow As Long

    If ws Is Nothing Then Exit Function

    GetBulkEditScanEndRow = MaxRow(GetLastDataRow(ws), ROW_DATA_START + DATA_ROWS_DEFAULT - 1)
    snapshotLastRow = GetBulkEditSnapshotLastRow()
    If snapshotLastRow > GetBulkEditScanEndRow Then GetBulkEditScanEndRow = snapshotLastRow
End Function

Private Function BulkEditValueText(ByVal targetCell As Range) As String
    Dim cellValue As Variant

    On Error GoTo ErrorHandler
    cellValue = targetCell.Value2

    If IsError(cellValue) Then
        BulkEditValueText = "#ERR:" & CStr(targetCell.Text)
    ElseIf IsEmpty(cellValue) Then
        BulkEditValueText = ""
    Else
        BulkEditValueText = CStr(cellValue)
    End If

    BulkEditValueText = Replace(BulkEditValueText, Chr$(29), " ")
    BulkEditValueText = Replace(BulkEditValueText, Chr$(30), " ")
    Exit Function

ErrorHandler:
    BulkEditValueText = "#ERR"
End Function

Private Function BuildBulkEditRowSignature(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    Dim colIndex As Long
    Dim signatureText As String

    If ws Is Nothing Then Exit Function
    If targetRow < ROW_DATA_START Then Exit Function

    For colIndex = BULK_EDIT_SIGNATURE_FIRST_COLUMN To BULK_EDIT_SIGNATURE_LAST_COLUMN
        signatureText = signatureText & Chr$(30) & BulkEditValueText(ws.Cells(targetRow, colIndex))
    Next colIndex

    BuildBulkEditRowSignature = signatureText
End Function

Private Sub CaptureBulkEditSnapshot(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim wsState As Worksheet
    Dim lastRow As Long
    Dim rowCount As Long
    Dim snapshotData() As Variant
    Dim r As Long
    Dim rowIndex As Long
    Dim targetName As String

    If ws Is Nothing Then Exit Sub
    targetName = ws.Name

    Set wsState = GetBulkEditStateWorksheet(True)
    If wsState Is Nothing Then Exit Sub

    lastRow = GetBulkEditScanEndRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    rowCount = lastRow - ROW_DATA_START + 1
    ReDim snapshotData(1 To rowCount, 1 To 2)

    Application.StatusBar = "高速入力: 差分確認用スナップショットを作成しています..."
    For r = ROW_DATA_START To lastRow
        MaybeYieldDuringGanttRefresh r - ROW_DATA_START + 1, rowCount, "高速入力スナップショット"
        rowIndex = r - ROW_DATA_START + 1
        snapshotData(rowIndex, 1) = r
        snapshotData(rowIndex, 2) = BuildBulkEditRowSignature(ws, r)
    Next r

    wsState.Cells.Clear
    wsState.Range("A1:B1").Value = Array("capturedAt", Now)
    wsState.Range("A2:B2").Value = Array("lastRow", lastRow)
    wsState.Range("A4:B4").Value = Array("row", "signature")
    wsState.Range("A5").Resize(rowCount, 2).Value = snapshotData
    wsState.Visible = xlSheetVeryHidden

    LogAutomationEvent "BulkEditSnapshot", "rows=" & CStr(rowCount) & ", lastRow=" & CStr(lastRow), targetName
    Exit Sub

ErrorHandler:
    On Error Resume Next
    LogAutomationEvent "BulkEditSnapshotError", Err.Description, targetName
End Sub

Private Sub AddBulkEditAffectedRow(ByVal affectedRows As Object, ByVal targetRow As Long, ByVal scanEndRow As Long)
    If affectedRows Is Nothing Then Exit Sub
    If targetRow < ROW_DATA_START Then Exit Sub
    If scanEndRow >= ROW_DATA_START And targetRow > scanEndRow Then Exit Sub
    affectedRows(CStr(targetRow)) = True
End Sub

Private Function DetectBulkEditAffectedRows(ByVal ws As Worksheet) As Object
    On Error GoTo ErrorHandler

    Dim wsState As Worksheet
    Dim previousByRow As Object
    Dim affectedRows As Object
    Dim stateData As Variant
    Dim lastSnapshotRow As Long
    Dim scanEndRow As Long
    Dim scanCount As Long
    Dim i As Long
    Dim r As Long
    Dim rowKey As String
    Dim currentSignature As String
    Dim previousSignature As String
    Dim targetName As String

    If ws Is Nothing Then Exit Function
    targetName = ws.Name

    Set wsState = GetBulkEditStateWorksheet(False)
    If wsState Is Nothing Then Exit Function

    lastSnapshotRow = wsState.Cells(wsState.Rows.Count, "A").End(xlUp).Row
    If lastSnapshotRow < 5 Then Exit Function

    Set previousByRow = CreateObject("Scripting.Dictionary")
    Set affectedRows = CreateObject("Scripting.Dictionary")
    stateData = wsState.Range("A5:B" & lastSnapshotRow).Value2

    For i = 1 To UBound(stateData, 1)
        If IsNumeric(stateData(i, 1)) Then
            previousByRow(CStr(CLng(stateData(i, 1)))) = CStr(stateData(i, 2))
        End If
    Next i

    scanEndRow = GetBulkEditScanEndRow(ws)
    If scanEndRow < ROW_DATA_START Then scanEndRow = ROW_DATA_START
    scanCount = scanEndRow - ROW_DATA_START + 1

    Application.StatusBar = "高速入力OFF: 変更行を確認しています..."
    For r = ROW_DATA_START To scanEndRow
        MaybeYieldDuringGanttRefresh r - ROW_DATA_START + 1, scanCount, "高速入力差分確認"
        rowKey = CStr(r)
        currentSignature = BuildBulkEditRowSignature(ws, r)
        previousSignature = ""
        If previousByRow.Exists(rowKey) Then previousSignature = CStr(previousByRow(rowKey))

        If (Not previousByRow.Exists(rowKey)) Or previousSignature <> currentSignature Then
            AddBulkEditAffectedRow affectedRows, r - 1, scanEndRow
            AddBulkEditAffectedRow affectedRows, r, scanEndRow
            AddBulkEditAffectedRow affectedRows, r + 1, scanEndRow
        End If
    Next r

    LogAutomationEvent "BulkEditChangedRows", "affectedRows=" & CStr(affectedRows.Count) & ", scanRows=" & CStr(scanCount), targetName
    Set DetectBulkEditAffectedRows = affectedRows
    Exit Function

ErrorHandler:
    On Error Resume Next
    LogAutomationEvent "BulkEditDiffError", Err.Description, targetName
End Function

Private Sub ClearBulkEditSnapshot()
    On Error GoTo Fallback

    Dim wsState As Worksheet
    Dim prevAlerts As Boolean

    Set wsState = GetBulkEditStateWorksheet(False)
    If wsState Is Nothing Then Exit Sub

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsState.Visible = xlSheetVisible
    wsState.Delete
    Application.DisplayAlerts = prevAlerts
    Exit Sub

Fallback:
    On Error Resume Next
    Application.DisplayAlerts = prevAlerts
    If Not wsState Is Nothing Then
        wsState.Cells.Clear
        wsState.Visible = xlSheetVeryHidden
    End If
End Sub

Private Function ElapsedSeconds(ByVal startedAt As Double) As Double
    ElapsedSeconds = Timer - startedAt
    If ElapsedSeconds < 0 Then ElapsedSeconds = ElapsedSeconds + 86400#
End Function

Private Sub RefreshAffectedRowsAfterBulkEdit(ByVal ws As Worksheet, ByVal affectedRows As Object)
    Dim rowKey As Variant
    Dim rowIndex As Long
    Dim rowCount As Long

    If ws Is Nothing Then Exit Sub
    If affectedRows Is Nothing Then Exit Sub
    If affectedRows.Count = 0 Then Exit Sub

    rowCount = affectedRows.Count
    For Each rowKey In affectedRows.Keys
        rowIndex = rowIndex + 1
        MaybeYieldDuringGanttRefresh rowIndex, rowCount, "変更行の表示更新"
        RefreshTaskRowDisplayState ws, CLng(rowKey)
    Next rowKey
End Sub

Private Sub LogGanttReconcile(ByVal ws As Worksheet, ByVal reconcileMode As String, _
                              ByVal changedRowCount As Long, ByVal lastRow As Long, _
                              ByVal startedAt As Double)
    Dim details As String
    Dim targetName As String

    If Not ws Is Nothing Then targetName = ws.Name
    details = "mode=" & reconcileMode & _
              ", changedRows=" & CStr(changedRowCount) & _
              ", lastRow=" & CStr(lastRow) & _
              ", elapsedSec=" & Format$(ElapsedSeconds(startedAt), "0.00")
    LogAutomationEvent "GanttReconcile", details, targetName
End Sub

Private Sub LogGanttPhase(ByVal ws As Worksheet, ByVal phaseName As String, _
                          ByVal startedAt As Double, Optional ByVal details As String = "")
    Dim logDetails As String
    Dim targetName As String

    If Not ws Is Nothing Then targetName = ws.Name
    logDetails = "phase=" & phaseName & _
                 ", elapsedSec=" & Format$(ElapsedSeconds(startedAt), "0.00")
    If Trim$(details) <> "" Then logDetails = logDetails & ", " & details
    LogAutomationEvent "GanttPhase", logDetails, targetName
End Sub

Public Sub CancelDeferredBulkEditReconcile()
End Sub

Public Sub QueueDeferredBulkEditReconcile(Optional ByVal note As String = "")
    Dim ws As Worksheet

    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Exit Sub

    UpdateBulkEditModeIndicator ws, note
End Sub

Private Function RequireMainWorksheet(ByVal operationName As String, Optional ByVal requireActiveMainSheet As Boolean = False) As Worksheet
    Dim ws As Worksheet
    Set ws = GetMainWorksheet()

    If ws Is Nothing Then
        MsgBox "メインシート '" & MAIN_SHEET_NAME & "' が見つかりません。" & vbCrLf & _
               "先に RunSetupWizard を実行してください。", vbExclamation, operationName
        Exit Function
    End If

    If requireActiveMainSheet Then
        If ActiveSheet Is Nothing Or Not ActiveSheet Is ws Then
            MsgBox operationName & " は '" & MAIN_SHEET_NAME & "' シートを表示した状態で実行してください。", vbExclamation, operationName
            Exit Function
        End If
    End If

    Set RequireMainWorksheet = ws
End Function

Private Function TryParseProgressValue(ByVal progressValue As Variant, ByRef normalizedValue As Double) As Boolean
    Dim textValue As String

    If IsEmpty(progressValue) Then Exit Function

    textValue = Trim$(CStr(progressValue))
    If textValue = "" Then Exit Function

    textValue = Replace$(textValue, "%", "")
    If Not IsNumeric(textValue) Then Exit Function

    normalizedValue = CDbl(textValue)
    If normalizedValue < 0 Then Exit Function
    If normalizedValue > 1 Then normalizedValue = normalizedValue / 100
    If normalizedValue < 0 Or normalizedValue > 1 Then Exit Function

    TryParseProgressValue = True
End Function

Public Function NormalizeProgressValue(ByVal progressValue As Variant, Optional ByVal fallback As Double = 0) As Double
    Dim normalizedValue As Double

    If TryParseProgressValue(progressValue, normalizedValue) Then
        NormalizeProgressValue = normalizedValue
    Else
        NormalizeProgressValue = fallback
    End If
End Function

Public Function TryGetDateValue(ByVal rawValue As Variant, ByRef parsedDate As Date) As Boolean
    Dim textValue As String
    Dim serialValue As Double

    If IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function

    textValue = Trim$(CStr(rawValue))
    If textValue = "" Then Exit Function

    If IsDate(rawValue) Then
        parsedDate = CDate(rawValue)
        TryGetDateValue = True
        Exit Function
    End If

    If Not IsNumeric(rawValue) Then Exit Function

    serialValue = CDbl(rawValue)
    If serialValue <= 0 Or serialValue >= 2958466 Then Exit Function

    parsedDate = CDate(DateSerial(1899, 12, 30) + serialValue)
    TryGetDateValue = True
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

Public Function FormatDevelopmentHours(ByVal normalizedHours As Double) As String
    FormatDevelopmentHours = Format$(normalizedHours, "0.0") & "h"
End Function

Public Function GetDevelopmentHoursValue(ByVal hoursValue As Variant, Optional ByVal fallback As Double = 0) As Double
    Dim normalizedHours As Double

    If TryParseDevelopmentHours(hoursValue, normalizedHours) Then
        GetDevelopmentHoursValue = normalizedHours
    Else
        GetDevelopmentHoursValue = fallback
    End If
End Function

Private Function NormalizeStatusText(ByVal statusValue As Variant) As String
    Dim textValue As String

    textValue = Trim$(CStr(statusValue))
    Select Case textValue
        Case ""
            NormalizeStatusText = STATUS_NOT_STARTED
        Case STATUS_NOT_STARTED, STATUS_IN_PROGRESS, STATUS_COMPLETED, STATUS_ON_HOLD
            NormalizeStatusText = textValue
        Case Else
            NormalizeStatusText = STATUS_NOT_STARTED
    End Select
End Function

Public Sub SyncTaskStatusAndProgressRow(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim statusText As String
    Dim rawStatusText As String
    Dim progressRate As Double
    Dim progressText As String

    If ws Is Nothing Then Exit Sub
    If targetRow < ROW_DATA_START Then Exit Sub
    If Not HasTaskContentInRow(ws, targetRow) Then Exit Sub

    rawStatusText = Trim$(CStr(ws.Cells(targetRow, COL_STATUS).Value))
    statusText = NormalizeStatusText(rawStatusText)
    progressText = Trim$(CStr(ws.Cells(targetRow, COL_PROGRESS).Value))

    If progressText = "" Then
        progressRate = 0
    Else
        progressRate = NormalizeProgressValue(ws.Cells(targetRow, COL_PROGRESS).Value, 0)
    End If

    If rawStatusText = "" Then
        If progressRate <= 0 Then
            statusText = STATUS_NOT_STARTED
            progressRate = 0
        ElseIf progressRate >= 1 Then
            statusText = STATUS_COMPLETED
            progressRate = 1
        Else
            statusText = STATUS_IN_PROGRESS
        End If
        ws.Cells(targetRow, COL_STATUS).Value = statusText
        ws.Cells(targetRow, COL_PROGRESS).Value = progressRate
        Exit Sub
    End If

    Select Case statusText
        Case STATUS_COMPLETED
            progressRate = 1
        Case STATUS_NOT_STARTED
            progressRate = 0
        Case STATUS_IN_PROGRESS
            If progressText = "" Or progressRate <= 0 Or progressRate >= 1 Then
                progressRate = 0.5
            End If
        Case STATUS_ON_HOLD
            If progressText = "" Then
                progressRate = 0
            End If
        Case Else
            If progressRate <= 0 Then
                statusText = STATUS_NOT_STARTED
                progressRate = 0
            ElseIf progressRate >= 1 Then
                statusText = STATUS_COMPLETED
                progressRate = 1
            Else
                statusText = STATUS_IN_PROGRESS
            End If
    End Select

    ws.Cells(targetRow, COL_STATUS).Value = statusText
    ws.Cells(targetRow, COL_PROGRESS).Value = progressRate
End Sub

Public Sub NormalizeTaskStatusAndProgressRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim r As Long

    If ws Is Nothing Then Exit Sub
    If startRow < ROW_DATA_START Then startRow = ROW_DATA_START
    If endRow < startRow Then Exit Sub

    For r = startRow To endRow
        If HasTaskContentInRow(ws, r) Then
            SyncTaskStatusAndProgressRow ws, r
        End If
    Next r
End Sub

Public Sub ValidateDevelopmentHoursInput(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo ErrorHandler

    Dim normalizedHours As Double

    If Target.Row < ROW_DATA_START Then Exit Sub
    If Trim$(CStr(Target.Value)) = "" Then Exit Sub

    If Not TryParseDevelopmentHours(Target.Value, normalizedHours) Then
        MsgBox "開発LTは 3h / 3.5h / 3 の形式で入力してください。", vbExclamation, "入力エラー"
        Target.ClearContents
        Exit Sub
    End If

    Target.Value = normalizedHours
    Target.NumberFormat = DEV_HOURS_NUMBER_FORMAT
    Exit Sub

ErrorHandler:
    MsgBox "開発LTの検証中にエラーが発生しました: " & Err.Description, vbExclamation, "開発LTエラー"
End Sub

' ==========================================
'  初期セットアップ (ヘッダー作成＆書式設定)
' ==========================================
Sub SetupInazumaGantt(Optional ByVal silentMode As Boolean = False, Optional ByVal overrideStartDate As Variant = Null)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim hadExistingContent As Boolean
    hadExistingContent = (Application.WorksheetFunction.CountA(ws.UsedRange) > 0)

    If ws.Name <> MAIN_SHEET_NAME Then
        If hadExistingContent Then
            MsgBox "現在のシートには既存データがあります。" & vbCrLf & _
                   "新しい空シートでセットアップを実行してください。", vbExclamation, "セットアップ"
            Exit Sub
        End If

        On Error Resume Next
        ws.Name = MAIN_SHEET_NAME
        If Err.Number <> 0 Then
            MsgBox "シート名を '" & MAIN_SHEET_NAME & "' に変更できませんでした。" & vbCrLf & "既に同名のシートが存在する可能性があります。", vbExclamation
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    End If

    ' P2修正: 元の設定を保存
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

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
    ws.Range(COL_DEV_LT & ROW_HEADER).Value = "開発LT"
    ws.Range(COL_START_PLAN & ROW_HEADER).Value = "開始予定"
    ws.Range(COL_END_PLAN & ROW_HEADER).Value = "完了予定"
    ws.Range(COL_START_ACTUAL & ROW_HEADER).Value = "開始実績"
    ws.Range(COL_END_ACTUAL & ROW_HEADER).Value = "完了実績"

    ' ヘッダー行のスタイル（8行目、A～O列）
    With ws.Range("A" & ROW_HEADER & ":" & COL_END_ACTUAL & ROW_HEADER)
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
    ws.Columns("F").ColumnWidth = 22    ' TASK Lv4
    ws.Columns("G").ColumnWidth = 22    ' タスク補足
    ws.Columns("H").ColumnWidth = 7     ' 状況
    ws.Columns("I").ColumnWidth = 7     ' 進捗率
    ws.Columns("J").ColumnWidth = 7     ' 担当
    ws.Columns("K").ColumnWidth = 8     ' 開発LT
    ws.Columns("L").ColumnWidth = 8.7   ' 開始予定
    ws.Columns("M").ColumnWidth = 8.7   ' 完了予定
    ws.Columns("N").ColumnWidth = 8.7   ' 開始実績
    ws.Columns("O").ColumnWidth = 8.7   ' 完了実績

    ' 行高さ統一（22）
    ws.Rows.RowHeight = 22


    EnsureGuideSheet

    ' 説明シート作成後、メインシートに戻る
    ws.Activate

    ' 日付開始日を入力させる（キャンセル時はロールバック）
    Dim startDateInput As Variant
    If silentMode And Not IsNull(overrideStartDate) Then
        startDateInput = overrideStartDate
    ElseIf silentMode Then
        startDateInput = Format(Date, "yy/mm/dd")
    Else
        startDateInput = Application.InputBox("ガントチャートの開始日を入力してください (例: 24/12/25)", "開始日設定", Format(Date, "yy/mm/dd"), Type:=2)
    End If

    ' キャンセル処理（ロールバック）
    If Not silentMode And (startDateInput = False Or VarType(startDateInput) = vbBoolean) Then
        If Not hadExistingContent Then
            ' 新規セットアップ開始時のみロールバックを許可
            Dim rollbackEndCol As Long
            rollbackEndCol = ws.Columns(COL_GANTT_START).Column + GANTT_DAYS - 1
            Dim rollbackEndRow As Long
            rollbackEndRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
            ws.Range(ws.Cells(1, 1), ws.Cells(rollbackEndRow, rollbackEndCol)).Clear
        End If
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreenUpdating
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
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

    Dim todayDate As Date
    todayDate = Date
    If IsDate(ws.Range(CELL_TODAY).Value) Then
        todayDate = CDate(ws.Range(CELL_TODAY).Value)
    End If

    ' 週・日付・曜日ヘッダーの作成（統合関数呼び出し）
    RegenerateDateHeaders ws, ganttStartDate

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

    ' フィルタ自動設定 (7行目（日付行）A-O列)
    If Not ws.AutoFilterMode Then
        ws.Range("A" & ROW_DATE_HEADER & ":" & COL_END_ACTUAL & ROW_DATE_HEADER).AutoFilter
    End If

    ' Initial numbering for No.1 to No.1000
    Dim noRow As Long
    For noRow = ROW_DATA_START To ROW_DATA_START + DATA_ROWS_DEFAULT - 1
        ws.Cells(noRow, COL_NO).Value = noRow - ROW_DATA_START + 1
    Next noRow

    ' コントロールボタンの作成
    CreateControlButtons ws, True

    Application.Calculation = prevCalc  ' P2修正: 元設定に復元
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating

    If Application.DisplayAlerts Then
        MsgBox "セットアップ完了！" & vbCrLf & "データを入力後、RefreshInazumaGantt を実行してください。", vbInformation, "イナズマガント"
    End If
    Exit Sub

ErrorHandler:
    Application.Calculation = prevCalc  ' P2修正: 元設定に復元
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub


' ==========================================
'  入力規則と日付書式の適用
' ==========================================
Private Sub ApplyDataValidationAndFormats(ByVal ws As Worksheet, ByVal lastRow As Long)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START

    ' 開発LTは数値で保持し、表示だけ h 付きにする
    With ws.Range(COL_DEV_LT & ROW_DATA_START & ":" & COL_DEV_LT & lastRow)
        .NumberFormat = DEV_HOURS_NUMBER_FORMAT
        .HorizontalAlignment = xlCenter
    End With

    ' 進捗率のドロップダウン
    With ws.Range(COL_PROGRESS & ROW_DATA_START & ":" & COL_PROGRESS & lastRow)
        .NumberFormat = "0%"
        With .Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="100"
            .ShowInput = False
            .ErrorTitle = "進捗率の入力エラー"
            .ErrorMessage = "0 から 100 までの数値、または 0% から 100% の割合で入力してください。"
            .InCellDropdown = False
        End With
    End With

    ' 状況のドロップダウン
    With ws.Range(COL_STATUS & ROW_DATA_START & ":" & COL_STATUS & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="未着手,進行中,完了,保留"
        .ShowInput = False
        .ErrorTitle = "状況の入力エラー"
        .ErrorMessage = "未着手、進行中、完了、保留のいずれかを入力してください。"
        .InCellDropdown = True
    End With

    ' 日付列の書式
    ws.Range(COL_START_PLAN & ROW_DATA_START & ":" & COL_END_ACTUAL & lastRow).NumberFormat = "yy/mm/dd"
End Sub

' ==========================================
'  データ最終行の取得
' ==========================================
Public Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ROW_HEADER

    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_TASK).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row) ' Lv2
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row) ' Lv3
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "F").End(xlUp).Row) ' Lv4
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_TASK_DETAIL).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_STATUS).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_PROGRESS).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_ASSIGNEE).End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, COL_DEV_LT).End(xlUp).Row)
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

Private Sub SetGanttRefreshStatus(ByVal phaseText As String)
    Application.StatusBar = phaseText
    DoEvents
End Sub

Public Sub MaybeYieldDuringGanttRefresh(ByVal currentIndex As Long, ByVal totalCount As Long, Optional ByVal phaseText As String = "ガント更新")
    If currentIndex <= 0 Then Exit Sub
    If totalCount <= 0 Then totalCount = currentIndex
    If (currentIndex Mod GANTT_DOEVENTS_INTERVAL) <> 0 And currentIndex < totalCount Then Exit Sub

    If Not mIsRefreshingGantt Then
        DoEvents
        Exit Sub
    End If

    Application.StatusBar = phaseText & " " & Format$(currentIndex / totalCount, "0%")
    DoEvents
End Sub

' ==========================================
'  説明シートの作成
' ==========================================
Private Sub EnsureGuideSheet()
    On Error GoTo ErrorHandler

    Dim prevAlerts As Boolean
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Dim wsGuide As Worksheet
    On Error Resume Next
    Set wsGuide = ThisWorkbook.Worksheets(GUIDE_SHEET_NAME)
    On Error GoTo ErrorHandler

    If wsGuide Is Nothing Then
        Set wsGuide = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsGuide.Name = GUIDE_SHEET_NAME
    Else
        wsGuide.Cells.Clear
    End If

    wsGuide.Activate
    ActiveWindow.DisplayGridlines = False

    ' コンテンツを配列に準備（一括書き込みで安定性を向上）
    Dim content(1 To 30, 1 To 2) As Variant

    ' タイトル
    content(1, 1) = "マクロ機能"

    ' ボタン機能
    content(3, 1) = "■ ボタン機能"
    content(4, 1) = "【ガント更新】": content(4, 2) = "親集計、警告表示、ヘッダー、罫線、土日祝色、ガントをまとめて最新化します。"
    content(5, 2) = "高速入力後の反映や、表示が崩れた場合の再描画にも使います。"
    content(6, 1) = "【土日切替】": content(6, 2) = "土日列の表示/非表示を切替えます。"
    content(7, 2) = "画面を広く使いたい時に便利です。"
    content(8, 1) = "※ 旧書式リセット": content(8, 2) = "ボタンは廃止し、ガント更新に統合しました。"
    content(9, 2) = "旧 ResetFormatting マクロは互換用に残し、同じ更新処理を実行します。"

    ' ダブルクリック完了
    content(11, 1) = "■ ダブルクリックでタスク完了"
    content(12, 1) = "No.列(B列) または状況列(H列) をダブルクリックすると、そのタスクが完了になります。"
    content(13, 1) = ""
    content(14, 1) = "  ・ 状況 → 「完了」"
    content(15, 1) = "  ・ 進捗率 → 100%"
    content(16, 1) = "  ・ 完了実績 → 今日の日付（設定マスタで「自動」時）"
    content(17, 1) = ""
    content(18, 1) = "※ すでに完了しているタスクは変更されません。"

    ' SHIFT+右クリック機能
    content(20, 1) = "■ SHIFT+右クリックで折りたたみ"
    content(21, 1) = "LV1タスク（C列）でSHIFT+右クリックすると、"
    content(22, 1) = "配下のLV2-4タスクを折りたたみ/展開します。"
    content(23, 1) = ""
    content(24, 1) = "  ・ 再度SHIFT+右クリックで展開"
    content(25, 1) = "  ・ LV1タスク（大項目）のみ対象です"
    content(26, 1) = "  ・ 開発LT列(K列)では配下タスクの合計工数を集計"

    ' 一括書き込み
    On Error Resume Next
    wsGuide.Range("A1").Resize(30, 2).Value = content
    If Err.Number <> 0 Then
        MsgBox "配列書き込みエラー: " & Err.Description & vbCrLf & "Err.Number: " & Err.Number, vbCritical
        Err.Clear
    End If
    On Error GoTo ErrorHandler

    ' 書式設定
    With wsGuide
        .Range("A1").Font.Size = 14
        .Range("A1,A3,A11,A20").Font.Bold = True
        ' インデントや列幅
        .Columns(1).ColumnWidth = 35
        .Columns(2).ColumnWidth = 55
    End With

    Application.DisplayAlerts = prevAlerts
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = prevAlerts
    MsgBox "EnsureGuideSheet Error: " & Err.Description, vbCritical
End Sub

' ==========================================
'  ガント全体の罫線（罫線サマリに基づく詳細パターン適用）
' ==========================================
Private Sub ApplyGanttBorders(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim ganttStartCol As Long
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

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
    ApplyBorder ws.Range(COL_END_ACTUAL & "6"), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic

    ' --- P5: 7行目 (日付行) ---
    ' 7行目の背景色をヘッダーと同じ色で塗りつぶし
    ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)).Interior.Color = COLOR_HEADER_BG
    ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)).Font.Color = RGB(255, 255, 255)

    ApplyBorder ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)), xlEdgeTop, xlContinuous, xlMedium, xlColorIndexAutomatic
    ' 7行目下部に黒色の太線
    ApplyBorder ws.Range(ws.Cells(7, 1), ws.Cells(7, ganttEndCol)), xlEdgeBottom, xlContinuous, xlMedium, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(7, ws.Columns(COL_END_ACTUAL).Column), ws.Cells(7, ganttEndCol)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
    ApplyBorder ws.Range("A7"), xlEdgeLeft, xlContinuous, xlMedium, xlColorIndexAutomatic
    ApplyBorder ws.Range(ws.Cells(7, ganttStartCol), ws.Cells(7, ganttEndCol)), xlEdgeLeft, xlContinuous, xlThin, xlColorIndexAutomatic

    ' 7行目のP列以降のガントチャート部は太字
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

        ' ガントチャート部(P列以降)にもC-D間と同じ縦罫線
        ApplyBorderWithColorIndex ws.Range(ws.Cells(ROW_DATE_HEADER, ganttStartCol), ws.Cells(lastRow, ganttEndCol)), xlInsideVertical, xlContinuous, xlHairline, 15

        ' A-B, F-O列: 細線 自動
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, 2)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ' A列B列間は黒細線
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 1), ws.Cells(lastRow, 1)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 6), ws.Cells(lastRow, ws.Columns(COL_END_ACTUAL).Column)), xlEdgeRight, xlContinuous, xlThin, xlColorIndexAutomatic
        ApplyBorder ws.Range(ws.Cells(ROW_DATA_START, 6), ws.Cells(lastRow, ws.Columns(COL_END_ACTUAL).Column)), xlInsideVertical, xlContinuous, xlThin, xlColorIndexAutomatic
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
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

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
Sub DrawGanttBars(Optional ByVal skipRuntimeStateRepair As Boolean = False)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("ガント描画")
    If ws Is Nothing Then Exit Sub
    If Not skipRuntimeStateRepair Then
        Call RepairBulkEditRuntimeState(ws)
    End If

    ' P2修正: 元の設定を保存
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START

    Dim ganttStartCol As Long
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        ganttStartDate = Date
    End If

    Dim touchedShapes As Object
    Set touchedShapes = CreateObject("Scripting.Dictionary")

    ' 各行のバーを描画
    Dim r As Long
    Dim shp As Shape
    Dim startPlan As Variant, endPlan As Variant
    Dim startActual As Variant, endActual As Variant
    Dim progress As Double
    Dim startCol As Long, endCol As Long, progressCol As Long
    Dim cellTop As Double, cellLeft As Double, cellWidth As Double, cellHeight As Double
    Dim planBarHeight As Double
    Dim actualBarHeight As Double

    planBarHeight = 6
    actualBarHeight = 6

    Dim inazumaPoints() As Variant
    ReDim inazumaPoints(1 To lastRow - ROW_DATA_START + 1, 1 To 2)
    Dim inazumaCount As Long
    inazumaCount = 0

    For r = ROW_DATA_START To lastRow
        MaybeYieldDuringGanttRefresh r - ROW_DATA_START + 1, lastRow - ROW_DATA_START + 1, "ガントバー描画"
        ' 日付を取得
        startPlan = ws.Cells(r, COL_START_PLAN).Value
        endPlan = ws.Cells(r, COL_END_PLAN).Value
        startActual = ws.Cells(r, COL_START_ACTUAL).Value
        endActual = ws.Cells(r, COL_END_ACTUAL).Value

        ' 進捗率を取得
        progress = NormalizeProgressValue(ws.Cells(r, COL_PROGRESS).Value, 0)

        ' 予定バーを描画
        If IsDate(startPlan) And IsDate(endPlan) Then
            startCol = DateToColumn(ganttStartDate, CDate(startPlan), ganttStartCol)
            endCol = DateToColumn(ganttStartDate, CDate(endPlan), ganttStartCol)

            ' P1修正: 開始が範囲外でも終了が範囲内ならクランプして描画
            If startCol < ganttStartCol Then startCol = ganttStartCol
            If endCol > ganttStartCol + GANTT_DAYS - 1 Then endCol = ganttStartCol + GANTT_DAYS - 1

            If startCol <= ganttStartCol + GANTT_DAYS - 1 And endCol >= ganttStartCol Then
                If endCol >= startCol Then
                    cellTop = GetStackedBarTop(ws, r, startCol, planBarHeight, actualBarHeight, False)
                    cellLeft = ws.Cells(r, startCol).Left
                    cellWidth = ws.Cells(r, endCol).Left + ws.Cells(r, endCol).Width - cellLeft

                    ' 予定バー（薄い灰色 + 黒枠線）
                    Set shp = UpsertGanttRectangle(ws, touchedShapes, "Bar_Plan_" & r, _
                                                    cellLeft, cellTop, cellWidth, planBarHeight, _
                                                    COLOR_PLAN, True, RGB(0, 0, 0), 1)

                    ' 進捗バー（紺色 + 黒枠線）
                    If progress > 0 Then
                        progressCol = startCol + CLng((endCol - startCol + 1) * progress) - 1
                        If progressCol < startCol Then progressCol = startCol
                        If progressCol >= startCol Then
                            Dim progressWidth As Double
                            progressWidth = ws.Cells(r, progressCol).Left + ws.Cells(r, progressCol).Width - cellLeft
                            If progressWidth < ws.Cells(r, startCol).Width Then progressWidth = ws.Cells(r, startCol).Width
                            If progress >= 1 Then progressWidth = cellWidth

                            Set shp = UpsertGanttRectangle(ws, touchedShapes, "Bar_Progress_" & r, _
                                                            cellLeft, cellTop, progressWidth, planBarHeight, _
                                                            COLOR_PROGRESS, True, RGB(0, 0, 0), 1)
                        End If
                    End If

                    ' イナズマ線用のポイントを記録（全可視タスクを対象）
                    ' 将来タスクも開始位置で結び、行の途中で線が途切れないようにする
                    Dim inazumaX As Double
                    inazumaX = 0
                    Dim todayDate As Date
                    todayDate = Date

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

                        ' 未着手かつ開始予定日が未来の場合は、今日線と同じ位置からオレンジ線を開始する
                        If progress = 0 And CDate(startPlan) > Date And todayX > 0 Then
                            useTodayPosition = True
                        End If
                    End If

                    ' 今日の位置を使用する場合
                    If useTodayPosition And todayX > 0 Then
                        inazumaX = todayX
                    End If

                    inazumaX = ClampGanttXPosition(ws, r, ganttStartCol, inazumaX)
                    inazumaCount = inazumaCount + 1
                    inazumaPoints(inazumaCount, 1) = inazumaX
                    inazumaPoints(inazumaCount, 2) = cellTop + planBarHeight / 2
                End If
            End If
        End If

        ' 実績バー（緑色の塗りつぶしバー、予定の下に配置）
        If IsDate(startActual) And IsDate(startPlan) And IsDate(endPlan) Then
            ' 実績バーの右端は進捗バーの右端と揃える
            Dim actualStartCol As Long
            Dim planStartCol As Long
            Dim planEndCol As Long
            Dim visibleStartCol As Long
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

            visibleStartCol = actualStartCol
            If visibleStartCol < ganttStartCol Then visibleStartCol = ganttStartCol

            ' 緑バーは予定終了日まで。表示範囲外の開始日でも途中から描画する
            Dim greenEndCol As Long
            greenEndCol = progressEndCol
            If greenEndCol > planEndCol Then greenEndCol = planEndCol
            If greenEndCol > ganttStartCol + GANTT_DAYS - 1 Then greenEndCol = ganttStartCol + GANTT_DAYS - 1

            If greenEndCol >= ganttStartCol And visibleStartCol <= ganttStartCol + GANTT_DAYS - 1 And greenEndCol >= visibleStartCol Then
                cellTop = GetStackedBarTop(ws, r, visibleStartCol, planBarHeight, actualBarHeight, True)
                cellLeft = ws.Cells(r, visibleStartCol).Left
                cellWidth = ws.Cells(r, greenEndCol).Left + ws.Cells(r, greenEndCol).Width - cellLeft

                Set shp = UpsertGanttRectangle(ws, touchedShapes, "Bar_Actual_" & r, _
                                                cellLeft, cellTop, cellWidth, actualBarHeight, _
                                                COLOR_ACTUAL, False, 0, 0)
            End If

            ' 完了時に実績で超過している場合は超過部分を別色で描画
            If progress >= 1 And IsDate(endActual) Then
                Dim actualEndDate As Date
                actualEndDate = CDate(endActual)
                If actualEndDate > CDate(endPlan) Then
                    Dim overrunStartCol As Long
                    Dim overrunEndCol As Long
                    overrunStartCol = planEndCol + 1
                    If overrunStartCol < ganttStartCol Then overrunStartCol = ganttStartCol
                    overrunEndCol = DateToColumn(ganttStartDate, actualEndDate, ganttStartCol)
                    If overrunEndCol > ganttStartCol + GANTT_DAYS - 1 Then overrunEndCol = ganttStartCol + GANTT_DAYS - 1

                    If overrunStartCol <= ganttStartCol + GANTT_DAYS - 1 And overrunEndCol >= overrunStartCol Then
                        Dim overrunLeft As Double
                        Dim overrunWidth As Double
                        cellTop = GetStackedBarTop(ws, r, overrunStartCol, planBarHeight, actualBarHeight, True)
                        overrunLeft = ws.Cells(r, overrunStartCol).Left
                        overrunWidth = ws.Cells(r, overrunEndCol).Left + ws.Cells(r, overrunEndCol).Width - overrunLeft

                        Set shp = UpsertGanttRectangle(ws, touchedShapes, "Bar_Overrun_" & r, _
                                                        overrunLeft, cellTop, overrunWidth, actualBarHeight, _
                                                        COLOR_ACTUAL_OVERRUN, False, 0, 0)
                    End If
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

        Set shp = UpsertGanttLine(ws, touchedShapes, "Today_Line", _
                                  todayLeft, todayTop, todayLeft, todayBottom, _
                                  COLOR_TODAY, TODAY_LINE_WEIGHT)
    End If

    ' イナズマ線を描画（複数ポイントがある場合）
    If inazumaCount >= 2 Then
        DeleteShapeIfExists ws, "Inazuma_Line"

        Dim p As Long
        Dim segmentName As String
        For p = 2 To inazumaCount
            segmentName = "Inazuma_Line_" & CStr(p - 1)
            DeleteShapeIfExists ws, segmentName
            Set shp = ws.Shapes.AddLine(inazumaPoints(p - 1, 1), inazumaPoints(p - 1, 2), _
                                        inazumaPoints(p, 1), inazumaPoints(p, 2))
            shp.Name = segmentName
            shp.Line.ForeColor.RGB = COLOR_INAZUMA
            shp.Line.Weight = INAZUMA_LINE_WEIGHT
            shp.Placement = xlMoveAndSize
            shp.ZOrder msoBringToFront
            touchedShapes(shp.Name) = True
        Next p
    End If

    Dim deletedShapeCount As Long
    deletedShapeCount = DeleteUntouchedGanttShapes(ws, touchedShapes)
    LogGanttRefresh ws, ROW_DATA_START, lastRow, touchedShapes.Count, deletedShapeCount
    Application.Calculation = prevCalc  ' P2修正: 元設定に復元
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub

ErrorHandler:
    Application.Calculation = prevCalc  ' P2修正: 元設定に復元
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
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

Private Function ClampGanttXPosition(ByVal ws As Worksheet, ByVal targetRow As Long, _
                                     ByVal ganttStartCol As Long, ByVal rawX As Double) As Double
    Dim leftBound As Double
    Dim rightBound As Double
    Dim edgePadding As Double

    edgePadding = INAZUMA_LINE_WEIGHT + 48
    leftBound = ws.Cells(targetRow, ganttStartCol).Left + edgePadding
    rightBound = ws.Cells(targetRow, ganttStartCol + GANTT_DAYS - 1).Left + _
                 ws.Cells(targetRow, ganttStartCol + GANTT_DAYS - 1).Width - edgePadding

    ClampGanttXPosition = rawX
    If ClampGanttXPosition < leftBound Then ClampGanttXPosition = leftBound
    If ClampGanttXPosition > rightBound Then ClampGanttXPosition = rightBound
End Function

' ==========================================
'  全描画実行
' ==========================================
Sub RefreshInazumaGantt()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("ガント更新")
    If ws Is Nothing Then Exit Sub
    If mIsRefreshingGantt Then
        Application.StatusBar = "ガント更新はすでに実行中です。"
        DoEvents
        Exit Sub
    End If

    ' P2修正: 元の設定を保存
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevScreenUpdating = Application.ScreenUpdating

    mIsRefreshingGantt = True
    Call RepairBulkEditRuntimeState(ws)
    CancelDeferredBulkEditReconcile

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    SetGanttRefreshStatus "ガント更新を開始しています..."
    ReconcileDeferredTaskState ws

    Application.Calculation = prevCalc  ' P2修正: 元設定に復元
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating

    Application.StatusBar = False
    mIsRefreshingGantt = False
    UpdateBulkEditModeIndicator ws, "最新状態へ更新しました。"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.Calculation = prevCalc  ' P2修正: 元設定に復元
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
    mIsRefreshingGantt = False
    MsgBox "更新中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

Private Sub DeleteExistingGanttShapes(ByVal ws As Worksheet)
    Dim shapeIndex As Long

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        With ws.Shapes(shapeIndex)
            If Left(.Name, 4) = "Bar_" Or Left(.Name, 6) = "Today_" Or Left(.Name, 8) = "Inazuma_" Then
                .Delete
            End If
        End With
    Next shapeIndex
End Sub

Private Function IsManagedGanttShapeName(ByVal shapeName As String) As Boolean
    IsManagedGanttShapeName = (Left$(shapeName, 4) = "Bar_" Or _
                               Left$(shapeName, 6) = "Today_" Or _
                               Left$(shapeName, 8) = "Inazuma_")
End Function

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String)
    On Error Resume Next
    ws.Shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Function GetShapeIfExists(ByVal ws As Worksheet, ByVal shapeName As String) As Shape
    On Error Resume Next
    Set GetShapeIfExists = ws.Shapes(shapeName)
    On Error GoTo 0
End Function

Private Function UpsertGanttRectangle(ByVal ws As Worksheet, ByVal touchedShapes As Object, _
                                      ByVal shapeName As String, ByVal leftPos As Double, _
                                      ByVal topPos As Double, ByVal shapeWidth As Double, _
                                      ByVal shapeHeight As Double, ByVal fillColor As Long, _
                                      ByVal showLine As Boolean, ByVal lineColor As Long, _
                                      ByVal lineWeight As Double) As Shape
    Dim shp As Shape

    Set shp = GetShapeIfExists(ws, shapeName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddShape(msoShapeRectangle, leftPos, topPos, shapeWidth, shapeHeight)
        shp.Name = shapeName
    Else
        With shp
            .Left = leftPos
            .Top = topPos
            .Width = shapeWidth
            .Height = shapeHeight
        End With
    End If

    With shp
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = IIf(showLine, msoTrue, msoFalse)
        If showLine Then
            .Line.ForeColor.RGB = lineColor
            .Line.Weight = lineWeight
        End If
        .Placement = xlMoveAndSize
    End With

    touchedShapes(shapeName) = True
    Set UpsertGanttRectangle = shp
End Function

Private Function UpsertGanttLine(ByVal ws As Worksheet, ByVal touchedShapes As Object, _
                                 ByVal shapeName As String, ByVal beginX As Double, _
                                 ByVal beginY As Double, ByVal endX As Double, _
                                 ByVal endY As Double, ByVal lineColor As Long, _
                                 ByVal lineWeight As Double) As Shape
    Dim shp As Shape

    Set shp = GetShapeIfExists(ws, shapeName)
    If shp Is Nothing Then
        Set shp = ws.Shapes.AddLine(beginX, beginY, endX, endY)
        shp.Name = shapeName
    Else
        With shp
            .Left = beginX
            .Top = beginY
            .Width = endX - beginX
            .Height = endY - beginY
        End With
    End If

    With shp
        .Line.ForeColor.RGB = lineColor
        .Line.Weight = lineWeight
        .Placement = xlMoveAndSize
        .ZOrder msoBringToFront
    End With

    touchedShapes(shapeName) = True
    Set UpsertGanttLine = shp
End Function

Private Function DeleteUntouchedGanttShapes(ByVal ws As Worksheet, ByVal touchedShapes As Object) As Long
    Dim shapeIndex As Long
    Dim shapeName As String
    Dim deletedCount As Long

    For shapeIndex = ws.Shapes.Count To 1 Step -1
        shapeName = ws.Shapes(shapeIndex).Name
        If IsManagedGanttShapeName(shapeName) Then
            If touchedShapes Is Nothing Or Not touchedShapes.Exists(shapeName) Then
                ws.Shapes(shapeIndex).Delete
                deletedCount = deletedCount + 1
            End If
        End If
    Next shapeIndex
    DeleteUntouchedGanttShapes = deletedCount
End Function

Private Function GetStackedBarTop(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal targetCol As Long, _
                                  ByVal upperBarHeight As Double, ByVal lowerBarHeight As Double, _
                                  ByVal lowerBar As Boolean) As Double
    Dim baseTop As Double
    Dim cellHeight As Double
    Dim gapHeight As Double
    Dim totalBarHeight As Double

    cellHeight = ws.Cells(targetRow, targetCol).Height
    gapHeight = 2
    totalBarHeight = upperBarHeight + lowerBarHeight + gapHeight
    baseTop = ws.Cells(targetRow, targetCol).Top + (cellHeight - totalBarHeight) / 2

    If lowerBar Then
        GetStackedBarTop = baseTop + upperBarHeight + gapHeight
    Else
        GetStackedBarTop = baseTop
    End If
End Function


' ==========================================
'  祝日列の色塗り（設定マスタ A16 以降）
' ==========================================
Private Sub ApplyHolidayColors(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim wsSettings As Worksheet
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    If wsSettings Is Nothing Then Exit Sub

    Dim ganttStartCol As Long
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        Exit Sub
    End If

    Dim lastHolidayRow As Long
    lastHolidayRow = wsSettings.Cells(wsSettings.Rows.Count, "A").End(xlUp).Row
    If lastHolidayRow < HOLIDAY_DATA_START_ROW Then Exit Sub

    Dim r As Long
    Dim holidayDate As Date
    Dim colIndex As Long

    For r = HOLIDAY_DATA_START_ROW To lastHolidayRow
        If IsDate(wsSettings.Cells(r, "A").Value) Then
            holidayDate = CDate(wsSettings.Cells(r, "A").Value)
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

Public Function IsAlertMarkerText(ByVal textValue As String) As Boolean
    textValue = Trim$(textValue)
    IsAlertMarkerText = (textValue = WBSParentRollup.ALERT_MARK_TODAY Or textValue = WBSParentRollup.ALERT_MARK_DELAY)
End Function

Public Function ExtractAlertMarkerText(ByVal textValue As String) As String
    textValue = Trim$(textValue)

    If textValue = "" Then Exit Function

    If textValue = WBSParentRollup.ALERT_MARK_DELAY Or _
       Left$(textValue, Len(WBSParentRollup.ALERT_MARK_DELAY & " ")) = WBSParentRollup.ALERT_MARK_DELAY & " " Then
        ExtractAlertMarkerText = WBSParentRollup.ALERT_MARK_DELAY
    ElseIf textValue = WBSParentRollup.ALERT_MARK_TODAY Or _
           Left$(textValue, Len(WBSParentRollup.ALERT_MARK_TODAY & " ")) = WBSParentRollup.ALERT_MARK_TODAY & " " Then
        ExtractAlertMarkerText = WBSParentRollup.ALERT_MARK_TODAY
    End If
End Function

Public Function BuildAuxiliaryDisplayText(Optional ByVal markerText As String = "") As String
    markerText = Trim$(markerText)
    If markerText <> "" Then
        BuildAuxiliaryDisplayText = markerText & " " & AUXILIARY_TASK_PLACEHOLDER
    Else
        BuildAuxiliaryDisplayText = AUXILIARY_TASK_PLACEHOLDER
    End If
End Function

Public Function IsAuxiliaryPlaceholderText(ByVal textValue As String) As Boolean
    textValue = Trim$(textValue)
    IsAuxiliaryPlaceholderText = (textValue = AUXILIARY_TASK_PLACEHOLDER Or _
                                  textValue = BuildAuxiliaryDisplayText(WBSParentRollup.ALERT_MARK_TODAY) Or _
                                  textValue = BuildAuxiliaryDisplayText(WBSParentRollup.ALERT_MARK_DELAY))
End Function

Private Function IsDisplayOnlyTaskText(ByVal textValue As String) As Boolean
    IsDisplayOnlyTaskText = (IsAlertMarkerText(textValue) Or IsAuxiliaryPlaceholderText(textValue))
End Function

Public Function GetVisibleTaskLabelForRow(ByVal ws As Worksheet, ByVal targetRow As Long) As String
    Dim cellText As String

    If ws Is Nothing Then Exit Function
    If targetRow < ROW_DATA_START Then Exit Function

    cellText = Trim$(CStr(ws.Cells(targetRow, "F").Value))
    If cellText <> "" Then
        GetVisibleTaskLabelForRow = cellText
        Exit Function
    End If

    cellText = Trim$(CStr(ws.Cells(targetRow, "E").Value))
    If cellText <> "" Then
        GetVisibleTaskLabelForRow = cellText
        Exit Function
    End If

    cellText = Trim$(CStr(ws.Cells(targetRow, "D").Value))
    If cellText <> "" Then
        GetVisibleTaskLabelForRow = cellText
        Exit Function
    End If

    cellText = Trim$(CStr(ws.Cells(targetRow, "C").Value))
    If cellText <> "" And Not IsDisplayOnlyTaskText(cellText) Then
        GetVisibleTaskLabelForRow = cellText
    ElseIf HasTaskContentInRow(ws, targetRow) Then
        GetVisibleTaskLabelForRow = AUXILIARY_TASK_PLACEHOLDER
    End If
End Function

Private Function GetRetainedTaskLevel(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    Dim r As Long

    If ws Is Nothing Then Exit Function
    If targetRow < ROW_DATA_START Then Exit Function

    If IsNumeric(ws.Cells(targetRow, COL_HIERARCHY).Value) Then
        GetRetainedTaskLevel = CLng(ws.Cells(targetRow, COL_HIERARCHY).Value)
        If GetRetainedTaskLevel > 0 Then Exit Function
    End If

    For r = targetRow - 1 To ROW_DATA_START Step -1
        If IsNumeric(ws.Cells(r, COL_HIERARCHY).Value) Then
            GetRetainedTaskLevel = CLng(ws.Cells(r, COL_HIERARCHY).Value)
            If GetRetainedTaskLevel > 0 Then Exit Function
        End If
    Next r
End Function

Private Sub ClearTaskLabelPresentation(ByVal taskCell As Range)
    If taskCell Is Nothing Then Exit Sub

    taskCell.Font.Italic = False
    taskCell.Font.Bold = False
    taskCell.Font.ColorIndex = xlColorIndexAutomatic
    taskCell.HorizontalAlignment = xlGeneral
End Sub

Public Sub RefreshTaskLabelPresentation(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim taskCell As Range
    Dim textValue As String
    Dim markerText As String

    If ws Is Nothing Then Exit Sub
    If targetRow < ROW_DATA_START Then Exit Sub

    Set taskCell = ws.Cells(targetRow, "C")
    textValue = Trim$(CStr(taskCell.Value))

    If IsAuxiliaryPlaceholderText(textValue) Then
        markerText = ExtractAlertMarkerText(textValue)
        taskCell.Font.Italic = True
        taskCell.Font.Bold = (markerText <> "")
        If markerText <> "" Then
            taskCell.Font.Color = WBSParentRollup.ALERT_COLOR_RED
        Else
            taskCell.Font.Color = RGB(127, 127, 127)
        End If
        taskCell.HorizontalAlignment = xlLeft
    ElseIf IsAlertMarkerText(textValue) Then
        taskCell.Font.Italic = False
        taskCell.Font.Bold = True
        taskCell.Font.Color = WBSParentRollup.ALERT_COLOR_RED
        taskCell.HorizontalAlignment = xlCenter
    ElseIf ExtractAlertMarkerText(textValue) <> "" Then
        taskCell.Font.Italic = False
        taskCell.Font.Bold = True
        taskCell.Font.Color = WBSParentRollup.ALERT_COLOR_RED
        taskCell.HorizontalAlignment = xlLeft
    Else
        ClearTaskLabelPresentation taskCell
    End If
End Sub

Public Sub RefreshTaskRowDisplayState(ByVal ws As Worksheet, ByVal targetRow As Long)
    Dim taskLevel As Long
    Dim retainedLevel As Long
    Dim taskCell As Range

    If ws Is Nothing Then Exit Sub
    If targetRow < ROW_DATA_START Then Exit Sub

    Set taskCell = ws.Cells(targetRow, "C")

    If HasPrimaryTaskContentInRow(ws, targetRow) Then
        taskLevel = GetTaskLevelForRow(ws, targetRow)
        If taskLevel > 0 Then
            ws.Cells(targetRow, COL_HIERARCHY).Value = taskLevel
            ws.Cells(targetRow, COL_HIERARCHY).NumberFormat = "General"
        End If

        If IsAuxiliaryPlaceholderText(CStr(taskCell.Value)) Then
            taskCell.ClearContents
        End If

        RefreshTaskLabelPresentation ws, targetRow
        Exit Sub
    End If

    If HasTaskContentInRow(ws, targetRow) Then
        retainedLevel = GetRetainedTaskLevel(ws, targetRow)
        If retainedLevel > 0 Then
            ws.Cells(targetRow, COL_HIERARCHY).Value = retainedLevel
            ws.Cells(targetRow, COL_HIERARCHY).NumberFormat = HIDDEN_VALUE_NUMBER_FORMAT
        Else
            ws.Cells(targetRow, COL_HIERARCHY).ClearContents
            ws.Cells(targetRow, COL_HIERARCHY).NumberFormat = "General"
        End If

        taskCell.Value = BuildAuxiliaryDisplayText()
        RefreshTaskLabelPresentation ws, targetRow
        Exit Sub
    End If

    ResetTaskRowDisplay ws, targetRow
End Sub

' ==========================================
'  タスク入力列から階層を自動判定
' ==========================================
Public Function HasTaskContentInRow(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRow < ROW_DATA_START Then Exit Function

    If HasPrimaryTaskContentInRow(ws, targetRow) Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_TASK_DETAIL).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_STATUS).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_PROGRESS).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_ASSIGNEE).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_DEV_LT).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_START_PLAN).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_END_PLAN).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_START_ACTUAL).Value)) <> "" Then
        HasTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, COL_END_ACTUAL).Value)) <> "" Then
        HasTaskContentInRow = True
    End If
End Function

Public Function HasPrimaryTaskContentInRow(ByVal ws As Worksheet, ByVal targetRow As Long) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRow < ROW_DATA_START Then Exit Function

    If Trim$(CStr(ws.Cells(targetRow, "F").Value)) <> "" Then
        HasPrimaryTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, "E").Value)) <> "" Then
        HasPrimaryTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, "D").Value)) <> "" Then
        HasPrimaryTaskContentInRow = True
    ElseIf Trim$(CStr(ws.Cells(targetRow, "C").Value)) <> "" And _
           Not IsDisplayOnlyTaskText(CStr(ws.Cells(targetRow, "C").Value)) Then
        HasPrimaryTaskContentInRow = True
    End If
End Function

Private Function GetTaskLevelForRow(ByVal ws As Worksheet, ByVal targetRow As Long) As Long
    If ws Is Nothing Then Exit Function
    If targetRow < ROW_DATA_START Then Exit Function

    If Trim$(CStr(ws.Cells(targetRow, "F").Value)) <> "" Then
        GetTaskLevelForRow = 4
    ElseIf Trim$(CStr(ws.Cells(targetRow, "E").Value)) <> "" Then
        GetTaskLevelForRow = 3
    ElseIf Trim$(CStr(ws.Cells(targetRow, "D").Value)) <> "" Then
        GetTaskLevelForRow = 2
    ElseIf Trim$(CStr(ws.Cells(targetRow, "C").Value)) <> "" And _
           Not IsDisplayOnlyTaskText(CStr(ws.Cells(targetRow, "C").Value)) Then
        GetTaskLevelForRow = 1
    End If
End Function

Public Sub AutoDetectTaskLevelsInRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long)
    Dim r As Long
    Dim taskLevel As Long
    Dim lastSupportedRow As Long

    If ws Is Nothing Then Exit Sub

    If startRow < ROW_DATA_START Then startRow = ROW_DATA_START
    lastSupportedRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    If endRow > lastSupportedRow Then endRow = lastSupportedRow
    If endRow < startRow Then Exit Sub

    For r = startRow To endRow
        taskLevel = GetTaskLevelForRow(ws, r)
        If taskLevel > 0 Then
            ws.Cells(r, COL_HIERARCHY).Value = taskLevel
            ws.Cells(r, COL_HIERARCHY).NumberFormat = "General"
        End If
        RefreshTaskRowDisplayState ws, r
    Next r
End Sub

Public Sub AutoDetectTaskLevel(Optional ByVal targetRow As Long = 0)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("階層自動判定")
    If ws Is Nothing Then Exit Sub

    Dim startRow As Long
    Dim endRow As Long

    If targetRow > 0 Then
        startRow = targetRow
        endRow = targetRow
    Else
        startRow = ROW_DATA_START
        endRow = GetLastDataRow(ws)
        If endRow < ROW_DATA_START Then endRow = ROW_DATA_START + DATA_ROWS_DEFAULT - 1
    End If

    AutoDetectTaskLevelsInRange ws, startRow, endRow
    Exit Sub

ErrorHandler:
    MsgBox "階層自動判定エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  コントロールボタンの作成
' ==========================================
Private Sub CreateControlButtons(ByVal ws As Worksheet, Optional ByVal skipRuntimeStateRepair As Boolean = False)
    On Error Resume Next
    Dim repairNote As String

    If ws Is Nothing Then Exit Sub

    If Not skipRuntimeStateRepair Then
        repairNote = RepairBulkEditRuntimeState(ws)
    End If

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

    ' 一括編集切替ボタン
    btnLeft = btnLeft + btnWidth + 10
    Dim btnBulkEdit As Shape
    Set btnBulkEdit = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth + 20, btnHeight)
    With btnBulkEdit
        .Name = "Btn_BulkEdit"
        If IsBulkEditModeEnabled() Then
            .Fill.ForeColor.RGB = RGB(192, 80, 77)
            .TextFrame2.TextRange.Characters.Text = "高速入力 ON"
        Else
            .Fill.ForeColor.RGB = RGB(84, 130, 53)
            .TextFrame2.TextRange.Characters.Text = "高速入力 OFF"
        End If
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "ToggleBulkEditMode"
    End With

    UpdateBulkEditModeIndicator ws, repairNote

    ' 日付シフトボタン (v3追加)
    btnLeft = btnLeft + btnWidth + 50
    Dim btnShift As Shape
    Set btnShift = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btnShift
        .Name = "Btn_ShiftDates"
        .Fill.ForeColor.RGB = RGB(0, 128, 128)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = "日付シフト"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "ShiftDates"
    End With

    ' PDF出力ボタン (v3追加)
    btnLeft = btnLeft + btnWidth + 10
    Dim btnPDF As Shape
    Set btnPDF = ws.Shapes.AddShape(msoShapeRoundedRectangle, btnLeft, btnTop, btnWidth, btnHeight)
    With btnPDF
        .Name = "Btn_ExportPDF"
        .Fill.ForeColor.RGB = RGB(192, 0, 0)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = "PDF出力"
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = "ExportToPDF"
    End With
End Sub

Public Sub ToggleBulkEditMode()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim targetEnabled As Boolean
    Dim affectedRows As Object

    If mIsTogglingBulkEdit Or mIsRefreshingGantt Then
        Application.StatusBar = "高速入力モードの切替またはガント更新がすでに実行中です。"
        DoEvents
        Exit Sub
    End If

    Set ws = RequireMainWorksheet("一括編集モード切替")
    If ws Is Nothing Then Exit Sub

    mIsTogglingBulkEdit = True
    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    targetEnabled = Not IsBulkEditModeEnabled()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    SetBulkEditMode targetEnabled
    CancelDeferredBulkEditReconcile
    CreateControlButtons ws, True

    If targetEnabled Then
        CaptureBulkEditSnapshot ws
        PrepareLiveTaskLevelHintsForBulkEdit ws
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreenUpdating
        Application.StatusBar = False
        UpdateBulkEditModeIndicator ws, "Ctrl+Z を優先しつつ、TASK入力時のLV表示だけ有効にしました。"
        mIsTogglingBulkEdit = False
        Exit Sub
    End If

    Set affectedRows = DetectBulkEditAffectedRows(ws)
    mIsRefreshingGantt = True
    ReconcileDeferredTaskState ws, affectedRows
    mIsRefreshingGantt = False
    ClearBulkEditSnapshot

    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    CreateControlButtons ws, True
    Application.StatusBar = False
    UpdateBulkEditModeIndicator ws, "再整合済み。"
    mIsTogglingBulkEdit = False
    Exit Sub

ErrorHandler:
    mIsRefreshingGantt = False
    mIsTogglingBulkEdit = False
    If targetEnabled Then
        SetBulkEditMode False
        Application.EnableEvents = True
    Else
        SetBulkEditMode True
        Application.EnableEvents = False
    End If
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    If Not ws Is Nothing Then CreateControlButtons ws, True
    MsgBox "高速入力モード切替エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  土日列の表示/非表示切替
' ==========================================
Sub ToggleWeekends()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("土日切替")
    If ws Is Nothing Then Exit Sub
    Call RepairBulkEditRuntimeState(ws)

    Dim ganttStartCol As Long
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

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
'  旧書式リセット互換
' ==========================================
Sub ResetFormatting()
    RefreshInazumaGantt
End Sub

Private Sub CreateWbsBackupSheetCore(Optional ByVal skipNotification As Boolean = False)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsBackupStaging As Worksheet
    Dim lastRow As Long
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim stateCaptured As Boolean
    Dim errDescription As String

    CaptureApplicationState prevCalc, prevEvents, prevScreenUpdating, stateCaptured
    Set ws = RequireMainWorksheet("WBS退避")
    If ws Is Nothing Then Exit Sub
    Call RepairBulkEditRuntimeState(ws)
    prevEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START

    Set wsBackupStaging = CreateTemporaryWorksheet(BACKUP_TEMP_SHEET_NAME)
    CopyWbsSnapshot ws, wsBackupStaging, lastRow, True
    WriteBackupMetadata wsBackupStaging, ws, lastRow
    ReplaceBackupWorksheetWithStaged wsBackupStaging

    RestoreApplicationState prevCalc, prevEvents, prevScreenUpdating, stateCaptured
    Application.StatusBar = False

    If (Not skipNotification) And Application.DisplayAlerts Then
        MsgBox "最新の WBS バックアップを '" & BACKUP_SHEET_NAME & "' に更新しました。", vbInformation, "WBS退避"
    End If
    Exit Sub

ErrorHandler:
    errDescription = Err.Description
    On Error Resume Next
    DeleteWorksheetIfExists BACKUP_TEMP_SHEET_NAME, MAIN_SHEET_NAME
    On Error GoTo 0
    RestoreApplicationState prevCalc, prevEvents, prevScreenUpdating, stateCaptured
    MsgBox "WBS退避エラー: " & errDescription, vbCritical, "エラー"
End Sub

Public Sub CreateWbsBackupSheet()
    CreateWbsBackupSheetCore False
End Sub

Public Sub CreateWbsBackupSheetSilent()
    CreateWbsBackupSheetCore True
End Sub

Private Sub RestoreWbsFromBackupSheetCore(Optional ByVal skipConfirmation As Boolean = False)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim wsBackup As Worksheet
    Dim wsRollback As Worksheet
    Dim backupLastRow As Long
    Dim clearEndRow As Long
    Dim rollbackLastRow As Long
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean
    Dim prevScreenUpdating As Boolean
    Dim stateCaptured As Boolean
    Dim restoreStarted As Boolean
    Dim errDescription As String

    CaptureApplicationState prevCalc, prevEvents, prevScreenUpdating, stateCaptured
    Set ws = RequireMainWorksheet("バックアップ復元")
    If ws Is Nothing Then Exit Sub
    Call RepairBulkEditRuntimeState(ws)

    Set wsBackup = GetBackupWorksheet()
    If wsBackup Is Nothing Then
        MsgBox "バックアップシート '" & BACKUP_SHEET_NAME & "' が見つかりません。", vbExclamation, "バックアップ復元"
        Exit Sub
    End If

    backupLastRow = GetLastDataRow(wsBackup)
    If backupLastRow < ROW_DATA_START Then backupLastRow = ROW_DATA_START

    If Not skipConfirmation Then
        If MsgBox(BuildRestoreConfirmationMessage(wsBackup, backupLastRow), _
                  vbYesNo + vbQuestion + vbDefaultButton2, "バックアップ復元") <> vbYes Then
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    SetBulkEditMode False

    clearEndRow = GetSnapshotClearEndRow(ws, backupLastRow)
    rollbackLastRow = clearEndRow
    Set wsRollback = CreateTemporaryWorksheet(RESTORE_ROLLBACK_SHEET_NAME)
    CopyWbsSnapshot ws, wsRollback, rollbackLastRow, True
    restoreStarted = True

    ClearMainWorksheetForRestore ws, clearEndRow
    CopyWbsSnapshot wsBackup, ws, backupLastRow
    ReconcileDeferredTaskState ws
    DeleteWorksheetIfExists RESTORE_ROLLBACK_SHEET_NAME, MAIN_SHEET_NAME

    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScreenUpdating
    If Not ws Is Nothing Then CreateControlButtons ws, True
    Application.StatusBar = False

    If (Not skipConfirmation) And Application.DisplayAlerts Then
        MsgBox "バックアップシートから WBS を復元しました。", vbInformation, "バックアップ復元"
    End If
    Exit Sub

ErrorHandler:
    errDescription = Err.Description
    On Error Resume Next
    If restoreStarted And Not wsRollback Is Nothing And Not ws Is Nothing Then
        ClearMainWorksheetForRestore ws, rollbackLastRow
        CopyWbsSnapshot wsRollback, ws, rollbackLastRow
    End If
    DeleteWorksheetIfExists RESTORE_ROLLBACK_SHEET_NAME, MAIN_SHEET_NAME
    On Error GoTo 0
    SetBulkEditMode False
    RestoreApplicationState prevCalc, prevEvents, prevScreenUpdating, stateCaptured, True
    If Not ws Is Nothing Then CreateControlButtons ws, True
    MsgBox "バックアップ復元エラー: " & errDescription, vbCritical, "エラー"
End Sub

Public Sub RestoreWbsFromBackupSheet()
    RestoreWbsFromBackupSheetCore False
End Sub

Public Sub RestoreWbsFromBackupSheetSilent()
    RestoreWbsFromBackupSheetCore True
End Sub


' ============================================================================
' # 新機能 (v3)
' ============================================================================

' ==========================================
'  日付ヘッダー再生成（開始日変更対応）
' ==========================================
Private Sub RegenerateDateHeaders(ByVal ws As Worksheet, Optional ByVal startDate As Variant)
    On Error GoTo ErrorHandler

    Dim ganttStartDate As Date

    If Not IsMissing(startDate) And IsDate(startDate) Then
        ganttStartDate = CDate(startDate)
    ElseIf IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        ganttStartDate = Date
    End If

    Dim ganttStartCol As Long
    ganttStartCol = ws.Columns(COL_GANTT_START).Column

    ' クリア（書式と値）
    ' Application.Intersectを使用すると重い可能性があるため、Range指定でクリア
    ' 注意: 列幅はリセットしない（ToggleWeekendsの状態を維持したいため...いや、再生成時は3にする？）
    ' 元のロジックでは列幅セットしている (ws.Columns(colIndex).ColumnWidth = 3)
    ' 既存の列幅設定処理を追加

    Dim weekStartCol As Long
    Dim weekEndCol As Long
    Dim currentDate As Date
    Dim colIndex As Long
    Dim i As Long

    ' 週ヘッダーのマージ解除
    ws.Range(ws.Cells(ROW_WEEK_HEADER, ganttStartCol), ws.Cells(ROW_WEEK_HEADER, ganttStartCol + GANTT_DAYS - 1)).UnMerge

    For i = 1 To GANTT_DAYS
        colIndex = ganttStartCol + i - 1
        currentDate = ganttStartDate + i - 1

        ' 列幅を設定 (標準は3、ただし非表示の場合は変更しない)
        If ws.Columns(colIndex).ColumnWidth > 0 Then
            ws.Columns(colIndex).ColumnWidth = 3
        End If

        ' 7行目: 日付（日のみ）
        ws.Cells(ROW_DATE_HEADER, colIndex).Value = Day(currentDate)
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Size = 9
        ws.Cells(ROW_DATE_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_DATE_HEADER, colIndex).Interior.Color = COLOR_HEADER_BG
        ws.Cells(ROW_DATE_HEADER, colIndex).Font.Color = RGB(255, 255, 255)

        ' 8行目: 曜日
        ws.Cells(ROW_HEADER, colIndex).Value = Format$(currentDate, "aaa")
        ws.Cells(ROW_HEADER, colIndex).Font.Size = 8
        ws.Cells(ROW_HEADER, colIndex).HorizontalAlignment = xlCenter
        ws.Cells(ROW_HEADER, colIndex).Interior.Color = COLOR_HEADER_BG
        ws.Cells(ROW_HEADER, colIndex).Font.Color = RGB(255, 255, 255)

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

    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, "RegenerateDateHeaders", Err.Description
End Sub

' ==========================================
'  ガント領域の背景色クリア
' ==========================================
Private Sub ClearGanttColors(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal ganttStartCol As Long)
    On Error Resume Next
    Dim ganttEndCol As Long
    ganttEndCol = ganttStartCol + GANTT_DAYS - 1

    ' データ行のガント領域の背景色をクリア
    ws.Range(ws.Cells(ROW_DATA_START, ganttStartCol), ws.Cells(lastRow, ganttEndCol)).Interior.ColorIndex = xlNone
End Sub

' ==========================================
'  設定マスタシート作成
' ==========================================

Private Function GetSettingsWorksheet() As Worksheet
    On Error Resume Next
    Set GetSettingsWorksheet = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function GetOrCreateLogWorksheet() As Worksheet
    Dim wsLog As Worksheet

    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsLog.Name = LOG_SHEET_NAME
        wsLog.Range("A1:F1").Value = Array("timestamp", "event", "sheet", "target", "details", "user")
        wsLog.Rows(1).Font.Bold = True
        wsLog.Columns("A:F").ColumnWidth = 18
    End If

    On Error Resume Next
    wsLog.Visible = xlSheetVeryHidden
    On Error GoTo 0
    Set GetOrCreateLogWorksheet = wsLog
End Function

Public Sub LogAutomationEvent(ByVal eventName As String, Optional ByVal details As String = "", Optional ByVal targetAddress As String = "")
    On Error Resume Next

    Dim wsLog As Worksheet
    Dim nextRow As Long

    Set wsLog = GetOrCreateLogWorksheet()
    If wsLog Is Nothing Then Exit Sub

    nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    If nextRow > LOG_MAX_ROWS Then
        wsLog.Rows("2:" & (nextRow - LOG_MAX_ROWS + 1)).Delete
        nextRow = wsLog.Cells(wsLog.Rows.Count, "A").End(xlUp).Row + 1
    End If

    wsLog.Cells(nextRow, "A").Value = Now
    wsLog.Cells(nextRow, "A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    wsLog.Cells(nextRow, "B").Value = eventName
    If Not ActiveSheet Is Nothing Then wsLog.Cells(nextRow, "C").Value = ActiveSheet.Name
    wsLog.Cells(nextRow, "D").Value = targetAddress
    wsLog.Cells(nextRow, "E").Value = details
    wsLog.Cells(nextRow, "F").Value = Application.UserName
End Sub

Private Sub LogGanttRefresh(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, _
                            ByVal touchedCount As Long, ByVal deletedCount As Long)
    Dim details As String

    details = "rows=" & startRow & ":" & endRow & _
              ", touchedShapes=" & touchedCount & _
              ", deletedShapes=" & deletedCount
    LogAutomationEvent "GanttRefresh", details, ws.Name
End Sub

Private Sub AddSettingsCommandButton(ByVal wsSettings As Worksheet, ByVal shapeName As String, _
                                     ByVal caption As String, ByVal macroName As String, _
                                     ByVal leftPos As Double, ByVal topPos As Double, _
                                     ByVal buttonWidth As Double, ByVal buttonHeight As Double, _
                                     ByVal fillColor As Long)
    Dim btn As Shape

    Set btn = wsSettings.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, buttonWidth, buttonHeight)
    With btn
        .Name = shapeName
        .Fill.ForeColor.RGB = fillColor
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Characters.Text = caption
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .OnAction = macroName
    End With
End Sub

Private Sub EnsureSettingsCommandButtons(ByVal wsSettings As Worksheet)
    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim buttonWidth As Double
    Dim buttonHeight As Double

    If wsSettings Is Nothing Then Exit Sub

    On Error Resume Next
    For Each shp In wsSettings.Shapes
        If Left$(shp.Name, 13) = "Btn_Settings_" Then shp.Delete
    Next shp
    On Error GoTo 0

    wsSettings.Range("E3").Value = "WBS操作"
    wsSettings.Range("E3").Font.Bold = True
    wsSettings.Range("E4:I4").ClearContents
    wsSettings.Range("E4").Value = "退避・復元・サマリ作成はここから実行します。"
    wsSettings.Range("E4:I4").Merge
    wsSettings.Range("E4:I4").WrapText = True

    leftPos = wsSettings.Range("E5").Left
    topPos = wsSettings.Range("E5").Top
    buttonWidth = 110
    buttonHeight = 24

    AddSettingsCommandButton wsSettings, "Btn_Settings_WbsBackup", "WBS退避", "CreateWbsBackupSheet", _
                             leftPos, topPos, buttonWidth, buttonHeight, RGB(96, 73, 122)
    AddSettingsCommandButton wsSettings, "Btn_Settings_WbsRestore", "バックアップ復元", "RestoreWbsFromBackupSheet", _
                             leftPos + buttonWidth + 8, topPos, buttonWidth + 20, buttonHeight, RGB(112, 48, 160)
    AddSettingsCommandButton wsSettings, "Btn_Settings_WbsSummary", "WBSサマリ作成", "WBSRoadmapReport.CreateRoadmapOverview", _
                             leftPos + buttonWidth * 2 + 36, topPos, buttonWidth + 15, buttonHeight, RGB(84, 130, 53)

    wsSettings.Columns("E:I").ColumnWidth = 14
End Sub

Private Sub NormalizeHolidayMasterRows(ByVal wsSettings As Worksheet)
    Dim holidayDates As Object
    Dim r As Long
    Dim writeRow As Long
    Dim dateKey As String
    Dim cellValue As Variant
    Dim key As Variant

    If wsSettings Is Nothing Then Exit Sub

    Set holidayDates = CreateObject("Scripting.Dictionary")
    For r = 13 To 45
        cellValue = wsSettings.Cells(r, "A").Value
        If IsDate(cellValue) Then
            dateKey = Format$(CDate(cellValue), "yyyy-mm-dd")
            If Not holidayDates.Exists(dateKey) Then holidayDates.Add dateKey, CDate(cellValue)
        End If
    Next r

    wsSettings.Range("A13:A45").ClearContents
    writeRow = HOLIDAY_DATA_START_ROW
    For Each key In holidayDates.Keys
        wsSettings.Cells(writeRow, "A").Value = holidayDates(key)
        writeRow = writeRow + 1
        If writeRow > 45 Then Exit For
    Next key
End Sub

Private Sub EnsureBulkEditSettingSection(ByVal wsSettings As Worksheet)
    Dim bulkModeValue As String

    If wsSettings Is Nothing Then Exit Sub

    wsSettings.Range("A9").Value = "高速入力モード"
    bulkModeValue = UCase$(Trim$(CStr(wsSettings.Range("B9").Value)))
    wsSettings.Range("B9").Value = (bulkModeValue = "TRUE" Or bulkModeValue = "1" Or bulkModeValue = "ON")
    wsSettings.Range("C9").Value = "TRUE: 大量貼り付け中だけ、親集計・色分け・ガント更新を止めます。通常作業では FALSE のまま使います。"
    wsSettings.Range("A10").Value = "更新の戻し方"
    wsSettings.Range("B10").Value = "FALSE/ガント更新"
    wsSettings.Range("C10").Value = "B9 を FALSE に戻すか、メインシートの「ガント更新」を押すと、保留した更新をまとめて反映します。"
    wsSettings.Range("B9:B10").HorizontalAlignment = xlCenter
    wsSettings.Range("A9:C10").WrapText = True
    wsSettings.Range("A9:A10").Font.Bold = True
    wsSettings.Range("A9:C10").Interior.Color = RGB(255, 242, 204)
    wsSettings.Range("A9:C10").VerticalAlignment = xlCenter
    wsSettings.Rows("9:10").RowHeight = 36
    With wsSettings.Range("B9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="TRUE,FALSE"
    End With

    wsSettings.Range("A11").Value = "WBSサマリ表示階層"
    If Trim$(CStr(wsSettings.Range("B11").Value)) = "" Then wsSettings.Range("B11").Value = "LV1のみ"
    wsSettings.Range("C11").Value = "← LV1のみ / LV2まで。LV2までの場合は折りたたみ可能な詳細行を追加"
    wsSettings.Range("B11").HorizontalAlignment = xlCenter
    With wsSettings.Range("B11").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="LV1のみ,LV2まで"
    End With

    wsSettings.Range("A12").Value = "自動テストモード"
    If Trim$(CStr(wsSettings.Range("B12").Value)) = "" Then wsSettings.Range("B12").Value = False
    wsSettings.Range("C12").Value = "TRUE: 自動テスト中だけ確認メッセージをログ化します。通常利用では FALSE のまま使います。"
    wsSettings.Range("B12").HorizontalAlignment = xlCenter
    With wsSettings.Range("B12").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="TRUE,FALSE"
    End With

    wsSettings.Range("A12:C12").WrapText = True
    wsSettings.Range("A12").Font.Bold = True
    wsSettings.Range("A12:C12").Interior.Color = RGB(226, 239, 218)
    wsSettings.Rows("12:12").RowHeight = 36

    With wsSettings.Range("A9:C12").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 48
    End With
End Sub

Sub EnsureSettingsSheet()
    Dim wsSettings As Worksheet
    Set wsSettings = GetSettingsWorksheet()

    If wsSettings Is Nothing Then
        Set wsSettings = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsSettings.Name = SETTINGS_SHEET_NAME
    End If

    On Error Resume Next
    wsSettings.Range("E4:H4").UnMerge
    On Error GoTo 0
    NormalizeHolidayMasterRows wsSettings

    ' === タイトル (A1) ===
    wsSettings.Range("A1").Value = "設定マスタ"
    wsSettings.Range("A1").Font.Bold = True
    wsSettings.Range("A1").Font.Size = 14

    ' === ダブルクリック機能セクション (A3-C7) ===
    wsSettings.Range("A3").Value = "ダブルクリック機能"
    wsSettings.Range("A3").Font.Bold = True

    wsSettings.Range("A4").Value = "機能有効"
    If Trim$(CStr(wsSettings.Range("B4").Value)) = "" Then wsSettings.Range("B4").Value = True
    wsSettings.Range("C4").Value = "← TRUE: ダブルクリックで完了処理を行う"

    wsSettings.Range("A5").Value = "完了日自動入力"
    If Trim$(CStr(wsSettings.Range("B5").Value)) = "" Then wsSettings.Range("B5").Value = True
    wsSettings.Range("C5").Value = "← TRUE: 完了実績日に今日を入力"

    wsSettings.Range("A6").Value = "取り消し線"
    If Trim$(CStr(wsSettings.Range("B6").Value)) = "" Then wsSettings.Range("B6").Value = True
    wsSettings.Range("C6").Value = "← TRUE: タスクに取り消し線を入れる"

    wsSettings.Range("A7").Value = "灰色変更"
    If Trim$(CStr(wsSettings.Range("B7").Value)) = "" Then wsSettings.Range("B7").Value = True
    wsSettings.Range("C7").Value = "← TRUE: タスクを濃い灰色に変更"

    wsSettings.Columns("A").ColumnWidth = 18
    wsSettings.Columns("B").ColumnWidth = 16
    wsSettings.Columns("C").ColumnWidth = 70
    wsSettings.Range("B4:B7").HorizontalAlignment = xlCenter

    With wsSettings.Range("A4:C7").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 48
    End With

    EnsureBulkEditSettingSection wsSettings
    EnsureSettingsCommandButtons wsSettings

    ' === 祝日マスタセクション (A15, A16-A45) ===
    wsSettings.Range("A15").Value = "祝日マスタ"
    wsSettings.Range("A15").Font.Bold = True
    wsSettings.Range("A15").Interior.Color = RGB(48, 84, 150)
    wsSettings.Range("A15").Font.Color = RGB(255, 255, 255)

    wsSettings.Range("B15").Value = "【祝日マスタの使い方】"
    wsSettings.Range("B15").Font.Bold = True

    wsSettings.Range("A16:A45").NumberFormat = "yy/mm/dd"
    With wsSettings.Range("A16:A45").Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 48
    End With

    wsSettings.Range("B16").Value = "A16 以下に祝日の日付を入力してください。"
    wsSettings.Range("B17").Value = "入力した日付はガントチャート上で濃い灰色で表示されます。"
    wsSettings.Range("B19").Value = "例: 26/01/01, 26/01/13, 26/02/11 ..."
    wsSettings.Range("B19").Font.Color = RGB(128, 128, 128)
    wsSettings.Range("B21").Value = "※ ガント更新後に反映されます。"

    If ThisWorkbook.Windows.Count > 0 Then
        ActiveWindow.DisplayGridlines = False
    End If
End Sub

' ==========================================
'  設定読み込み
' ==========================================
Public Function GetSettingValue(ByVal settingRow As Long) As Boolean
    Dim wsSettings As Worksheet
    Set wsSettings = GetSettingsWorksheet()

    If wsSettings Is Nothing Then
        GetSettingValue = True
        Exit Function
    End If

    GetSettingValue = (wsSettings.Cells(settingRow, "B").Value = True)
End Function

Public Function IsBulkEditModeEnabled() As Boolean
    Dim wsSettings As Worksheet
    Set wsSettings = GetSettingsWorksheet()

    If wsSettings Is Nothing Then Exit Function
    If Trim$(CStr(wsSettings.Cells(SETTINGS_ROW_BULK_EDIT_MODE, "B").Value)) = "" Then Exit Function

    IsBulkEditModeEnabled = CBool(wsSettings.Cells(SETTINGS_ROW_BULK_EDIT_MODE, "B").Value)
End Function

Public Function IsAutomationModeEnabled() As Boolean
    Dim wsSettings As Worksheet
    Set wsSettings = GetSettingsWorksheet()

    If wsSettings Is Nothing Then Exit Function
    If Trim$(CStr(wsSettings.Cells(SETTINGS_ROW_AUTOMATION_MODE, "B").Value)) = "" Then Exit Function

    IsAutomationModeEnabled = CBool(wsSettings.Cells(SETTINGS_ROW_AUTOMATION_MODE, "B").Value)
End Function

Public Sub SetAutomationMode(ByVal isEnabled As Boolean)
    EnsureSettingsSheet
    GetSettingsWorksheet().Cells(SETTINGS_ROW_AUTOMATION_MODE, "B").Value = isEnabled
    LogAutomationEvent "AutomationMode", "enabled=" & CStr(isEnabled), SETTINGS_SHEET_NAME & "!B" & SETTINGS_ROW_AUTOMATION_MODE
End Sub

Public Function CheckInazumaRuntimeState(Optional ByVal writeLog As Boolean = True) As String
    Dim ws As Worksheet
    Dim wsSettings As Worksheet
    Dim issues As Collection
    Dim shp As Shape
    Dim bulkEnabled As Boolean
    Dim bulkButtonCount As Long
    Dim indicatorText As String
    Dim resultText As String
    Dim issue As Variant

    Set issues = New Collection
    Set ws = GetMainWorksheet()
    Set wsSettings = GetSettingsWorksheet()

    If ws Is Nothing Then issues.Add "main sheet missing"
    If wsSettings Is Nothing Then issues.Add "settings sheet missing"

    If Not wsSettings Is Nothing Then
        bulkEnabled = IsBulkEditModeEnabled()
        If bulkEnabled And Application.EnableEvents Then issues.Add "bulk mode ON but events enabled"
        If (Not bulkEnabled) And (Not Application.EnableEvents) Then issues.Add "bulk mode OFF but events disabled"
    End If

    If Not ws Is Nothing Then
        For Each shp In ws.Shapes
            If shp.Name = "Btn_BulkEdit" Then bulkButtonCount = bulkButtonCount + 1
        Next shp
        If bulkButtonCount <> 1 Then issues.Add "bulk button count=" & CStr(bulkButtonCount)

        indicatorText = CStr(ws.Range("A3").Value)
        If bulkEnabled Then
            If InStr(1, indicatorText, "高速入力 ON", vbTextCompare) = 0 Then issues.Add "indicator not ON"
        Else
            If InStr(1, indicatorText, "高速入力 OFF", vbTextCompare) = 0 Then issues.Add "indicator not OFF"
        End If
    End If

    If issues.Count = 0 Then
        resultText = "OK"
    Else
        resultText = "NG"
        For Each issue In issues
            resultText = resultText & "; " & CStr(issue)
        Next issue
    End If

    If writeLog Then LogAutomationEvent "RuntimeStateCheck", resultText, MAIN_SHEET_NAME
    CheckInazumaRuntimeState = resultText
End Function

Public Sub ShowInazumaRuntimeState()
    MsgBox CheckInazumaRuntimeState(True), vbInformation, "Inazuma runtime state"
End Sub

Public Function GetWbsSummaryDisplayDepth() As Long
    Dim wsSettings As Worksheet
    Dim settingText As String

    GetWbsSummaryDisplayDepth = 1
    EnsureSettingsSheet
    Set wsSettings = GetSettingsWorksheet()
    If wsSettings Is Nothing Then Exit Function

    settingText = Trim$(CStr(wsSettings.Cells(SETTINGS_ROW_WBS_SUMMARY_DEPTH, "B").Value))
    If settingText = "" Then Exit Function

    If InStr(1, settingText, "2", vbTextCompare) > 0 Or _
       InStr(1, settingText, "LV2", vbTextCompare) > 0 Then
        GetWbsSummaryDisplayDepth = 2
    End If
End Function

Public Sub SetBulkEditMode(ByVal isEnabled As Boolean)
    EnsureSettingsSheet
    GetSettingsWorksheet().Cells(SETTINGS_ROW_BULK_EDIT_MODE, "B").Value = isEnabled
End Sub

Private Function HasChildTaskRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal currentLevel As Long) As Boolean
    Dim r As Long
    Dim nextLevel As Variant

    For r = startRow + 1 To endRow
        nextLevel = ws.Cells(r, COL_HIERARCHY).Value
        If IsNumeric(nextLevel) Then
            If CLng(nextLevel) <= currentLevel Then Exit Function
            HasChildTaskRows = True
            Exit Function
        End If
    Next r
End Function

Public Sub RollupDevelopmentHours(ByVal targetRow As Long)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("開発LT集計")
    If ws Is Nothing Then Exit Sub

    If targetRow < ROW_DATA_START Then Exit Sub

    Dim targetLevel As Variant
    targetLevel = ws.Cells(targetRow, COL_HIERARCHY).Value
    If Not IsNumeric(targetLevel) Then Exit Sub

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow <= targetRow Then Exit Sub

    Dim totalHours As Double
    Dim invalidCount As Long
    Dim hasLeafTask As Boolean
    Dim r As Long
    Dim rowLevel As Variant
    Dim hoursValue As Double

    For r = targetRow + 1 To lastRow
        rowLevel = ws.Cells(r, COL_HIERARCHY).Value
        If IsNumeric(rowLevel) Then
            If CLng(rowLevel) <= CLng(targetLevel) Then Exit For

            If Not HasChildTaskRows(ws, r, lastRow, CLng(rowLevel)) Then
                hasLeafTask = True
                If TryParseDevelopmentHours(ws.Cells(r, COL_DEV_LT).Value, hoursValue) Then
                    totalHours = totalHours + hoursValue
                ElseIf Trim$(CStr(ws.Cells(r, COL_DEV_LT).Value)) <> "" Then
                    invalidCount = invalidCount + 1
                End If
            End If
        End If
    Next r

    If Not hasLeafTask Then Exit Sub

    ws.Cells(targetRow, COL_DEV_LT).Value = totalHours
    ws.Cells(targetRow, COL_DEV_LT).NumberFormat = DEV_HOURS_NUMBER_FORMAT

    If invalidCount > 0 Then
        MsgBox "配下タスクに集計できない開発LTが " & invalidCount & " 件ありました。", vbExclamation, "開発LT集計"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "開発LT集計エラー: " & Err.Description, vbCritical, "開発LT集計"
End Sub

' ==========================================
'  タスク行の折りたたみ/展開（右ダブルクリック用）
' ==========================================
Public Sub ToggleTaskCollapse(ByVal targetRow As Long)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("折りたたみ切替")
    If ws Is Nothing Then Exit Sub

    If targetRow < ROW_DATA_START Then Exit Sub

    Dim lvValue As Variant
    lvValue = ws.Cells(targetRow, "A").Value
    If Not IsNumeric(lvValue) Then Exit Sub
    If CLng(lvValue) <> 1 Then Exit Sub  ' LV1のみ折りたたみ対象

    Application.ScreenUpdating = False

    Dim r As Long, lastRow As Long
    lastRow = GetLastDataRow(ws)

    ' 次のLV1まで、または最終行まで
    Dim endRow As Long
    endRow = lastRow
    For r = targetRow + 1 To lastRow
        If IsNumeric(ws.Cells(r, "A").Value) Then
            If CLng(ws.Cells(r, "A").Value) = 1 Then
                endRow = r - 1
                Exit For
            End If
        End If
    Next r

    If endRow <= targetRow Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' 現在の状態を確認（最初の子行が非表示かどうか）
    Dim isHidden As Boolean
    isHidden = ws.Rows(targetRow + 1).Hidden

    ' 子行の表示/非表示を切り替え
    For r = targetRow + 1 To endRow
        ws.Rows(r).Hidden = Not isHidden
    Next r

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

' ==========================================
'  No.列の自動採番
' ==========================================
Public Sub ResetTaskRowDisplay(ByVal ws As Worksheet, ByVal targetRow As Long)
    If ws Is Nothing Then Exit Sub
    If targetRow < ROW_DATA_START Then Exit Sub
    If HasTaskContentInRow(ws, targetRow) Then Exit Sub

    ws.Cells(targetRow, COL_HIERARCHY).ClearContents
    ws.Cells(targetRow, COL_HIERARCHY).NumberFormat = "General"
    ws.Cells(targetRow, "B").ClearContents
    ws.Cells(targetRow, COL_STATUS).ClearContents
    ws.Cells(targetRow, COL_PROGRESS).ClearContents
    ws.Cells(targetRow, COL_DEV_LT).ClearContents
    ws.Cells(targetRow, COL_START_PLAN).ClearContents
    ws.Cells(targetRow, COL_END_PLAN).ClearContents
    ws.Cells(targetRow, COL_START_ACTUAL).ClearContents
    ws.Cells(targetRow, COL_END_ACTUAL).ClearContents
    If IsDisplayOnlyTaskText(CStr(ws.Cells(targetRow, "C").Value)) Then
        ws.Cells(targetRow, "C").ClearContents
    End If
    ws.Range("C" & targetRow & ":F" & targetRow).Font.Strikethrough = False
    ws.Range("C" & targetRow & ":F" & targetRow).Font.ColorIndex = xlColorIndexAutomatic
    ws.Range("C" & targetRow & ":F" & targetRow).Font.Bold = False
    ws.Cells(targetRow, "C").Font.Italic = False
    ws.Cells(targetRow, "C").HorizontalAlignment = xlGeneral
End Sub

Public Sub RenumberRowsForWorksheet(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim numArray() As Variant
    Dim r As Long
    Dim num As Long

    If ws Is Nothing Then Exit Sub

    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START

    ReDim numArray(1 To lastRow - ROW_DATA_START + 1, 1 To 1)

    num = 1
    For r = ROW_DATA_START To lastRow
        If HasTaskContentInRow(ws, r) Then
            numArray(r - ROW_DATA_START + 1, 1) = num
            num = num + 1
        Else
            numArray(r - ROW_DATA_START + 1, 1) = ""
        End If
    Next r

    ws.Range("B" & ROW_DATA_START & ":B" & lastRow).Value = numArray
End Sub

Sub RenumberRows()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("再採番")
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    RenumberRowsForWorksheet ws
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "採番エラー: " & Err.Description, vbCritical
End Sub

' ==========================================
'  日付一括シフト（祝日考慮）
' ==========================================
Sub ShiftDates()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("日付シフト", True)
    If ws Is Nothing Then Exit Sub
    Call RepairBulkEditRuntimeState(ws)

    Dim shiftDays As Variant
    shiftDays = Application.InputBox("シフトする営業日数を入力（例: 5 または -3）" & vbCrLf & _
                                     "※祝日マスタの祝日も考慮されます", _
                                     "日付シフト", 0, Type:=1)

    If VarType(shiftDays) = vbBoolean Then Exit Sub
    If shiftDays = 0 Then
        MsgBox "シフト日数が0のため処理を中止しました", vbInformation
        Exit Sub
    End If

    ' 祝日マスタ（設定マスタ内）を取得
    Dim wsSettings As Worksheet
    Dim holidays As Range
    On Error Resume Next
    Set wsSettings = ThisWorkbook.Worksheets(SETTINGS_SHEET_NAME)
    On Error GoTo ErrorHandler


    If Not wsSettings Is Nothing Then
        Dim lastHolidayRow As Long
        lastHolidayRow = wsSettings.Cells(wsSettings.Rows.Count, "A").End(xlUp).Row
        ' 設定マスタは16行目から祝日データ
        If lastHolidayRow >= HOLIDAY_DATA_START_ROW Then
            Set holidays = wsSettings.Range("A" & HOLIDAY_DATA_START_ROW & ":A" & lastHolidayRow)
        End If
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim cell As Range
    Dim shiftCount As Long
    shiftCount = 0

    For Each cell In Selection
        If IsDate(cell.Value) Then
            If holidays Is Nothing Then
                cell.Value = WorksheetFunction.WorkDay(cell.Value, CLng(shiftDays))
            Else
                cell.Value = WorksheetFunction.WorkDay(cell.Value, CLng(shiftDays), holidays)
            End If
            shiftCount = shiftCount + 1
        End If
    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox shiftCount & " 個の日付を " & shiftDays & " 営業日シフトしました" & vbCrLf & _
           "(祝日マスタ: " & IIf(holidays Is Nothing, "未使用", "使用") & ")", _
           vbInformation, "日付シフト"
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "日付シフトエラー: " & Err.Description, vbCritical
End Sub

' ==========================================
'  PDFエクスポート（当月ガント含む）
' ==========================================
Sub ExportToPDF()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = RequireMainWorksheet("PDF出力")
    If ws Is Nothing Then Exit Sub
    Call RepairBulkEditRuntimeState(ws)

    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START + 10

    ' 開始日から当月末までのガント列を計算
    Dim ganttStartDate As Date
    If IsDate(ws.Range(CELL_PROJECT_START).Value) Then
        ganttStartDate = CDate(ws.Range(CELL_PROJECT_START).Value)
    Else
        ganttStartDate = Date
    End If

    ' 当月末日を計算
    Dim monthEndDate As Date
    monthEndDate = DateSerial(Year(Date), Month(Date) + 1, 0)

    ' ガント終了列を計算
    Dim ganttEndCol As Long
    Dim daysToShow As Long
    daysToShow = monthEndDate - ganttStartDate + 1
    If daysToShow < 1 Then daysToShow = 31  ' 最低31日
    If daysToShow > GANTT_DAYS Then daysToShow = GANTT_DAYS

    ganttEndCol = ws.Columns(COL_GANTT_START).Column + daysToShow - 1

    ' 出力範囲を設定（A列から当月末のガント列まで）
    Dim exportRange As Range
    Set exportRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, ganttEndCol))

    ' ファイル保存ダイアログ
    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="InazumaGantt_" & Format(Date, "yyyymmdd"), _
        FileFilter:="PDF Files (*.pdf), *.pdf")

    If savePath = "False" Then Exit Sub

    ' 印刷設定を調整
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    exportRange.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard

    MsgBox "PDFを出力しました:" & vbCrLf & savePath & vbCrLf & vbCrLf & _
           "出力範囲: A列から" & Format(monthEndDate, "m/d") & "まで", _
           vbInformation, "PDF出力完了"
    Exit Sub

ErrorHandler:
    MsgBox "PDF出力エラー: " & Err.Description, vbCritical
End Sub

' ==========================================
'  日付バリデーション（L-O列）
' ==========================================
Public Sub ValidateDateInput(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo ErrorHandler

    If Target.Row < ROW_DATA_START Then Exit Sub
    If Target.Value = "" Then Exit Sub

    If Not IsDate(Target.Value) Then
        MsgBox "日付形式で入力してください（例: 26/01/10）", vbExclamation, "入力エラー"
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = True
        Exit Sub
    End If

    Dim startPlan As Variant, endPlan As Variant
    Dim startActual As Variant, endActual As Variant
    startPlan = ws.Cells(Target.Row, COL_START_PLAN).Value
    endPlan = ws.Cells(Target.Row, COL_END_PLAN).Value
    startActual = ws.Cells(Target.Row, COL_START_ACTUAL).Value
    endActual = ws.Cells(Target.Row, COL_END_ACTUAL).Value

    If IsDate(startPlan) And IsDate(endPlan) Then
        If CDate(startPlan) > CDate(endPlan) Then
            MsgBox "開始予定日が完了予定日より後になっています", vbExclamation, "日付エラー"
        End If
    End If

    If IsDate(startActual) And IsDate(endActual) Then
        If CDate(startActual) > CDate(endActual) Then
            MsgBox "開始実績日が完了実績日より後になっています", vbExclamation, "日付エラー"
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "日付入力の検証中にエラーが発生しました: " & Err.Description, vbExclamation, "日付エラー"
End Sub

' ==========================================
'  進捗率バリデーション（I列）
' ==========================================
Public Sub ValidateProgressInput(ByVal ws As Worksheet, ByVal Target As Range)
    On Error GoTo ErrorHandler

    Dim prevEvents As Boolean
    Dim appStateCaptured As Boolean

    prevEvents = Application.EnableEvents
    appStateCaptured = True

    If Target.Row < ROW_DATA_START Then Exit Sub
    If Target.Value = "" Then Exit Sub

    Dim normalizedRate As Double
    If Not TryParseProgressValue(Target.Value, normalizedRate) Then
        If IsAutomationModeEnabled() Then
            LogAutomationEvent "InvalidProgressInput", "value=" & CStr(Target.Value), Target.Address(False, False)
        Else
            MsgBox "進捗率は 0.7 / 70 / 70% の形式で入力してください。", vbExclamation, "入力エラー"
        End If
        Application.EnableEvents = False
        Target.ClearContents
        Application.EnableEvents = prevEvents
        Exit Sub
    End If

    Application.EnableEvents = False
    Target.Value = normalizedRate
    Application.EnableEvents = prevEvents
    Exit Sub

ErrorHandler:
    If appStateCaptured Then
        Application.EnableEvents = prevEvents
    Else
        Application.EnableEvents = True
    End If
    If IsAutomationModeEnabled() Then
        LogAutomationEvent "ProgressValidationError", Err.Description, Target.Address(False, False)
    Else
        MsgBox "進捗率の検証中にエラーが発生しました: " & Err.Description, vbExclamation, "入力エラー"
    End If
End Sub

