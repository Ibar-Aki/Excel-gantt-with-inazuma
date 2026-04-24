' ==========================================
'  InazumaGantt_v3 シートモジュール用コード
' ==========================================
Option Explicit
' このコードは「InazumaGantt_v3」シートのシートモジュールに貼り付けてください
'
' 【設定方法】
' 1. Excelで Alt+F11 を押してVBAエディタを開く
' 2. プロジェクトエクスプローラーで「InazumaGantt_v3」シートをダブルクリック
' 3. 開いたコードウィンドウに以下のコードを貼り付ける
' 4. VBAエディタを閉じる
'
' ==========================================

Private isHandlingWorksheetChange As Boolean

' API宣言（Shiftキー検知用）
#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

' データ開始行（InazumaGantt_v3モジュールと同期）
' Private Const ROW_DATA_START As Long = 9

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' タスク行のダブルクリック処理
    ' B列: 完了処理
    On Error GoTo ErrorHandler

    If Target.Row < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    ' B列(2): 完了処理
    If Target.Column <> 2 Then Exit Sub
    If InazumaGantt_v3.IsBulkEditModeEnabledAfterRuntimeRepair() Then Exit Sub

    ' 設定マスタから機能有効を確認
    If Not InazumaGantt_v3.GetSettingValue(4) Then Exit Sub

    ' 既に完了済みの場合は変更しない
    If Me.Cells(Target.Row, "H").Value = "完了" Then Exit Sub

    Application.EnableEvents = False

    ' 進捗率を100%に
    Me.Cells(Target.Row, "I").Value = 1

    ' 状況を「完了」に
    Me.Cells(Target.Row, "H").Value = "完了"

    ' 設定：完了日自動入力
    If InazumaGantt_v3.GetSettingValue(5) Then
        If IsDate(Me.Cells(Target.Row, InazumaGantt_v3.COL_START_ACTUAL).Value) Then
            If Trim$(CStr(Me.Cells(Target.Row, InazumaGantt_v3.COL_END_ACTUAL).Value)) = "" Then
                Me.Cells(Target.Row, InazumaGantt_v3.COL_END_ACTUAL).Value = Date
            End If
        End If
    End If

    ' 設定：取り消し線
    If InazumaGantt_v3.GetSettingValue(6) Then
        Me.Range("C" & Target.Row & ":F" & Target.Row).Font.Strikethrough = True
    End If

    ' 設定：濃い灰色に変更
    If InazumaGantt_v3.GetSettingValue(7) Then
        Me.Range("C" & Target.Row & ":F" & Target.Row).Font.Color = RGB(128, 128, 128)
    End If

    Application.EnableEvents = True
    Cancel = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    ' Shift + 右クリック: 開発LT集計または折りたたみ/展開
    On Error GoTo ErrorHandler

    ' Shiftキーが押されていない場合は通常の右クリックメニュー
    If (GetKeyState(vbKeyShift) And &H8000) = 0 Then Exit Sub

    If Target.Row < InazumaGantt_v3.ROW_DATA_START Then Exit Sub
    If InazumaGantt_v3.IsBulkEditModeEnabledAfterRuntimeRepair() Then Exit Sub

    If Target.Column = Me.Columns(InazumaGantt_v3.COL_DEV_LT).Column Then
        Application.EnableEvents = False
        WBSParentRollup.RecalculateTaskRowAndAncestors Me, Target.Row
        WBSParentRollup.RefreshTaskAlertMarkersForRowAndAncestors Me, Target.Row
        Application.EnableEvents = True
        Cancel = True
        Exit Sub
    End If

    ' C-F列(3-6)でのみ有効
    If Target.Column >= 3 And Target.Column <= 6 Then
        InazumaGantt_v3.ToggleTaskCollapse Target.Row
        Cancel = True
    End If
    Exit Sub

ErrorHandler:
    ' エラーは無視
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo ErrorHandler
    Dim affectedRows As Object
    Dim bulkEditMode As Boolean
    Dim prevCalc As XlCalculation
    Set affectedRows = CreateObject("Scripting.Dictionary")

    If isHandlingWorksheetChange Then Exit Sub
    If Target Is Nothing Then Exit Sub

    isHandlingWorksheetChange = True
    bulkEditMode = InazumaGantt_v3.IsBulkEditModeEnabledAfterRuntimeRepair()
    If bulkEditMode Then
        Application.StatusBar = "高速入力中: Ctrl+Z を優先し、自動更新を停止しています"
        isHandlingWorksheetChange = False
        Exit Sub
    End If
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' タスク入力列（C～F列）に変更があった場合
    If Not Intersect(Target, Me.Range("C:F")) Is Nothing Then
        PrepareChangedTaskRows Intersect(Target, Me.Range("C:F")), affectedRows
        RefreshTaskStructureForRange Intersect(Target, Me.Range("C:F"))
    End If

    ' 状況列（H列）または進捗率列（I列）に変更があった場合、相互同期
    If Not Intersect(Target, Me.Range("H:I")) Is Nothing Then
        Dim statusProgressCell As Range
        For Each statusProgressCell In Intersect(Target, Me.Range("H:I"))
            If statusProgressCell.Row >= InazumaGantt_v3.ROW_DATA_START Then
                If statusProgressCell.Column = Me.Columns("I").Column Then
                    InazumaGantt_v3.ValidateProgressInput Me, Me.Cells(statusProgressCell.Row, "I")
                End If

                If InazumaGantt_v3.HasTaskContentInRow(Me, statusProgressCell.Row) Then
                    If statusProgressCell.Column = Me.Columns("I").Column Then
                        UpdateStatusByProgress statusProgressCell.Row
                    Else
                        InazumaGantt_v3.SyncTaskStatusAndProgressRow Me, statusProgressCell.Row
                    End If
                Else
                    InazumaGantt_v3.ResetTaskRowDisplay Me, statusProgressCell.Row
                End If
                CollectAffectedRow affectedRows, statusProgressCell.Row
            End If
        Next statusProgressCell
    End If

    ' 開発LT列（K列）の入力を検証
    If Not Intersect(Target, Me.Columns(InazumaGantt_v3.COL_DEV_LT)) Is Nothing Then
        Dim ltCell As Range
        For Each ltCell In Intersect(Target, Me.Columns(InazumaGantt_v3.COL_DEV_LT))
            If ltCell.Row >= InazumaGantt_v3.ROW_DATA_START Then
                InazumaGantt_v3.ValidateDevelopmentHoursInput Me, ltCell
                CollectAffectedRow affectedRows, ltCell.Row
            End If
        Next ltCell
    End If

    ' 日付列（L-O列）の入力を検証
    If Not Intersect(Target, Me.Range(InazumaGantt_v3.COL_START_PLAN & ":" & InazumaGantt_v3.COL_END_ACTUAL)) Is Nothing Then
        Dim validateCell As Range
        For Each validateCell In Intersect(Target, Me.Range(InazumaGantt_v3.COL_START_PLAN & ":" & InazumaGantt_v3.COL_END_ACTUAL))
            If validateCell.Row >= InazumaGantt_v3.ROW_DATA_START Then
                InazumaGantt_v3.ValidateDateInput Me, validateCell
                CollectAffectedRow affectedRows, validateCell.Row
            End If
        Next validateCell
    End If

    ' 予定日付列（L, M列）に土日祝日を入力した場合に確認メッセージ
    If Not bulkEditMode And Not Intersect(Target, Me.Range(InazumaGantt_v3.COL_START_PLAN & ":" & InazumaGantt_v3.COL_END_PLAN)) Is Nothing Then
        Dim planDateCell As Range
        Dim inputDate As Date
        Dim isWeekend As Boolean
        Dim isHoliday As Boolean
        Dim warningMsg As String

        For Each planDateCell In Intersect(Target, Me.Range(InazumaGantt_v3.COL_START_PLAN & ":" & InazumaGantt_v3.COL_END_PLAN))
            If planDateCell.Row >= InazumaGantt_v3.ROW_DATA_START Then
                If IsDate(planDateCell.Value) Then
                    inputDate = CDate(planDateCell.Value)
                    isWeekend = (Weekday(inputDate, vbMonday) >= 6)
                    isHoliday = CheckHoliday(inputDate)

                    If isWeekend Or isHoliday Then
                        If isHoliday Then
                            warningMsg = "祝日"
                        ElseIf Weekday(inputDate, vbMonday) = 6 Then
                            warningMsg = "土曜日"
                        Else
                            warningMsg = "日曜日"
                        End If

                        If MsgBox(Format(inputDate, "yy/mm/dd") & " は " & warningMsg & " です。" & vbCrLf & _
                                  "この日付を入力しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then
                            Application.EnableEvents = False
                            planDateCell.ClearContents
                            Application.EnableEvents = True
                        End If
                    End If
                End If
                CollectAffectedRow affectedRows, planDateCell.Row
            End If
        Next planDateCell
    End If

    RefreshAffectedRowDisplayState affectedRows

    ApplyRollupAndAlerts affectedRows

    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    isHandlingWorksheetChange = False
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    isHandlingWorksheetChange = False
End Sub

' ==========================================
'  祝日チェック
' ==========================================
Private Function CheckHoliday(ByVal targetDate As Date) As Boolean
    Dim wsSettings As Worksheet
    On Error Resume Next
    ' InazumaGantt_v3の定数を使用
    Set wsSettings = ThisWorkbook.Worksheets(InazumaGantt_v3.SETTINGS_SHEET_NAME)
    On Error GoTo 0

    CheckHoliday = False
    If wsSettings Is Nothing Then Exit Function

    ' 祝日マスタハ設定マスタのA13から
    Dim lastRow As Long
    lastRow = wsSettings.Cells(wsSettings.Rows.Count, "A").End(xlUp).Row
    If lastRow < InazumaGantt_v3.HOLIDAY_DATA_START_ROW Then Exit Function

    Dim r As Long
    For r = InazumaGantt_v3.HOLIDAY_DATA_START_ROW To lastRow
        If IsDate(wsSettings.Cells(r, "A").Value) Then
            If CDate(wsSettings.Cells(r, "A").Value) = targetDate Then
                CheckHoliday = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub UpdateStatusByProgress(ByVal targetRow As Long)
    Dim progressValue As Variant
    Dim rate As Double

    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub
    If Not InazumaGantt_v3.HasTaskContentInRow(Me, targetRow) Then
        InazumaGantt_v3.ResetTaskRowDisplay Me, targetRow
        Exit Sub
    End If

    progressValue = Me.Cells(targetRow, "I").Value
    If Trim$(CStr(progressValue)) = "" Then
        Me.Cells(targetRow, "H").Value = InazumaGantt_v3.STATUS_NOT_STARTED
        Me.Cells(targetRow, "I").Value = 0
        Exit Sub
    End If

    rate = InazumaGantt_v3.NormalizeProgressValue(progressValue, 0)
    If rate >= 1 Then
        Me.Cells(targetRow, "H").Value = InazumaGantt_v3.STATUS_COMPLETED
        Me.Cells(targetRow, "I").Value = 1
    ElseIf rate <= 0 Then
        Me.Cells(targetRow, "H").Value = InazumaGantt_v3.STATUS_NOT_STARTED
        Me.Cells(targetRow, "I").Value = 0
    Else
        Me.Cells(targetRow, "H").Value = InazumaGantt_v3.STATUS_IN_PROGRESS
        Me.Cells(targetRow, "I").Value = rate
    End If
End Sub

Private Sub CollectAffectedRow(ByVal affectedRows As Object, ByVal rowNumber As Long)
    If rowNumber < InazumaGantt_v3.ROW_DATA_START Then Exit Sub
    affectedRows(CStr(rowNumber)) = True
End Sub

Private Sub PrepareChangedTaskRows(ByVal changedRange As Range, ByVal affectedRows As Object)
    Dim area As Range
    Dim currentRow As Long
    Dim lastSupportedRow As Long

    If changedRange Is Nothing Then Exit Sub

    lastSupportedRow = InazumaGantt_v3.ROW_DATA_START + InazumaGantt_v3.DATA_ROWS_DEFAULT - 1
    For Each area In changedRange.Areas
        For currentRow = area.Row To area.Row + area.Rows.Count - 1
            If currentRow >= InazumaGantt_v3.ROW_DATA_START And currentRow <= lastSupportedRow Then
                NormalizeTaskRowState currentRow
                CollectAffectedRow affectedRows, currentRow - 1
                CollectAffectedRow affectedRows, currentRow
                CollectAffectedRow affectedRows, currentRow + 1
            End If
        Next currentRow
    Next area
End Sub

Private Sub NormalizeTaskRowState(ByVal targetRow As Long)
    If targetRow < InazumaGantt_v3.ROW_DATA_START Then Exit Sub

    If InazumaGantt_v3.HasTaskContentInRow(Me, targetRow) Then
        InazumaGantt_v3.SyncTaskStatusAndProgressRow Me, targetRow
    Else
        InazumaGantt_v3.ResetTaskRowDisplay Me, targetRow
    End If

    InazumaGantt_v3.RefreshTaskRowDisplayState Me, targetRow
End Sub

Private Sub RefreshTaskStructureForRange(ByVal changedRange As Range)
    Dim startRow As Long
    Dim endRow As Long
    Dim lastSupportedRow As Long

    If changedRange Is Nothing Then Exit Sub

    lastSupportedRow = InazumaGantt_v3.ROW_DATA_START + InazumaGantt_v3.DATA_ROWS_DEFAULT - 1
    startRow = changedRange.Row - 1
    If startRow < InazumaGantt_v3.ROW_DATA_START Then startRow = InazumaGantt_v3.ROW_DATA_START
    endRow = changedRange.Row + changedRange.Rows.Count
    If endRow > lastSupportedRow Then endRow = lastSupportedRow

    InazumaGantt_v3.AutoDetectTaskLevelsInRange Me, startRow, endRow
    InazumaGantt_v3.RenumberRowsForWorksheet Me
End Sub

Private Sub ApplyRollupAndAlerts(ByVal affectedRows As Object)
    Dim rowKey As Variant

    For Each rowKey In affectedRows.Keys
        WBSParentRollup.RecalculateTaskRowAndAncestors Me, CLng(rowKey)
        WBSParentRollup.RefreshTaskAlertMarkersForRowAndAncestors Me, CLng(rowKey)
    Next rowKey
End Sub

Private Sub RefreshAffectedRowDisplayState(ByVal affectedRows As Object)
    Dim rowKey As Variant

    For Each rowKey In affectedRows.Keys
        InazumaGantt_v3.RefreshTaskRowDisplayState Me, CLng(rowKey)
    Next rowKey
End Sub
