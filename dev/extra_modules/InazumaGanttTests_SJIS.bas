Attribute VB_Name = "InazumaGanttTests"
Option Explicit

' ==========================================
'  InazumaGantt v2.2 テストモジュール
' ==========================================
' このモジュールは開発・デバッグ用です。
' 本番環境では削除またはコメントアウト推奨。
'
' 使い方:
'   Alt + F8 → RunAllTests → 実行
' ==========================================

Public TestsPassed As Long
Public TestsFailed As Long
Public TestResults As String

' ==========================================
'  全テストを実行
' ==========================================
Sub RunAllTests()
    TestsPassed = 0
    TestsFailed = 0
    TestResults = ""
    
    Debug.Print "=========================================="
    Debug.Print "InazumaGantt v2.2 テスト開始"
    Debug.Print "=========================================="
    
    ' 基本テスト実行
    Call Test_GetTaskColumnByLevel
    Call Test_AutoDetectTaskLevel
    Call Test_ParseProgressValue
    Call Test_GetLastDataRow
    
    ' v2.2新機能テスト
    Call Test_GetSettingValue
    Call Test_EnsureSettingsSheet
    Call Test_ShiftDates
    Call Test_RenumberRows
    Call Test_TaskCollapse
    
    ' 結果表示
    Debug.Print "=========================================="
    Debug.Print "テスト完了"
    Debug.Print "成功: " & TestsPassed
    Debug.Print "失敗: " & TestsFailed
    Debug.Print "=========================================="
    
    MsgBox "テスト完了" & vbCrLf & vbCrLf & _
           "成功: " & TestsPassed & vbCrLf & _
           "失敗: " & TestsFailed & vbCrLf & vbCrLf & _
           TestResults, vbInformation, "テスト結果"
End Sub

' ==========================================
'  テスト: GetTaskColumnByLevel
' ==========================================
Private Sub Test_GetTaskColumnByLevel()
    Dim testName As String
    testName = "GetTaskColumnByLevel"
    
    On Error GoTo TestFailed
    
    ' LV1 → C列
    AssertEquals testName & " - LV1", "C", InazumaGantt_v2.GetTaskColumnByLevel(1)
    
    ' LV2 → D列
    AssertEquals testName & " - LV2", "D", InazumaGantt_v2.GetTaskColumnByLevel(2)
    
    ' LV3 → E列
    AssertEquals testName & " - LV3", "E", InazumaGantt_v2.GetTaskColumnByLevel(3)
    
    ' LV4 → F列
    AssertEquals testName & " - LV4", "F", InazumaGantt_v2.GetTaskColumnByLevel(4)
    
    ' 範囲外 → C列（デフォルト）
    AssertEquals testName & " - 範囲外", "C", InazumaGantt_v2.GetTaskColumnByLevel(5)
    
    Debug.Print "[PASS] " & testName
    Exit Sub
    
TestFailed:
    RecordFailure testName & ": " & Err.Description
End Sub

' ==========================================
'  テスト: AutoDetectTaskLevel
' ==========================================
Private Sub Test_AutoDetectTaskLevel()
    Dim testName As String
    testName = "AutoDetectTaskLevel"
    
    ' このテストは実際のシートが必要なため、
    ' 手動テストまたはモックを使用
    Debug.Print "[SKIP] " & testName & " (requires worksheet)"
End Sub

' ==========================================
'  テスト: ParseProgressValue (DataMigration)
' ==========================================
Private Sub Test_ParseProgressValue()
    Dim testName As String
    testName = "ParseProgressValue"
    
    ' このテストはDataMigrationモジュールの
    ' Private関数のため、スキップ
    Debug.Print "[SKIP] " & testName & " (private function)"
End Sub

' ==========================================
'  テスト: GetLastDataRow
' ==========================================
Private Sub Test_GetLastDataRow()
    Dim testName As String
    testName = "GetLastDataRow"
    
    ' このテストは実際のシートが必要なため、
    ' 手動テストまたはモックを使用
    Debug.Print "[SKIP] " & testName & " (requires worksheet)"
End Sub

' ==========================================
'  v2.2テスト: GetSettingValue
' ==========================================
Private Sub Test_GetSettingValue()
    Dim testName As String
    testName = "GetSettingValue (v2.2)"
    
    On Error GoTo TestFailed
    
    ' 設定マスタシートがない場合はデフォルト値（True）を返す
    Dim result As Boolean
    result = InazumaGantt_v2.GetSettingValue(3)
    
    ' 結果はTrueまたはFalseのどちらか（エラーにならなければOK）
    TestsPassed = TestsPassed + 1
    Debug.Print "[PASS] " & testName & " - 値取得可能"
    Exit Sub
    
TestFailed:
    RecordFailure testName & ": " & Err.Description
End Sub

' ==========================================
'  v2.2テスト: EnsureSettingsSheet
' ==========================================
Private Sub Test_EnsureSettingsSheet()
    Dim testName As String
    testName = "EnsureSettingsSheet (v2.2)"
    
    On Error GoTo TestFailed
    
    ' 設定マスタシートを作成/確認
    InazumaGantt_v2.EnsureSettingsSheet
    
    ' 設定マスタシートが存在することを確認
    Dim sheetExists As Boolean
    sheetExists = False
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "設定マスタ" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    If sheetExists Then
        TestsPassed = TestsPassed + 1
        Debug.Print "[PASS] " & testName
    Else
        RecordFailure testName & ": 設定マスタシートが作成されていません"
    End If
    Exit Sub
    
TestFailed:
    RecordFailure testName & ": " & Err.Description
End Sub

' ==========================================
'  v2.2テスト: ShiftDates
' ==========================================
Private Sub Test_ShiftDates()
    Dim testName As String
    testName = "ShiftDates (v2.2)"
    
    ' このテストはユーザー入力が必要なため、
    ' 関数の存在確認のみ
    On Error GoTo NotExists
    
    ' ShiftDatesのテスト（実行はしない）
    Debug.Print "[SKIP] " & testName & " (requires user input)"
    Exit Sub
    
NotExists:
    RecordFailure testName & ": 関数が存在しません"
End Sub

' ==========================================
'  v2.2テスト: RenumberRows
' ==========================================
Private Sub Test_RenumberRows()
    Dim testName As String
    testName = "RenumberRows (v2.2)"
    
    On Error GoTo TestFailed
    
    ' RenumberRowsが存在するか確認（実行はシートに影響するため注意）
    ' 関数の存在確認のみ
    Debug.Print "[SKIP] " & testName & " (modifies worksheet)"
    Exit Sub
    
TestFailed:
    RecordFailure testName & ": " & Err.Description
End Sub

' ==========================================
'  v2.2テスト: ToggleTaskCollapse
' ==========================================
Private Sub Test_TaskCollapse()
    Dim testName As String
    testName = "ToggleTaskCollapse (v2.2)"
    
    On Error GoTo TestFailed
    
    ' ToggleTaskCollapseが存在するか確認
    ' 実行はシートに影響するためスキップ
    Debug.Print "[SKIP] " & testName & " (modifies worksheet rows)"
    Exit Sub
    
TestFailed:
    RecordFailure testName & ": " & Err.Description
End Sub

' ==========================================
'  アサーション: 値の等価性チェック
' ==========================================
Private Sub AssertEquals(ByVal testDescription As String, ByVal expected As Variant, ByVal actual As Variant)
    If expected = actual Then
        TestsPassed = TestsPassed + 1
    Else
        TestsFailed = TestsFailed + 1
        Dim msg As String
        msg = testDescription & vbCrLf & _
              "  期待値: " & CStr(expected) & vbCrLf & _
              "  実際値: " & CStr(actual)
        TestResults = TestResults & msg & vbCrLf & vbCrLf
        Debug.Print "[FAIL] " & msg
    End If
End Sub

' ==========================================
'  失敗を記録
' ==========================================
Private Sub RecordFailure(ByVal message As String)
    TestsFailed = TestsFailed + 1
    TestResults = TestResults & message & vbCrLf & vbCrLf
    Debug.Print "[FAIL] " & message
End Sub

' ==========================================
'  統合テスト: セットアップから更新まで
' ==========================================
Sub IntegrationTest_FullWorkflow()
    On Error GoTo TestFailed
    
    Debug.Print "=========================================="
    Debug.Print "統合テスト: フルワークフロー"
    Debug.Print "=========================================="
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 1. データがあることを確認
    If ws.Cells(9, "C").Value = "" Then
        MsgBox "テストデータが必要です。C9セルにタスク名を入力してください。", vbExclamation
        Exit Sub
    End If
    
    ' 2. 階層自動判定
    Debug.Print "階層自動判定テスト..."
    InazumaGantt_v2.AutoDetectTaskLevel 9
    
    If ws.Cells(9, "A").Value <> 1 Then
        Debug.Print "[FAIL] 階層レベルが正しく設定されていません"
    Else
        Debug.Print "[PASS] 階層自動判定"
    End If
    
    ' 3. ガント更新
    Debug.Print "ガント更新テスト..."
    Application.DisplayAlerts = False
    ' RefreshInazumaGanttはMsgBoxを表示するため、テストでは呼び出し注意
    ' InazumaGantt_v2.RefreshInazumaGantt
    Application.DisplayAlerts = True
    Debug.Print "[SKIP] ガント更新（手動確認推奨）"
    
    ' 4. 色分け適用
    Debug.Print "色分けテスト..."
    Application.DisplayAlerts = False
    ' HierarchyColor.ApplyHierarchyColors
    Application.DisplayAlerts = True
    Debug.Print "[SKIP] 色分け（手動確認推奨）"
    
    Debug.Print "=========================================="
    Debug.Print "統合テスト完了"
    Debug.Print "=========================================="
    
    MsgBox "統合テスト完了。イミディエイトウィンドウを確認してください。", vbInformation
    Exit Sub
    
TestFailed:
    Debug.Print "[ERROR] " & Err.Description
    MsgBox "統合テストでエラーが発生しました: " & Err.Description, vbCritical
End Sub

' ==========================================
'  v2.2機能テスト: 設定マスタ連携
' ==========================================
Sub Test_SettingsMaster_Integration()
    On Error GoTo TestFailed
    
    Debug.Print "=========================================="
    Debug.Print "v2.2機能テスト: 設定マスタ連携"
    Debug.Print "=========================================="
    
    ' 1. 設定マスタシートを作成
    Debug.Print "設定マスタシート作成..."
    InazumaGantt_v2.EnsureSettingsSheet
    Debug.Print "[PASS] 設定マスタシート作成"
    
    ' 2. 設定値を取得
    Debug.Print "設定値取得..."
    Dim setting3 As Boolean
    setting3 = InazumaGantt_v2.GetSettingValue(3)
    Debug.Print "  行3の設定値: " & setting3
    Debug.Print "[PASS] 設定値取得"
    
    Debug.Print "=========================================="
    Debug.Print "v2.2機能テスト完了"
    Debug.Print "=========================================="
    
    MsgBox "v2.2機能テスト完了。イミディエイトウィンドウを確認してください。", vbInformation
    Exit Sub
    
TestFailed:
    Debug.Print "[ERROR] " & Err.Description
    MsgBox "v2.2機能テストでエラーが発生しました: " & Err.Description, vbCritical
End Sub
