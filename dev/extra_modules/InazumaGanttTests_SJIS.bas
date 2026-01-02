Attribute VB_Name = "InazumaGanttTests"
Option Explicit

' ==========================================
'  InazumaGantt v2 テストモジュール
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
    Debug.Print "InazumaGantt v2 テスト開始"
    Debug.Print "=========================================="
    
    ' テスト実行
    Call Test_GetTaskColumnByLevel
    Call Test_AutoDetectTaskLevel
    Call Test_ParseProgressValue
    Call Test_GetLastDataRow
    
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
