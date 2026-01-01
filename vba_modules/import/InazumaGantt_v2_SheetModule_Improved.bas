' ==========================================
'  InazumaGantt_v2 シートモジュール用コード（改善版）
' ==========================================
' このコードは「InazumaGantt_v2」シートのシートモジュールに貼り付けてください
'
' 【設定方法】
' 1. Excelで Alt+F11 を押してVBAエディタを開く
' 2. プロジェクトエクスプローラーで「InazumaGantt_v2」シートをダブルクリック
' 3. 開いたコードウィンドウに以下のコードを貼り付ける
' 4. VBAエディタを閉じる
'
' 【改善点】
' - マジックナンバー削除（9 → ROW_DATA_START）
' - エラーハンドリング改善（ErrorHandler使用）
' - 入力値検証追加
' ==========================================

' 定数定義（マジックナンバー対策）
Private Const ROW_DATA_START As Long = 9  ' データ開始行

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    ' 標準モジュールの関数を呼び出し
    InazumaGantt_v2.CompleteTaskByDoubleClick Target
    
    ' Excelの既定のダブルクリック動作（セル編集）をキャンセル
    Cancel = True
    Exit Sub
    
ErrorHandler:
    Cancel = True
    ' ErrorHandlerモジュールがインポートされている場合のみ使用
    ' ErrorHandler.HandleError "SheetModule", "Worksheet_BeforeDoubleClick", _
    '                          "ダブルクリック処理でエラーが発生しました。"
    Debug.Print "Worksheet_BeforeDoubleClick Error: " & Err.Description
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    
    ' タスク入力列（C～F列）に変更があった場合、階層を自動判定
    If Not Intersect(Target, Me.Range("C:F")) Is Nothing Then
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Range("C:F"))
            If cell.Row >= ROW_DATA_START Then
                InazumaGantt_v2.AutoDetectTaskLevel cell.Row
            End If
        Next cell
    End If
    
    ' 進捗率列（I列）に変更があった場合、状況を自動更新
    If Not Intersect(Target, Me.Columns("I")) Is Nothing Then
        Dim progressCell As Range
        For Each progressCell In Intersect(Target, Me.Columns("I"))
            If progressCell.Row >= ROW_DATA_START Then
                UpdateStatusByProgress progressCell.Row
            End If
        Next progressCell
    End If
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    Debug.Print "Worksheet_Change Error: " & Err.Description
End Sub

' ==========================================
'  進捗率から状況を自動更新（1行のみ処理）
' ==========================================
Private Sub UpdateStatusByProgress(ByVal targetRow As Long)
    On Error GoTo ErrorHandler
    
    Dim progressValue As Variant
    Dim rate As Double
    
    ' 進捗率を取得
    progressValue = Me.Cells(targetRow, "I").Value
    
    ' 空の場合は未着手
    If Trim$(CStr(progressValue)) = "" Then
        Me.Cells(targetRow, "H").Value = "未着手"
        Exit Sub
    End If
    
    ' 数値チェック
    If Not IsNumeric(progressValue) Then
        Exit Sub  ' 数値でない場合はスキップ
    End If
    
    rate = CDbl(progressValue)
    
    ' 範囲チェック
    If rate < 0 Or rate > 100 Then
        ' 範囲外の場合はスキップ
        Exit Sub
    End If
    
    ' 100超の値は100%として扱う（パーセント表記対応）
    If rate > 1 Then rate = rate / 100
    
    ' 状況を設定
    If rate >= 1 Then
        Me.Cells(targetRow, "H").Value = "完了"
    ElseIf rate <= 0 Then
        Me.Cells(targetRow, "H").Value = "未着手"
    Else
        Me.Cells(targetRow, "H").Value = "進行中"
    End If
    Exit Sub
    
ErrorHandler:
    Debug.Print "UpdateStatusByProgress Error (Row " & targetRow & "): " & Err.Description
End Sub
