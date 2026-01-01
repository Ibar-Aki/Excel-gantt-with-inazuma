' ==========================================
'  InazumaGantt_v2 シートモジュール用コード
' ==========================================
' このコードは「InazumaGantt_v2」シートのシートモジュールに貼り付けてください
'
' 【設定方法】
' 1. Excelで Alt+F11 を押してVBAエディタを開く
' 2. プロジェクトエクスプローラーで「InazumaGantt_v2」シートをダブルクリック
' 3. 開いたコードウィンドウに以下のコードを貼り付ける
' 4. VBAエディタを閉じる
'
' ==========================================

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' タスク行のダブルクリックで完了処理を実行
    On Error Resume Next
    
    ' 標準モジュールの関数を呼び出し
    InazumaGantt_v2.CompleteTaskByDoubleClick Target
    
    ' Excelの既定のダブルクリック動作（セル編集）をキャンセル
    Cancel = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    Application.EnableEvents = False
    
    ' タスク入力列（C～F列）に変更があった場合、階層を自動判定
    If Not Intersect(Target, Me.Range("C:F")) Is Nothing Then
        ' 変更された行の階層レベルを自動判定
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Range("C:F"))
            If cell.Row >= 9 Then ' データ開始行以降
                InazumaGantt_v2.AutoDetectTaskLevel cell.Row
            End If
        Next cell
    End If
    
    ' 進捗率列（I列）に変更があった場合、状況を自動更新
    If Not Intersect(Target, Me.Columns("I")) Is Nothing Then
        Dim progressCell As Range
        For Each progressCell In Intersect(Target, Me.Columns("I"))
            If progressCell.Row >= 9 Then ' データ開始行以降
                UpdateStatusByProgress progressCell.Row
            End If
        Next progressCell
    End If
    
    Application.EnableEvents = True
End Sub

' ==========================================
'  進捗率から状況を自動更新（1行のみ処理）
' ==========================================
Private Sub UpdateStatusByProgress(ByVal targetRow As Long)
    Dim progressValue As Variant
    Dim rate As Double
    
    ' 進捗率を取得
    progressValue = Me.Cells(targetRow, "I").Value
    
    ' 空の場合は未着手
    If Trim$(CStr(progressValue)) = "" Then
        Me.Cells(targetRow, "H").Value = "未着手"
        Exit Sub
    End If
    
    ' 数値に変換
    If IsNumeric(progressValue) Then
        rate = CDbl(progressValue)
        
        ' 100超の値は100%として扱う
        If rate > 1 Then rate = rate / 100
        
        ' 状況を設定
        If rate >= 1 Then
            Me.Cells(targetRow, "H").Value = "完了"
        ElseIf rate <= 0 Then
            Me.Cells(targetRow, "H").Value = "未着手"
        Else
            Me.Cells(targetRow, "H").Value = "進行中"
        End If
    End If
End Sub
