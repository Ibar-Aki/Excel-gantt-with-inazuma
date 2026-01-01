Attribute VB_Name = "DataMigration"
Option Explicit

' ==========================================
'  既存ガントチャートからv2への移管マクロ
' ==========================================
' 
' 【既存形式】
' A列: LV | B列: タスク | C列: 担当者 | D列: 進捗状況 | E列: 開始 | F列: 終了
'
' 【v2形式】
' A列: LV | B列: No. | C～F列: TASK(階層別) | G列: 詳細 | H列: 状況 | I列: 進捗率
' J列: 担当 | K列: 開始予定 | L列: 完了予定 | M列: 開始実績 | N列: 完了実績
'
' ==========================================

' 既存シートの列定義
Private Const OLD_COL_LV As String = "A"
Private Const OLD_COL_TASK As String = "B"
Private Const OLD_COL_ASSIGNEE As String = "C"
Private Const OLD_COL_PROGRESS As String = "D"
Private Const OLD_COL_START As String = "E"
Private Const OLD_COL_END As String = "F"
Private Const OLD_ROW_DATA_START As Long = 8  ' データ開始行

' v2シートの列定義
Private Const V2_COL_LV As String = "A"
Private Const V2_COL_NO As String = "B"
Private Const V2_COL_TASK_LV1 As String = "C"
Private Const V2_COL_TASK_LV2 As String = "D"
Private Const V2_COL_TASK_LV3 As String = "E"
Private Const V2_COL_TASK_LV4 As String = "F"
Private Const V2_COL_DETAIL As String = "G"
Private Const V2_COL_STATUS As String = "H"
Private Const V2_COL_PROGRESS As String = "I"
Private Const V2_COL_ASSIGNEE As String = "J"
Private Const V2_COL_START_PLAN As String = "K"
Private Const V2_COL_END_PLAN As String = "L"
Private Const V2_COL_START_ACTUAL As String = "M"
Private Const V2_COL_END_ACTUAL As String = "N"
Private Const V2_ROW_DATA_START As Long = 9   ' データ開始行

' ==========================================
'  メイン移管処理
' ==========================================
Sub MigrateToV2Format()
    On Error GoTo ErrorHandler
    
    ' 既存シートを選択
    Dim oldSheet As Worksheet
    Set oldSheet = ActiveSheet
    
    ' 確認ダイアログ
    Dim result As VbMsgBoxResult
    result = MsgBox("このシート「" & oldSheet.Name & "」のデータをv2形式に移管しますか？" & vbCrLf & vbCrLf & _
                    "移管先: InazumaGantt_v2 シート" & vbCrLf & vbCrLf & _
                    "※既存シートは変更されません", _
                    vbYesNo + vbQuestion, "データ移管")
    
    If result = vbNo Then Exit Sub
    
    ' v2シートを探す
    Dim v2Sheet As Worksheet
    On Error Resume Next
    Set v2Sheet = ThisWorkbook.Worksheets("InazumaGantt_v2")
    On Error GoTo ErrorHandler
    
    If v2Sheet Is Nothing Then
        MsgBox "移管先シート「InazumaGantt_v2」が見つかりません。" & vbCrLf & _
               "先にSetupInazumaGanttマクロを実行してシートを作成してください。", _
               vbCritical, "エラー"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' データ最終行を取得
    Dim lastRow As Long
    lastRow = oldSheet.Cells(oldSheet.Rows.Count, OLD_COL_TASK).End(xlUp).Row
    
    If lastRow < OLD_ROW_DATA_START Then
        MsgBox "移管するデータが見つかりません。", vbExclamation, "データ移管"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' v2シートのデータをクリア（データ開始行以降）
    Dim clearRange As Range
    Set clearRange = v2Sheet.Range("A" & V2_ROW_DATA_START & ":N" & V2_ROW_DATA_START + 200)
    clearRange.ClearContents
    clearRange.Interior.ColorIndex = xlNone
    
    ' データを移管
    Dim oldRow As Long
    Dim v2Row As Long
    Dim taskLevel As Long
    Dim taskText As String
    Dim progressText As String
    Dim progressValue As Double
    Dim taskNo As Long
    
    v2Row = V2_ROW_DATA_START
    taskNo = 1
    
    For oldRow = OLD_ROW_DATA_START To lastRow
        ' LVを取得
        taskLevel = 0
        If IsNumeric(oldSheet.Cells(oldRow, OLD_COL_LV).Value) Then
            taskLevel = CLng(oldSheet.Cells(oldRow, OLD_COL_LV).Value)
        End If
        
        ' タスク名を取得
        taskText = Trim$(CStr(oldSheet.Cells(oldRow, OLD_COL_TASK).Value))
        
        ' タスク名が空の場合はスキップ
        If taskText = "" Then GoTo NextRow
        
        ' A列: LV
        v2Sheet.Cells(v2Row, V2_COL_LV).Value = taskLevel
        
        ' B列: No.
        v2Sheet.Cells(v2Row, V2_COL_NO).Value = taskNo
        taskNo = taskNo + 1
        
        ' C～F列: タスク名（階層に応じて配置）
        Select Case taskLevel
            Case 1
                v2Sheet.Cells(v2Row, V2_COL_TASK_LV1).Value = taskText
            Case 2
                v2Sheet.Cells(v2Row, V2_COL_TASK_LV2).Value = taskText
            Case 3
                v2Sheet.Cells(v2Row, V2_COL_TASK_LV3).Value = taskText
            Case 4
                v2Sheet.Cells(v2Row, V2_COL_TASK_LV4).Value = taskText
            Case Else
                ' LVが指定されていない、または5以上の場合はLV1として扱う
                v2Sheet.Cells(v2Row, V2_COL_TASK_LV1).Value = taskText
        End Select
        
        ' H列: 状況（進捗率から推定）
        progressText = Trim$(CStr(oldSheet.Cells(oldRow, OLD_COL_PROGRESS).Value))
        If progressText <> "" Then
            ' パーセント表記を数値に変換
            progressValue = ParseProgressValue(progressText)
            
            ' I列: 進捗率（0～1の値）
            v2Sheet.Cells(v2Row, V2_COL_PROGRESS).Value = progressValue
            
            ' 状況を自動設定
            If progressValue >= 0.999 Then
                v2Sheet.Cells(v2Row, V2_COL_STATUS).Value = "完了"
            ElseIf progressValue > 0 Then
                v2Sheet.Cells(v2Row, V2_COL_STATUS).Value = "進行中"
            Else
                v2Sheet.Cells(v2Row, V2_COL_STATUS).Value = "未着手"
            End If
        End If
        
        ' J列: 担当
        v2Sheet.Cells(v2Row, V2_COL_ASSIGNEE).Value = oldSheet.Cells(oldRow, OLD_COL_ASSIGNEE).Value
        
        ' K列: 開始予定
        If IsDate(oldSheet.Cells(oldRow, OLD_COL_START).Value) Then
            v2Sheet.Cells(v2Row, V2_COL_START_PLAN).Value = CDate(oldSheet.Cells(oldRow, OLD_COL_START).Value)
        End If
        
        ' L列: 完了予定
        If IsDate(oldSheet.Cells(oldRow, OLD_COL_END).Value) Then
            v2Sheet.Cells(v2Row, V2_COL_END_PLAN).Value = CDate(oldSheet.Cells(oldRow, OLD_COL_END).Value)
        End If
        
        v2Row = v2Row + 1
        
NextRow:
    Next oldRow
    
    Application.ScreenUpdating = True
    
    ' 移管完了メッセージ
    MsgBox "データの移管が完了しました！" & vbCrLf & vbCrLf & _
           "移管元: " & oldSheet.Name & vbCrLf & _
           "移管先: InazumaGantt_v2" & vbCrLf & _
           "移管件数: " & (v2Row - V2_ROW_DATA_START) & " 件" & vbCrLf & vbCrLf & _
           "次の手順:" & vbCrLf & _
           "1. InazumaGantt_v2シートに移動" & vbCrLf & _
           "2. RefreshInazumaGanttマクロを実行" & vbCrLf & _
           "3. HierarchyColor.ApplyHierarchyColorsマクロを実行（任意）", _
           vbInformation, "移管完了"
    
    ' v2シートをアクティブにする
    v2Sheet.Activate
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "移管中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  進捗率の値を解析（文字列 → 0～1の数値）
' ==========================================
Private Function ParseProgressValue(ByVal progressText As String) As Double
    Dim cleanText As String
    Dim numValue As Double
    
    ' 空の場合は0
    If Trim$(progressText) = "" Then
        ParseProgressValue = 0
        Exit Function
    End If
    
    ' "%"記号を除去
    cleanText = Replace(progressText, "%", "")
    cleanText = Trim$(cleanText)
    
    ' 数値に変換
    If IsNumeric(cleanText) Then
        numValue = CDbl(cleanText)
        
        ' 100超の値は100%として扱う
        If numValue > 100 Then numValue = 100
        If numValue < 0 Then numValue = 0
        
        ' 0～1の範囲に正規化
        ParseProgressValue = numValue / 100
    Else
        ' 数値でない場合は0
        ParseProgressValue = 0
    End If
End Function
