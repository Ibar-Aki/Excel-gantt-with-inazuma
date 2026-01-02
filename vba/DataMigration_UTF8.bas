Attribute VB_Name = "DataMigration"
Option Explicit

' ==========================================
'  データ移管モジュール
' ==========================================
' 既存のガントチャート形式からv2形式へデータを移管する

' ==========================================
'  v2形式への移管実行
' ==========================================
Sub MigrateToV2Format()
    On Error GoTo ErrorHandler
    
    Dim oldSheet As Worksheet
    Dim newSheet As Worksheet
    
    Set oldSheet = ActiveSheet
    
    ' 確認
    Dim result As VbMsgBoxResult
    result = MsgBox("このシートのデータをv2形式に移管しますか？" & vbCrLf & vbCrLf & _
                   "移管元: " & oldSheet.Name & vbCrLf & _
                   "移管先: InazumaGantt_v2 シート（新規作成）", _
                   vbQuestion + vbYesNo, "データ移管")
    
    If result <> vbYes Then
        MsgBox "移管をキャンセルしました。", vbInformation
        Exit Sub
    End If
    
    ' v2シートを取得または作成
    On Error Resume Next
    Set newSheet = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Worksheets.Add(After:=oldSheet)
        newSheet.Name = InazumaGantt_v2.MAIN_SHEET_NAME
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 移管処理
    Dim oldRow As Long, newRow As Long
    Dim lastOldRow As Long
    
    ' 元データの最終行を取得（C列基準）
    lastOldRow = oldSheet.Cells(oldSheet.Rows.Count, "C").End(xlUp).Row
    If lastOldRow < 2 Then lastOldRow = 2
    
    ' データ行の開始（v2形式）
    newRow = InazumaGantt_v2.ROW_DATA_START
    
    ' ヘッダー行をスキップして移管
    For oldRow = 2 To lastOldRow
        ' 空行はスキップ
        If Trim$(CStr(oldSheet.Cells(oldRow, "C").Value)) <> "" Then
            ' タスク名（C列）
            newSheet.Cells(newRow, "C").Value = oldSheet.Cells(oldRow, "C").Value
            
            ' 可能な列をマッピング
            If oldSheet.Cells(1, "D").Value Like "*詳細*" Or oldSheet.Cells(1, "D").Value Like "*内容*" Then
                newSheet.Cells(newRow, "G").Value = oldSheet.Cells(oldRow, "D").Value
            End If
            
            ' 日付列のマッピング
            MapDateColumns oldSheet, newSheet, oldRow, newRow
            
            ' 進捗率のマッピング
            MapProgressColumn oldSheet, newSheet, oldRow, newRow
            
            ' 担当者のマッピング
            MapAssigneeColumn oldSheet, newSheet, oldRow, newRow
            
            newRow = newRow + 1
        End If
    Next oldRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 階層自動判定
    newSheet.Activate
    InazumaGantt_v2.AutoDetectTaskLevel
    
    MsgBox "移管完了！" & vbCrLf & vbCrLf & _
           "移管元: " & oldSheet.Name & vbCrLf & _
           "移管先: " & newSheet.Name & vbCrLf & _
           "移管行数: " & (newRow - InazumaGantt_v2.ROW_DATA_START), _
           vbInformation, "データ移管"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "移管エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  日付列のマッピング
' ==========================================
Private Sub MapDateColumns(ByVal oldSheet As Worksheet, ByVal newSheet As Worksheet, ByVal oldRow As Long, ByVal newRow As Long)
    Dim col As Long
    
    For col = 1 To oldSheet.Cells(1, oldSheet.Columns.Count).End(xlToLeft).Column
        Dim header As String
        header = CStr(oldSheet.Cells(1, col).Value)
        
        If header Like "*開始予定*" Or header Like "*Start*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "K").Value = oldSheet.Cells(oldRow, col).Value
            End If
        ElseIf header Like "*完了予定*" Or header Like "*End*" Or header Like "*終了予定*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "L").Value = oldSheet.Cells(oldRow, col).Value
            End If
        ElseIf header Like "*開始実績*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "M").Value = oldSheet.Cells(oldRow, col).Value
            End If
        ElseIf header Like "*完了実績*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "N").Value = oldSheet.Cells(oldRow, col).Value
            End If
        End If
    Next col
End Sub

' ==========================================
'  進捗率のマッピング
' ==========================================
Private Sub MapProgressColumn(ByVal oldSheet As Worksheet, ByVal newSheet As Worksheet, ByVal oldRow As Long, ByVal newRow As Long)
    Dim col As Long
    
    For col = 1 To oldSheet.Cells(1, oldSheet.Columns.Count).End(xlToLeft).Column
        Dim header As String
        header = CStr(oldSheet.Cells(1, col).Value)
        
        If header Like "*進捗*" Or header Like "*Progress*" Then
            Dim progressValue As Variant
            progressValue = oldSheet.Cells(oldRow, col).Value
            
            If IsNumeric(progressValue) Then
                Dim rate As Double
                rate = CDbl(progressValue)
                If rate > 1 Then rate = rate / 100
                newSheet.Cells(newRow, "I").Value = rate
            End If
            Exit For
        End If
    Next col
End Sub

' ==========================================
'  担当者のマッピング
' ==========================================
Private Sub MapAssigneeColumn(ByVal oldSheet As Worksheet, ByVal newSheet As Worksheet, ByVal oldRow As Long, ByVal newRow As Long)
    Dim col As Long
    
    For col = 1 To oldSheet.Cells(1, oldSheet.Columns.Count).End(xlToLeft).Column
        Dim header As String
        header = CStr(oldSheet.Cells(1, col).Value)
        
        If header Like "*担当*" Or header Like "*Assignee*" Then
            newSheet.Cells(newRow, "J").Value = oldSheet.Cells(oldRow, col).Value
            Exit For
        End If
    Next col
End Sub
