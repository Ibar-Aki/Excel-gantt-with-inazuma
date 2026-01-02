' ==========================================
'  InazumaGantt_v2 シートモジュール（レガシー版）
' ==========================================
' このコードは旧バージョンです。
' 新しいバージョンは vba/SheetModule_SJIS.bas を使用してください。
' ==========================================

Private Const ROW_DATA_START As Long = 9

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    On Error GoTo ErrorHandler
    
    If Target.Row < ROW_DATA_START Then Exit Sub
    
    Me.Cells(Target.Row, "I").Value = 1
    Me.Cells(Target.Row, "H").Value = "完了"
    
    If IsDate(Me.Cells(Target.Row, "M").Value) Then
        Me.Cells(Target.Row, "N").Value = Date
    End If
    
    Cancel = True
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    Application.EnableEvents = False
    
    If Not Intersect(Target, Me.Range("C:F")) Is Nothing Then
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Range("C:F"))
            If cell.Row >= ROW_DATA_START Then
                InazumaGantt_v2.AutoDetectTaskLevel cell.Row
            End If
        Next cell
    End If
    
    If Not Intersect(Target, Me.Columns("I")) Is Nothing Then
        Dim progressCell As Range
        For Each progressCell In Intersect(Target, Me.Columns("I"))
            If progressCell.Row >= ROW_DATA_START Then
                UpdateStatusByProgress progressCell.Row
            End If
        Next progressCell
    End If
    
    Application.EnableEvents = True
End Sub

Private Sub UpdateStatusByProgress(ByVal targetRow As Long)
    Dim progressValue As Variant
    Dim rate As Double
    
    progressValue = Me.Cells(targetRow, "I").Value
    
    If Trim$(CStr(progressValue)) = "" Then
        Me.Cells(targetRow, "H").Value = "未着手"
        Exit Sub
    End If
    
    If IsNumeric(progressValue) Then
        rate = CDbl(progressValue)
    Else
        Exit Sub
    End If
    
    If rate > 1 Then rate = rate / 100
    
    If rate >= 1 Then
        Me.Cells(targetRow, "H").Value = "完了"
    ElseIf rate <= 0 Then
        Me.Cells(targetRow, "H").Value = "未着手"
    Else
        Me.Cells(targetRow, "H").Value = "進行中"
    End If
End Sub
