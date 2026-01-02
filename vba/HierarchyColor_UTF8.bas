Attribute VB_Name = "HierarchyColor"
Option Explicit

' ==========================================
'  階層色分けモジュール
' ==========================================

' 階層別の色定義
Public Const COLOR_LV1 As Long = 14083324  ' RGB(252,228,214) サーモン
Public Const COLOR_LV2 As Long = 15983322  ' RGB(218,227,243) 薄い青
Public Const COLOR_LV3 As Long = 14348514  ' RGB(226,239,218) 薄い緑
Public Const COLOR_LV4 As Long = 13434879  ' RGB(255,242,204) 薄い黄色

' 色塗り開始列（No.列）
Public Const COL_COLOR_START As String = "B"
' 色塗り終了列（ガント開始列の手前）
Public Const COL_COLOR_END As String = "N"

' ==========================================
'  階層色分けを適用
' ==========================================
Sub ApplyHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lastRow As Long
    lastRow = InazumaGantt_v2.ROW_DATA_START
    
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "C").End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    lastRow = MaxRow(lastRow, ws.Cells(ws.Rows.Count, "F").End(xlUp).Row)
    
    If lastRow < InazumaGantt_v2.ROW_DATA_START Then
        lastRow = InazumaGantt_v2.ROW_DATA_START + 199
    End If
    
    ' 既存の塗りをクリア
    ws.Range(COL_COLOR_START & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow).Interior.ColorIndex = xlNone
    
    Dim r As Long
    Dim lvValue As Variant
    Dim taskLevel As Long
    Dim taskCol As String
    Dim colorValue As Long
    Dim colorRange As Range
    
    For r = InazumaGantt_v2.ROW_DATA_START To lastRow
        lvValue = ws.Cells(r, InazumaGantt_v2.COL_HIERARCHY).Value
        
        If IsNumeric(lvValue) And lvValue <> "" Then
            taskLevel = CLng(lvValue)
            
            Select Case taskLevel
                Case 1
                    colorValue = COLOR_LV1
                    taskCol = "C"
                Case 2
                    colorValue = COLOR_LV2
                    taskCol = "D"
                Case 3
                    colorValue = COLOR_LV3
                    taskCol = "E"
                Case 4
                    colorValue = COLOR_LV4
                    taskCol = "F"
                Case Else
                    colorValue = 0
            End Select
            
            If colorValue > 0 Then
                Set colorRange = ws.Range(taskCol & r & ":" & COL_COLOR_END & r)
                colorRange.Interior.Color = colorValue
            End If
        End If
    Next r
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "階層色分けを適用しました！", vbInformation, "階層色分け"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "色分けエラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  色分けをクリア
' ==========================================
Sub ClearHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < InazumaGantt_v2.ROW_DATA_START Then lastRow = InazumaGantt_v2.ROW_DATA_START + 199
    
    ws.Range(COL_COLOR_START & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow).Interior.ColorIndex = xlNone
    
    MsgBox "色分けをクリアしました！", vbInformation, "階層色分け"
    Exit Sub
    
ErrorHandler:
    MsgBox "クリアエラー: " & Err.Description, vbCritical, "エラー"
End Sub

Private Function MaxRow(ByVal a As Long, ByVal b As Long) As Long
    If b > a Then
        MaxRow = b
    Else
        MaxRow = a
    End If
End Function
