Attribute VB_Name = "HierarchyColor"
Option Explicit

' ==========================================
'  髫主ｱ､濶ｲ蛻・￠繝｢繧ｸ繝･繝ｼ繝ｫ
' ==========================================

' 髫主ｱ､蛻･縺ｮ濶ｲ螳夂ｾｩ
Public Const COLOR_LV1 As Long = 252& + 228& * 256& + 214& * 65536&  ' RGB(252,228,214) 繧ｵ繝ｼ繝｢繝ｳ
Public Const COLOR_LV2 As Long = 218& + 227& * 256& + 243& * 65536&  ' RGB(218,227,243) 阮・＞髱・
Public Const COLOR_LV3 As Long = 226& + 239& * 256& + 218& * 65536&  ' RGB(226,239,218) 阮・＞邱・
Public Const COLOR_LV4 As Long = 255& + 242& * 256& + 204& * 65536&  ' RGB(255,242,204) 阮・＞鮟・牡

' 濶ｲ蝪励ｊ髢句ｧ句・・・o.蛻暦ｼ・
Public Const COL_COLOR_START As String = "B"
' 濶ｲ蝪励ｊ邨ゆｺ・・・医ぎ繝ｳ繝磯幕蟋句・縺ｮ謇句燕・・
Public Const COL_COLOR_END As String = "N"

' ==========================================
'  髫主ｱ､濶ｲ蛻・￠繧帝←逕ｨ
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
    
    ' 譌｢蟄倥・蝪励ｊ繧偵け繝ｪ繧｢
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
    
    MsgBox "髫主ｱ､濶ｲ蛻・￠繧帝←逕ｨ縺励∪縺励◆・・, vbInformation, "髫主ｱ､濶ｲ蛻・￠"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "濶ｲ蛻・￠繧ｨ繝ｩ繝ｼ: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub

' ==========================================
'  濶ｲ蛻・￠繧偵け繝ｪ繧｢
' ==========================================
Sub ClearHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    If lastRow < InazumaGantt_v2.ROW_DATA_START Then lastRow = InazumaGantt_v2.ROW_DATA_START + 199
    
    ws.Range(COL_COLOR_START & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow).Interior.ColorIndex = xlNone
    
    MsgBox "濶ｲ蛻・￠繧偵け繝ｪ繧｢縺励∪縺励◆・・, vbInformation, "髫主ｱ､濶ｲ蛻・￠"
    Exit Sub
    
ErrorHandler:
    MsgBox "繧ｯ繝ｪ繧｢繧ｨ繝ｩ繝ｼ: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub

Private Function MaxRow(ByVal a As Long, ByVal b As Long) As Long
    If b > a Then
        MaxRow = b
    Else
        MaxRow = a
    End If
End Function
