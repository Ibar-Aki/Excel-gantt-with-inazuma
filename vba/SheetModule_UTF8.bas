' ==========================================
'  InazumaGantt_v2 繧ｷ繝ｼ繝医Δ繧ｸ繝･繝ｼ繝ｫ逕ｨ繧ｳ繝ｼ繝・
' ==========================================
' 縺薙・繧ｳ繝ｼ繝峨・縲栗nazumaGantt_v2縲阪す繝ｼ繝医・繧ｷ繝ｼ繝医Δ繧ｸ繝･繝ｼ繝ｫ縺ｫ雋ｼ繧贋ｻ倥￠縺ｦ縺上□縺輔＞
'
' 縲占ｨｭ螳壽婿豕輔・
' 1. Excel縺ｧ Alt+F11 繧呈款縺励※VBA繧ｨ繝・ぅ繧ｿ繧帝幕縺・
' 2. 繝励Ο繧ｸ繧ｧ繧ｯ繝医お繧ｯ繧ｹ繝励Ο繝ｼ繝ｩ繝ｼ縺ｧ縲栗nazumaGantt_v2縲阪す繝ｼ繝医ｒ繝繝悶Ν繧ｯ繝ｪ繝・け
' 3. 髢九＞縺溘さ繝ｼ繝峨え繧｣繝ｳ繝峨え縺ｫ莉･荳九・繧ｳ繝ｼ繝峨ｒ雋ｼ繧贋ｻ倥￠繧・
' 4. VBA繧ｨ繝・ぅ繧ｿ繧帝哩縺倥ｋ
'
' ==========================================

' 繝・・繧ｿ髢句ｧ玖｡鯉ｼ・nazumaGantt_v2繝｢繧ｸ繝･繝ｼ繝ｫ縺ｨ蜷梧悄・・
Private Const ROW_DATA_START As Long = 9

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 繧ｿ繧ｹ繧ｯ陦後・繝繝悶Ν繧ｯ繝ｪ繝・け縺ｧ螳御ｺ・・逅・ｒ螳溯｡・
    On Error GoTo ErrorHandler
    
    If Target.Row < ROW_DATA_START Then Exit Sub
    
    ' 騾ｲ謐礼紫繧・00%縺ｫ
    Me.Cells(Target.Row, "I").Value = 1
    
    ' 迥ｶ豕√ｒ縲悟ｮ御ｺ・阪↓
    Me.Cells(Target.Row, "H").Value = "螳御ｺ・
    
    ' 髢句ｧ句ｮ溽ｸｾ縺後≠繧句ｴ蜷医∝ｮ御ｺ・ｮ溽ｸｾ縺ｫ莉頑律繧定ｨｭ螳・
    If IsDate(Me.Cells(Target.Row, "M").Value) Then
        Me.Cells(Target.Row, "N").Value = Date
    End If
    
    Cancel = True
    Exit Sub
    
ErrorHandler:
    ' 繧ｨ繝ｩ繝ｼ縺ｯ辟｡隕・
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    
    ' 繧ｿ繧ｹ繧ｯ蜈･蜉帛・・・・曦蛻暦ｼ峨↓螟画峩縺後≠縺｣縺溷ｴ蜷医・嚴螻､繧定・蜍募愛螳・
    If Not Intersect(Target, Me.Range("C:F")) Is Nothing Then
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Range("C:F"))
            If cell.Row >= ROW_DATA_START Then
                InazumaGantt_v2.AutoDetectTaskLevel cell.Row
            End If
        Next cell
    End If
    
    ' 騾ｲ謐礼紫蛻暦ｼ・蛻暦ｼ峨↓螟画峩縺後≠縺｣縺溷ｴ蜷医∫憾豕√ｒ閾ｪ蜍墓峩譁ｰ
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
End Sub

Private Sub UpdateStatusByProgress(ByVal targetRow As Long)
    Dim progressValue As Variant
    Dim rate As Double
    Dim textValue As String
    
    progressValue = Me.Cells(targetRow, "I").Value
    
    If Trim$(CStr(progressValue)) = "" Then
        Me.Cells(targetRow, "H").Value = "譛ｪ逹謇・
        Exit Sub
    End If
    
    If IsNumeric(progressValue) Then
        rate = CDbl(progressValue)
    Else
        textValue = Replace$(Trim$(CStr(progressValue)), "%", "")
        If Not IsNumeric(textValue) Then
            Exit Sub
        End If
        rate = CDbl(textValue)
    End If
    
    ' 100雜・・蛟､縺ｯ蜑ｲ蜷医→縺励※謇ｱ縺・
    If rate > 1 Then rate = rate / 100
    If rate < 0 Then rate = 0
    If rate > 1 Then rate = 1
    
    ' 迥ｶ豕√ｒ險ｭ螳・
    If rate >= 1 Then
        Me.Cells(targetRow, "H").Value = "螳御ｺ・
    ElseIf rate <= 0 Then
        Me.Cells(targetRow, "H").Value = "譛ｪ逹謇・
    Else
        Me.Cells(targetRow, "H").Value = "騾ｲ陦御ｸｭ"
    End If
End Sub
