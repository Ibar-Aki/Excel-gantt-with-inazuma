Attribute VB_Name = "DataMigration"
Option Explicit

' ==========================================
'  繝・・繧ｿ遘ｻ邂｡繝｢繧ｸ繝･繝ｼ繝ｫ
' ==========================================
' 譌｢蟄倥・繧ｬ繝ｳ繝医メ繝｣繝ｼ繝亥ｽ｢蠑上°繧益2蠖｢蠑上∈繝・・繧ｿ繧堤ｧｻ邂｡縺吶ｋ

' ==========================================
'  v2蠖｢蠑上∈縺ｮ遘ｻ邂｡螳溯｡・
' ==========================================
Sub MigrateToV2Format()
    On Error GoTo ErrorHandler
    
    Dim oldSheet As Worksheet
    Dim newSheet As Worksheet
    
    Set oldSheet = ActiveSheet
    
    ' 遒ｺ隱・
    Dim result As VbMsgBoxResult
    result = MsgBox("縺薙・繧ｷ繝ｼ繝医・繝・・繧ｿ繧致2蠖｢蠑上↓遘ｻ邂｡縺励∪縺吶°・・ & vbCrLf & vbCrLf & _
                   "遘ｻ邂｡蜈・ " & oldSheet.Name & vbCrLf & _
                   "遘ｻ邂｡蜈・ InazumaGantt_v2 繧ｷ繝ｼ繝茨ｼ域眠隕丈ｽ懈・・・, _
                   vbQuestion + vbYesNo, "繝・・繧ｿ遘ｻ邂｡")
    
    If result <> vbYes Then
        MsgBox "遘ｻ邂｡繧偵く繝｣繝ｳ繧ｻ繝ｫ縺励∪縺励◆縲・, vbInformation
        Exit Sub
    End If
    
    ' v2繧ｷ繝ｼ繝医ｒ蜿門ｾ励∪縺溘・菴懈・
    On Error Resume Next
    Set newSheet = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Worksheets.Add(After:=oldSheet)
        newSheet.Name = InazumaGantt_v2.MAIN_SHEET_NAME
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 遘ｻ邂｡蜃ｦ逅・
    Dim oldRow As Long, newRow As Long
    Dim lastOldRow As Long
    
    ' 蜈・ョ繝ｼ繧ｿ縺ｮ譛邨り｡後ｒ蜿門ｾ暦ｼ・蛻怜渕貅厄ｼ・
    lastOldRow = oldSheet.Cells(oldSheet.Rows.Count, "C").End(xlUp).Row
    If lastOldRow < 2 Then lastOldRow = 2
    
    ' 繝・・繧ｿ陦後・髢句ｧ具ｼ・2蠖｢蠑擾ｼ・
    newRow = InazumaGantt_v2.ROW_DATA_START
    
    ' 繝倥ャ繝繝ｼ陦後ｒ繧ｹ繧ｭ繝・・縺励※遘ｻ邂｡
    For oldRow = 2 To lastOldRow
        ' 遨ｺ陦後・繧ｹ繧ｭ繝・・
        If Trim$(CStr(oldSheet.Cells(oldRow, "C").Value)) <> "" Then
            ' 繧ｿ繧ｹ繧ｯ蜷搾ｼ・蛻暦ｼ・
            newSheet.Cells(newRow, "C").Value = oldSheet.Cells(oldRow, "C").Value
            
            ' 蜿ｯ閭ｽ縺ｪ蛻励ｒ繝槭ャ繝斐Φ繧ｰ
            If oldSheet.Cells(1, "D").Value Like "*隧ｳ邏ｰ*" Or oldSheet.Cells(1, "D").Value Like "*蜀・ｮｹ*" Then
                newSheet.Cells(newRow, "G").Value = oldSheet.Cells(oldRow, "D").Value
            End If
            
            ' 譌･莉伜・縺ｮ繝槭ャ繝斐Φ繧ｰ
            MapDateColumns oldSheet, newSheet, oldRow, newRow
            
            ' 騾ｲ謐礼紫縺ｮ繝槭ャ繝斐Φ繧ｰ
            MapProgressColumn oldSheet, newSheet, oldRow, newRow
            
            ' 諡・ｽ楢・・繝槭ャ繝斐Φ繧ｰ
            MapAssigneeColumn oldSheet, newSheet, oldRow, newRow
            
            newRow = newRow + 1
        End If
    Next oldRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 髫主ｱ､閾ｪ蜍募愛螳・
    newSheet.Activate
    InazumaGantt_v2.AutoDetectTaskLevel
    
    MsgBox "遘ｻ邂｡螳御ｺ・ｼ・ & vbCrLf & vbCrLf & _
           "遘ｻ邂｡蜈・ " & oldSheet.Name & vbCrLf & _
           "遘ｻ邂｡蜈・ " & newSheet.Name & vbCrLf & _
           "遘ｻ邂｡陦梧焚: " & (newRow - InazumaGantt_v2.ROW_DATA_START), _
           vbInformation, "繝・・繧ｿ遘ｻ邂｡"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "遘ｻ邂｡繧ｨ繝ｩ繝ｼ: " & Err.Description, vbCritical, "繧ｨ繝ｩ繝ｼ"
End Sub

' ==========================================
'  譌･莉伜・縺ｮ繝槭ャ繝斐Φ繧ｰ
' ==========================================
Private Sub MapDateColumns(ByVal oldSheet As Worksheet, ByVal newSheet As Worksheet, ByVal oldRow As Long, ByVal newRow As Long)
    Dim col As Long
    
    For col = 1 To oldSheet.Cells(1, oldSheet.Columns.Count).End(xlToLeft).Column
        Dim header As String
        header = CStr(oldSheet.Cells(1, col).Value)
        
        If header Like "*髢句ｧ倶ｺ亥ｮ・" Or header Like "*Start*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "K").Value = oldSheet.Cells(oldRow, col).Value
            End If
        ElseIf header Like "*螳御ｺ・ｺ亥ｮ・" Or header Like "*End*" Or header Like "*邨ゆｺ・ｺ亥ｮ・" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "L").Value = oldSheet.Cells(oldRow, col).Value
            End If
        ElseIf header Like "*髢句ｧ句ｮ溽ｸｾ*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "M").Value = oldSheet.Cells(oldRow, col).Value
            End If
        ElseIf header Like "*螳御ｺ・ｮ溽ｸｾ*" Then
            If IsDate(oldSheet.Cells(oldRow, col).Value) Then
                newSheet.Cells(newRow, "N").Value = oldSheet.Cells(oldRow, col).Value
            End If
        End If
    Next col
End Sub

' ==========================================
'  騾ｲ謐礼紫縺ｮ繝槭ャ繝斐Φ繧ｰ
' ==========================================
Private Sub MapProgressColumn(ByVal oldSheet As Worksheet, ByVal newSheet As Worksheet, ByVal oldRow As Long, ByVal newRow As Long)
    Dim col As Long
    
    For col = 1 To oldSheet.Cells(1, oldSheet.Columns.Count).End(xlToLeft).Column
        Dim header As String
        header = CStr(oldSheet.Cells(1, col).Value)
        
        If header Like "*騾ｲ謐・" Or header Like "*Progress*" Then
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
'  諡・ｽ楢・・繝槭ャ繝斐Φ繧ｰ
' ==========================================
Private Sub MapAssigneeColumn(ByVal oldSheet As Worksheet, ByVal newSheet As Worksheet, ByVal oldRow As Long, ByVal newRow As Long)
    Dim col As Long
    
    For col = 1 To oldSheet.Cells(1, oldSheet.Columns.Count).End(xlToLeft).Column
        Dim header As String
        header = CStr(oldSheet.Cells(1, col).Value)
        
        If header Like "*諡・ｽ・" Or header Like "*Assignee*" Then
            newSheet.Cells(newRow, "J").Value = oldSheet.Cells(oldRow, col).Value
            Exit For
        End If
    Next col
End Sub
