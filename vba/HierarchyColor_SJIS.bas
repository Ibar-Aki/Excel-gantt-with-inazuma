Attribute VB_Name = "HierarchyColor"
Option Explicit

' ==========================================
'  階層色分けモジュール（条件付き書式版）
' ==========================================
' このモジュールは条件付き書式を設定します。
' 一度実行すれば、以降は自動的に色分けが適用されます。
' 
' 塗り範囲ルール:
'   LV1 (A列=1): C～N列を塗る
'   LV2 (A列=2): D～N列を塗る
'   LV3 (A列=3): E～N列を塗る
'   LV4 (A列=4): F～N列を塗る

' 階層別の色定義
Public Const COLOR_LV1 As Long = 14083324  ' RGB(252,228,214) サーモン
Public Const COLOR_LV2 As Long = 15983322  ' RGB(218,227,243) 薄い青
Public Const COLOR_LV3 As Long = 14348514  ' RGB(226,239,218) 薄い緑
Public Const COLOR_LV4 As Long = 13434879  ' RGB(255,242,204) 薄い黄色

' 色塗り終了列（ガント開始列の手前）
Public Const COL_COLOR_END As String = "N"

' ==========================================
'  階層色分けの条件付き書式を設定
' ==========================================
Sub SetupHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim lastRow As Long
    lastRow = InazumaGantt_v2.ROW_DATA_START + InazumaGantt_v2.DATA_ROWS_DEFAULT - 1
    
    ' 既存の条件付き書式をクリア（B～N列）
    ws.Range("B" & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow).FormatConditions.Delete
    
    ' LV1: A列が1のとき、C～N列をサーモン色に
    Dim rangeLV1 As Range
    Set rangeLV1 = ws.Range("C" & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow)
    Dim cf1 As FormatCondition
    Set cf1 = rangeLV1.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$A" & InazumaGantt_v2.ROW_DATA_START & "=1")
    cf1.Interior.Color = COLOR_LV1
    cf1.StopIfTrue = True
    
    ' LV2: A列が2のとき、D～N列を薄い青に
    Dim rangeLV2 As Range
    Set rangeLV2 = ws.Range("D" & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow)
    Dim cf2 As FormatCondition
    Set cf2 = rangeLV2.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$A" & InazumaGantt_v2.ROW_DATA_START & "=2")
    cf2.Interior.Color = COLOR_LV2
    cf2.StopIfTrue = True
    
    ' LV3: A列が3のとき、E～N列を薄い緑に
    Dim rangeLV3 As Range
    Set rangeLV3 = ws.Range("E" & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow)
    Dim cf3 As FormatCondition
    Set cf3 = rangeLV3.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$A" & InazumaGantt_v2.ROW_DATA_START & "=3")
    cf3.Interior.Color = COLOR_LV3
    cf3.StopIfTrue = True
    
    ' LV4: A列が4のとき、F～N列を薄い黄色に
    Dim rangeLV4 As Range
    Set rangeLV4 = ws.Range("F" & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow)
    Dim cf4 As FormatCondition
    Set cf4 = rangeLV4.FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$A" & InazumaGantt_v2.ROW_DATA_START & "=4")
    cf4.Interior.Color = COLOR_LV4
    cf4.StopIfTrue = True
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    If Application.DisplayAlerts Then
        MsgBox "階層色分けの条件付き書式を設定しました！" & vbCrLf & vbCrLf & _
               "塗り範囲ルール:" & vbCrLf & _
               "  LV1: C～N列" & vbCrLf & _
               "  LV2: D～N列" & vbCrLf & _
               "  LV3: E～N列" & vbCrLf & _
               "  LV4: F～N列", vbInformation, "階層色分け"
    End If
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "条件付き書式設定エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  階層色分けの条件付き書式をクリア
' ==========================================
Sub ClearHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = InazumaGantt_v2.ROW_DATA_START + InazumaGantt_v2.DATA_ROWS_DEFAULT - 1
    
    ' 対象範囲の条件付き書式をクリア（B～N列）
    ws.Range("B" & InazumaGantt_v2.ROW_DATA_START & ":" & COL_COLOR_END & lastRow).FormatConditions.Delete
    
    MsgBox "階層色分けの条件付き書式をクリアしました！", vbInformation, "階層色分け"
    Exit Sub
    
ErrorHandler:
    MsgBox "クリアエラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  旧互換: ApplyHierarchyColors
' ==========================================
' 旧バージョンとの互換性のため、SetupHierarchyColorsを呼び出す
Sub ApplyHierarchyColors()
    Call SetupHierarchyColors
End Sub
