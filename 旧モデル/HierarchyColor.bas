Attribute VB_Name = "HierarchyColor"
Option Explicit

' ==========================================
'  階層別色分けモジュール
'  InazumaGantt と連携して使用
' ==========================================

' 階層レベルに応じた薄い色の定義
' LV1: 薄いオレンジ/サーモン
Public Const COLOR_LV1 As Long = 252& + 228& * 256& + 214& * 65536&  ' RGB(252,228,214)
' LV2: 薄い青
Public Const COLOR_LV2 As Long = 221& + 235& * 256& + 247& * 65536&  ' RGB(221,235,247)
' LV3: 薄い緑
Public Const COLOR_LV3 As Long = 226& + 239& * 256& + 218& * 65536&  ' RGB(226,239,218)
' LV4: 薄い黄色
Public Const COLOR_LV4 As Long = 255& + 249& * 256& + 219& * 65536&  ' RGB(255,249,219)
' LV5以上: 薄い紫
Public Const COLOR_LV5 As Long = 242& + 230& * 256& + 255& * 65536&  ' RGB(242,230,255)

' 階層列（InazumaGantt_v2の設定に合わせる）
Public Const COL_HIERARCHY As String = "A"
' 色塗り開始列（No.列）
Public Const COL_COLOR_START As String = "B"
' 色塗り終了列（ガント開始列の手前）
Public Const COL_COLOR_END As String = "N"
' データ開始行（InazumaGantt_v2の設定に合わせる）
Public Const ROW_DATA_START As Long = 9

' ==========================================
'  階層レベルから色を取得
' ==========================================
Private Function GetHierarchyColor(ByVal level As Long) As Long
    Select Case level
        Case 1
            GetHierarchyColor = COLOR_LV1
        Case 2
            GetHierarchyColor = COLOR_LV2
        Case 3
            GetHierarchyColor = COLOR_LV3
        Case 4
            GetHierarchyColor = COLOR_LV4
        Case Else
            If level >= 5 Then
                GetHierarchyColor = COLOR_LV5
            Else
                GetHierarchyColor = -1  ' 色なし
            End If
    End Select
End Function

' ==========================================
'  データ最終行の取得
' ==========================================
Private Function GetLastDataRow(ByVal ws As Worksheet) As Long
    Dim lastRow As Long
    lastRow = ROW_DATA_START
    
    ' タスク列で最終行を確認
    Dim taskLastRow As Long
    taskLastRow = ws.Cells(ws.Rows.Count, COL_COLOR_START).End(xlUp).Row
    If taskLastRow > lastRow Then lastRow = taskLastRow
    
    ' 階層列で最終行を確認
    Dim hierLastRow As Long
    hierLastRow = ws.Cells(ws.Rows.Count, COL_HIERARCHY).End(xlUp).Row
    If hierLastRow > lastRow Then lastRow = hierLastRow
    
    GetLastDataRow = lastRow
End Function

' ==========================================
'  階層別の色塗りを適用（タスク入力列からN列まで）
' ==========================================
' C列→LV1, D列→LV2, E列→LV3, F列→LV4
' タスクが入力された列からN列まで色塗り
Sub ApplyHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Application.ScreenUpdating = False
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    ' 全範囲の色をクリア
    Dim fullRange As Range
    Set fullRange = ws.Range("C" & ROW_DATA_START & ":N" & lastRow)
    fullRange.Interior.ColorIndex = xlNone
    
    Dim r As Long
    Dim hierarchyLevel As Variant
    Dim colorStartCol As String
    Dim colorEndCol As String
    Dim rowColor As Long
    
    For r = ROW_DATA_START To lastRow
        ' 階層レベルを取得
        hierarchyLevel = ws.Cells(r, COL_HIERARCHY).Value
        
        If IsNumeric(hierarchyLevel) And hierarchyLevel >= 1 And hierarchyLevel <= 4 Then
            ' 階層レベルに応じた色を取得
            rowColor = GetHierarchyColor(CLng(hierarchyLevel))
            
            ' タスク入力列を取得（InazumaGantt_v2モジュールの関数を使用）
            On Error Resume Next
            colorStartCol = InazumaGantt_v2.GetTaskColumnByLevel(CLng(hierarchyLevel))
            On Error GoTo ErrorHandler
            
            If colorStartCol <> "" Then
                ' タスク入力列からN列まで色塗り
                colorEndCol = "N"
                Dim colorRange As Range
                Set colorRange = ws.Range(colorStartCol & r & ":" & colorEndCol & r)
                colorRange.Interior.Color = rowColor
            End If
        End If
    Next r
    
    Application.ScreenUpdating = True
    
    MsgBox "階層別の色塗りを適用しました！" & vbCrLf & vbCrLf & _
           "【色の意味】" & vbCrLf & _
           "・LV1(C列) = サーモン" & vbCrLf & _
           "・LV2(D列) = 薄い青" & vbCrLf & _
           "・LV3(E列) = 薄い緑" & vbCrLf & _
           "・LV4(F列) = 薄い黄色", vbInformation, "階層色分け"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  階層別の色塗りをクリア
' ==========================================
Sub ClearHierarchyColors()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lastRow As Long
    lastRow = GetLastDataRow(ws)
    If lastRow < ROW_DATA_START Then lastRow = ROW_DATA_START
    
    ' 色塗り対象範囲
    Dim colorRange As Range
    Set colorRange = ws.Range(COL_COLOR_START & ROW_DATA_START & ":" & COL_COLOR_END & lastRow)
    
    ' 条件付き書式をクリア
    colorRange.FormatConditions.Delete
    ' 背景色もクリア
    colorRange.Interior.ColorIndex = xlNone
    
    MsgBox "階層別の色塗りをクリアしました。", vbInformation, "階層色分け"
    Exit Sub
    
ErrorHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  セル値変更時に自動で色塗りを更新（Worksheet_Changeイベント用）
'  ※このコードはシートモジュールにコピーして使用
' ==========================================
' Private Sub Worksheet_Change(ByVal Target As Range)
'     ' 階層列が変更された場合のみ処理
'     If Not Intersect(Target, Me.Columns("B")) Is Nothing Then
'         Application.EnableEvents = False
'         HierarchyColor.ApplyHierarchyColors
'         Application.EnableEvents = True
'     End If
' End Sub
