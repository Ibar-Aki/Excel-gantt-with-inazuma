Attribute VB_Name = "GanttBooster"
Option Explicit

' ==========================================
'  ガントチャート最強化キット - 設定エリア
' ==========================================
' お使いのExcelシートの列に合わせて、ここのアルファベットを変更してください
' Vertex42 Simple Gantt Template 標準設定:
' タスク名: B, 開始日: D, 終了日: E, 進捗%: F
Public Const COL_TASK_NAME As String = "B"
Public Const COL_START_DATE As String = "D"
Public Const COL_END_DATE As String = "E"
Public Const COL_PROGRESS As String = "F"

' このマクロで追加・管理する列の設定
Public Const COL_STATUS As String = "A"      ' ステータス信号機 (空いている列を指定)
Public Const COL_CHECK As String = "G"       ' エラーチェック用 (進捗の隣などを指定)
Public Const ROW_HEADER As Long = 5          ' ヘッダー(項目名)がある行番号

' ==========================================
'  セットアップ実行マクロ
' ==========================================
Sub SetupGanttSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 1. ステータス列のロジック追加
    ' 数式: 100%なら"完了"、期限切れなら"遅延"、あと7日以内なら"注意"
    With ws.Range(COL_STATUS & (ROW_HEADER + 1) & ":" & COL_STATUS & "100")
        .Formula = "=IF(" & COL_PROGRESS & (ROW_HEADER + 1) & "=1,""完了"",IF(AND(" & COL_END_DATE & (ROW_HEADER + 1) & "<TODAY()," & COL_PROGRESS & (ROW_HEADER + 1) & "<1),""遅延"",IF(" & COL_END_DATE & (ROW_HEADER + 1) & "-TODAY()<7,""注意"",""順調"")))"
        .FormatConditions.Delete
        ' 遅延 (赤)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""遅延"""
        .FormatConditions(1).Interior.Color = RGB(255, 200, 200) ' 薄い赤
        .FormatConditions(1).Font.Color = RGB(200, 0, 0)
        ' 注意 (黄)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""注意"""
        .FormatConditions(2).Interior.Color = RGB(255, 255, 200) ' 薄い黄色
        .FormatConditions(2).Font.Color = RGB(200, 150, 0)
        ' 完了 (グレー)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""完了"""
        .FormatConditions(3).Interior.Color = RGB(220, 220, 220) ' グレー
        .FormatConditions(3).Font.Color = RGB(100, 100, 100)
    End With
    
    ' 2. 進捗%の入力規則 (0, 25, 50, 75, 100 のプルダウン)
    With ws.Range(COL_PROGRESS & (ROW_HEADER + 1) & ":" & COL_PROGRESS & "100").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="0%,25%,50%,75%,100%"
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
    ' 3. 完了行のグレーアウト (条件付き書式)
    ' タスク名から進捗列までをグレーにし、取り消し線を引く
    Dim rngTable As Range
    Set rngTable = ws.Range(COL_TASK_NAME & (ROW_HEADER + 1) & ":" & COL_PROGRESS & "100")
    rngTable.FormatConditions.Add Type:=xlExpression, Formula1:="=$" & COL_PROGRESS & (ROW_HEADER + 1) & "=1"
    rngTable.FormatConditions(rngTable.FormatConditions.Count).Interior.Color = RGB(240, 240, 240)
    rngTable.FormatConditions(rngTable.FormatConditions.Count).Font.Color = RGB(160, 160, 160)
    rngTable.FormatConditions(rngTable.FormatConditions.Count).Font.StrikeThrough = True

    MsgBox "セットアップ完了！" & vbCrLf & "ステータス列の追加と、条件付き書式の設定を行いました。", vbInformation, "Gantt Booster"
End Sub

' ==========================================
'  一発完了ボタン用マクロ
' ==========================================
Sub SetComplete()
    Dim cell As Range
    ' 選択されている行に対して実行
    For Each cell In Selection.Rows
        Dim r As Long
        r = cell.Row
        ' データ行かどうかチェック
        If r > ROW_HEADER Then
            ActiveSheet.Cells(r, Columns(COL_PROGRESS).Column).Value = 1 ' 100% に設定
        End If
    Next cell
End Sub

' ==========================================
'  フィルタ機能
' ==========================================
Sub FilterDelayed()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    
    ' ステータス列(A列想定)で「遅延」のみ表示
    ' ※注意: ステータス列がA列以外の場合は Field:=1 の数字を調整する必要があります
    ws.Range(COL_STATUS & ROW_HEADER & ":" & COL_PROGRESS & "100").AutoFilter Field:=1, Criteria1:="遅延"
End Sub

Sub FilterMyTasks()
    ' ※「担当者」列がどこにあるかVBA側で特定できないため、この機能は設定が必要です。
    MsgBox "この機能を使うには、VBAコード内の担当者列設定が必要です。", vbInformation
End Sub

Sub ClearFilters()
    If ActiveSheet.AutoFilterMode Then ActiveSheet.ShowAllData
End Sub

' ==========================================
'  表示モード切替
' ==========================================
Sub ToggleWeeklyView()
    ' 「開始日」列を隠して、スッキリさせるスイッチ
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws.Columns(COL_START_DATE).Hidden Then
        ws.Columns(COL_START_DATE).Hidden = False
    Else
        ws.Columns(COL_START_DATE).Hidden = True
    End If
End Sub
