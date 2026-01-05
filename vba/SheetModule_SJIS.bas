' ==========================================
'  InazumaGantt_v2 シートモジュール用コード
' ==========================================
' このコードは「InazumaGantt_v2」シートのシートモジュールに貼り付けてください
'
' 【設定方法】
' 1. Excelで Alt+F11 を押してVBAエディタを開く
' 2. プロジェクトエクスプローラーで「InazumaGantt_v2」シートをダブルクリック
' 3. 開いたコードウィンドウに以下のコードを貼り付ける
' 4. VBAエディタを閉じる
'
' ==========================================

' データ開始行（InazumaGantt_v2モジュールと同期）
Private Const ROW_DATA_START As Long = 9

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' タスク行のダブルクリックで完了処理を実行
    ' ※ No.列(B)またはLV列(A)のみ有効（誤操作防止）
    On Error GoTo ErrorHandler
    
    If Target.Row < ROW_DATA_START Then Exit Sub
    
    ' A列(1) or B列(2) のみ対象
    If Target.Column <> 1 And Target.Column <> 2 Then Exit Sub
    
    ' 既に完了済みの場合は変更しない（誤操作防止）
    If Me.Cells(Target.Row, "H").Value = "完了" Then Exit Sub
    
    ' 進捗率を100%に
    Me.Cells(Target.Row, "I").Value = 1
    
    ' 状況を「完了」に
    Me.Cells(Target.Row, "H").Value = "完了"
    
    ' 開始実績がある場合、完了実績に今日を設定（空の場合のみ）
    If IsDate(Me.Cells(Target.Row, "M").Value) Then
        If Trim(CStr(Me.Cells(Target.Row, "N").Value)) = "" Then
            Me.Cells(Target.Row, "N").Value = Date
        End If
    End If
    
    Cancel = True
    Exit Sub
    
ErrorHandler:
    ' エラーは無視
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Application.EnableEvents = False
    
    ' タスク入力列（C～F列）に変更があった場合
    If Not Intersect(Target, Me.Range("C:F")) Is Nothing Then
        Dim cell As Range
        For Each cell In Intersect(Target, Me.Range("C:F"))
            If cell.Row >= ROW_DATA_START Then
                ' タスクが入力された場合
                If Trim$(CStr(cell.Value)) <> "" Then
                    ' 階層を自動判定
                    InazumaGantt_v2.AutoDetectTaskLevel cell.Row
                    
                    ' No.が空なら自動入力
                    If Trim$(CStr(Me.Cells(cell.Row, "B").Value)) = "" Then
                        Me.Cells(cell.Row, "B").Value = GetNextNo()
                    End If
                    
                    ' 進捗率が空なら0%を入力
                    If Trim$(CStr(Me.Cells(cell.Row, "I").Value)) = "" Then
                        Me.Cells(cell.Row, "I").Value = 0
                    End If
                    
                    ' 状況が空なら「未着手」を入力
                    If Trim$(CStr(Me.Cells(cell.Row, "H").Value)) = "" Then
                        Me.Cells(cell.Row, "H").Value = "未着手"
                    End If
                Else
                    ' タスクが削除された場合も階層を更新
                    InazumaGantt_v2.AutoDetectTaskLevel cell.Row
                End If
            End If
        Next cell
    End If
    
    ' 進捗率列（I列）に変更があった場合、状況を自動更新
    If Not Intersect(Target, Me.Columns("I")) Is Nothing Then
        Dim progressCell As Range
        For Each progressCell In Intersect(Target, Me.Columns("I"))
            If progressCell.Row >= ROW_DATA_START Then
                UpdateStatusByProgress progressCell.Row
            End If
        Next progressCell
    End If
    
    ' 予定日付列（K, L列）に土日祝日を入力した場合に確認メッセージ
    If Not Intersect(Target, Me.Range("K:L")) Is Nothing Then
        Dim dateCell As Range
        Dim inputDate As Date
        Dim isWeekend As Boolean
        Dim isHoliday As Boolean
        Dim warningMsg As String
        
        For Each dateCell In Intersect(Target, Me.Range("K:L"))
            If dateCell.Row >= ROW_DATA_START Then
                If IsDate(dateCell.Value) Then
                    inputDate = CDate(dateCell.Value)
                    isWeekend = (Weekday(inputDate, vbMonday) >= 6)
                    isHoliday = CheckHoliday(inputDate)
                    
                    If isWeekend Or isHoliday Then
                        If isHoliday Then
                            warningMsg = "祝日"
                        ElseIf Weekday(inputDate, vbMonday) = 6 Then
                            warningMsg = "土曜日"
                        Else
                            warningMsg = "日曜日"
                        End If
                        
                        If MsgBox(Format(inputDate, "yy/mm/dd") & " は " & warningMsg & " です。" & vbCrLf & _
                                  "この日付を入力しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then
                            Application.EnableEvents = False
                            dateCell.ClearContents
                            Application.EnableEvents = True
                        End If
                    End If
                End If
            End If
        Next dateCell
    End If
    
    Application.EnableEvents = True
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
End Sub

' ==========================================
'  祝日チェック
' ==========================================
Private Function CheckHoliday(ByVal targetDate As Date) As Boolean
    Dim wsHoliday As Worksheet
    On Error Resume Next
    Set wsHoliday = ThisWorkbook.Worksheets("祝日マスタ")
    On Error GoTo 0
    
    CheckHoliday = False
    If wsHoliday Is Nothing Then Exit Function
    
    Dim lastRow As Long
    lastRow = wsHoliday.Cells(wsHoliday.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    Dim r As Long
    For r = 2 To lastRow
        If IsDate(wsHoliday.Cells(r, "A").Value) Then
            If CDate(wsHoliday.Cells(r, "A").Value) = targetDate Then
                CheckHoliday = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Sub UpdateStatusByProgress(ByVal targetRow As Long)
    Dim progressValue As Variant
    Dim rate As Double
    Dim textValue As String
    
    progressValue = Me.Cells(targetRow, "I").Value
    
    If Trim$(CStr(progressValue)) = "" Then
        Me.Cells(targetRow, "H").Value = "未着手"
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
    
    ' 100超の値は割合として扱う
    If rate > 1 Then rate = rate / 100
    If rate < 0 Then rate = 0
    If rate > 1 Then rate = 1
    
    ' 状況を設定
    If rate >= 1 Then
        Me.Cells(targetRow, "H").Value = "完了"
    ElseIf rate <= 0 Then
        Me.Cells(targetRow, "H").Value = "未着手"
    Else
        Me.Cells(targetRow, "H").Value = "進行中"
    End If
End Sub

' ==========================================
'  次のNo.を取得
' ==========================================
Private Function GetNextNo() As Long
    Dim lastNo As Long
    Dim r As Long
    Dim cellValue As Variant
    
    lastNo = 0
    
    ' B列から最大のNo.を探す
    For r = ROW_DATA_START To Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
        cellValue = Me.Cells(r, "B").Value
        If IsNumeric(cellValue) Then
            If CLng(cellValue) > lastNo Then
                lastNo = CLng(cellValue)
            End If
        End If
    Next r
    
    GetNextNo = lastNo + 1
End Function
