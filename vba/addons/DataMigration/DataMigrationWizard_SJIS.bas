Attribute VB_Name = "DataMigrationWizard"
Option Explicit

' ==========================================
'  データ移管ウィザード - メインモジュール
' ==========================================

' マッピング設定の型定義
Public Type MappingConfig
    TemplateName As String           ' 設定名（"Backlogからの移管"など）
    SourceSheetName As String        ' 移管元シート名
    WBSColumn As String              ' WBS番号列（例: "A"）
    TaskNameColumn As String         ' タスク名列（例: "B"）
    AssigneeColumn As String         ' 担当者列（例: "C"）
    EndPlanColumn As String          ' 完了予定列（例: "D"）
    ProgressColumn As String         ' 進捗率列（任意、例: "E"）
    StartPlanColumn As String        ' 開始予定列（任意）
    StartActualColumn As String      ' 開始実績列（任意）
    EndActualColumn As String        ' 完了実績列（任意）
    DataStartRow As Long             ' データ開始行（例: 2）
    HierarchyMode As Long            ' 階層判定モード (0: WBS, 1: Level)
End Type

' ==========================================
'  ウィザード起動（エントリーポイント）
' ==========================================
Public Sub ShowMigrationWizard()
    On Error GoTo ErrorHandler
    
    ' UserFormの存在チェック
    Dim formExists As Boolean
    formExists = False
    
    On Error Resume Next
    Dim testForm As Object
    Set testForm = VBA.UserForms.Add("frmMigrationWizard")
    If Not testForm Is Nothing Then
        formExists = True
        Unload testForm
    End If
    On Error GoTo ErrorHandler
    
    If Not formExists Then
        MsgBox "ウィザードフォームが見つかりません。" & vbCrLf & vbCrLf & _
               "最初に CreateMigrationWizardForm() を実行して" & vbCrLf & _
               "フォームを作成してください。", vbExclamation, "フォーム未作成"
        Exit Sub
    End If
    
    ' ウィザードフォームを表示
    frmMigrationWizard.Show
    Exit Sub
    
ErrorHandler:
    MsgBox "ウィザード起動エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  移管実行
' ==========================================
Public Sub ExecuteMigration(ByRef config As MappingConfig)
    On Error GoTo ErrorHandler
    
    Dim oldSheet As Worksheet
    Dim newSheet As Worksheet
    
    ' 移管元シートを取得
    On Error Resume Next
    Set oldSheet = ThisWorkbook.Worksheets(config.SourceSheetName)
    On Error GoTo ErrorHandler
    
    If oldSheet Is Nothing Then
        MsgBox "移管元シート '" & config.SourceSheetName & "' が見つかりません。", vbCritical, "エラー"
        Exit Sub
    End If
    
    ' 移管先シートを取得または作成
    On Error Resume Next
    Set newSheet = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        newSheet.Name = InazumaGantt_v2.MAIN_SHEET_NAME
    Else
        ' 既存シートが存在する場合は確認
        Dim result As VbMsgBoxResult
        result = MsgBox("'" & InazumaGantt_v2.MAIN_SHEET_NAME & "' シートが既に存在します。" & vbCrLf & _
                       "データを追加しますか？" & vbCrLf & vbCrLf & _
                       "はい: 既存データの下に追加" & vbCrLf & _
                       "いいえ: キャンセル", vbQuestion + vbYesNo, "既存シート確認")
        If result <> vbYes Then
            Exit Sub
        End If
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 移管処理
    Dim oldRow As Long, newRow As Long
    Dim lastOldRow As Long
    Dim wbsCol As Long, taskCol As Long
    
    ' 列番号を取得
    wbsCol = Range(config.WBSColumn & "1").Column
    taskCol = Range(config.TaskNameColumn & "1").Column
    
    ' 元データの最終行を取得
    lastOldRow = oldSheet.Cells(oldSheet.Rows.Count, wbsCol).End(xlUp).Row
    If lastOldRow < config.DataStartRow Then lastOldRow = config.DataStartRow
    
    ' 移管先の開始行を取得（既存データがある場合は最終行の次）
    newRow = newSheet.Cells(newSheet.Rows.Count, "C").End(xlUp).Row + 1
    If newRow < InazumaGantt_v2.ROW_DATA_START Then
        newRow = InazumaGantt_v2.ROW_DATA_START
    End If
    
    Dim migratedCount As Long
    migratedCount = 0
    
    ' データ行を移管
    For oldRow = config.DataStartRow To lastOldRow
        ' 階層判定用の値を取得（旧WBS列）
        Dim hierarchyText As String
        hierarchyText = Trim$(CStr(oldSheet.Cells(oldRow, wbsCol).Value))
        
        ' 値が空の行はスキップ
        If hierarchyText <> "" Then
            ' 階層レベルを判定
            Dim level As Long
            level = WBSParser.ParseHierarchyLevel(hierarchyText, config.HierarchyMode)
            
            ' 無効なレベルはスキップ
            If level >= 1 And level <= 4 Then
                ' A列に階層レベルを設定
                newSheet.Cells(newRow, "A").Value = level
                
                ' タスク名をC?F列に配置（階層に応じて列を変える）
                Dim taskName As String
                taskName = Trim$(CStr(oldSheet.Cells(oldRow, taskCol).Value))
                
                Select Case level
                    Case 1
                        newSheet.Cells(newRow, "C").Value = taskName
                    Case 2
                        newSheet.Cells(newRow, "D").Value = taskName
                    Case 3
                        newSheet.Cells(newRow, "E").Value = taskName
                    Case 4
                        newSheet.Cells(newRow, "F").Value = taskName
                End Select
                
                ' 担当者のマッピング
                If config.AssigneeColumn <> "" Then
                    Dim assigneeCol As Long
                    assigneeCol = Range(config.AssigneeColumn & "1").Column
                    newSheet.Cells(newRow, "J").Value = oldSheet.Cells(oldRow, assigneeCol).Value
                End If
                
                ' 完了予定日のマッピング
                If config.EndPlanColumn <> "" Then
                    Dim endPlanCol As Long
                    endPlanCol = Range(config.EndPlanColumn & "1").Column
                    If IsDate(oldSheet.Cells(oldRow, endPlanCol).Value) Then
                        newSheet.Cells(newRow, "L").Value = oldSheet.Cells(oldRow, endPlanCol).Value
                    End If
                End If
                
                ' 進捗率のマッピング
                If config.ProgressColumn <> "" Then
                    Dim progressCol As Long
                    progressCol = Range(config.ProgressColumn & "1").Column
                    Dim progressValue As Variant
                    progressValue = oldSheet.Cells(oldRow, progressCol).Value
                    
                    If IsNumeric(progressValue) Then
                        Dim rate As Double
                        rate = CDbl(progressValue)
                        If rate > 1 Then rate = rate / 100
                        newSheet.Cells(newRow, "I").Value = rate
                    End If
                End If
                
                ' 開始予定日のマッピング
                If config.StartPlanColumn <> "" Then
                    Dim startPlanCol As Long
                    startPlanCol = Range(config.StartPlanColumn & "1").Column
                    If IsDate(oldSheet.Cells(oldRow, startPlanCol).Value) Then
                        newSheet.Cells(newRow, "K").Value = oldSheet.Cells(oldRow, startPlanCol).Value
                    End If
                End If
                
                ' 開始実績日のマッピング
                If config.StartActualColumn <> "" Then
                    Dim startActualCol As Long
                    startActualCol = Range(config.StartActualColumn & "1").Column
                    If IsDate(oldSheet.Cells(oldRow, startActualCol).Value) Then
                        newSheet.Cells(newRow, "M").Value = oldSheet.Cells(oldRow, startActualCol).Value
                    End If
                End If
                
                ' 完了実績日のマッピング
                If config.EndActualColumn <> "" Then
                    Dim endActualCol As Long
                    endActualCol = Range(config.EndActualColumn & "1").Column
                    If IsDate(oldSheet.Cells(oldRow, endActualCol).Value) Then
                        newSheet.Cells(newRow, "N").Value = oldSheet.Cells(oldRow, endActualCol).Value
                    End If
                End If
                
                migratedCount = migratedCount + 1
                newRow = newRow + 1
            End If
        End If
    Next oldRow
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' 移管完了メッセージ
    newSheet.Activate
    MsgBox "移管完了！" & vbCrLf & vbCrLf & _
           "移管元: " & config.SourceSheetName & vbCrLf & _
           "移管先: " & InazumaGantt_v2.MAIN_SHEET_NAME & vbCrLf & _
           "移管行数: " & migratedCount, vbInformation, "データ移管"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "移管エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  マッピング設定の保存（Excelシート）
' ==========================================
Public Sub SaveMappingConfig(ByRef config As MappingConfig)
    On Error GoTo ErrorHandler
    
    ' 設定マスタシートを取得または作成
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.SETTINGS_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = InazumaGantt_v2.SETTINGS_SHEET_NAME
        
        ' ヘッダーを作成
        ws.Range("A1").Value = "設定名"
        ws.Range("B1").Value = "移管元シート"
        ws.Range("C1").Value = "WBS列"
        ws.Range("D1").Value = "タスク名列"
        ws.Range("E1").Value = "担当者列"
        ws.Range("F1").Value = "完了予定列"
        ws.Range("G1").Value = "進捗率列"
        ws.Range("H1").Value = "開始予定列"
        ws.Range("I1").Value = "開始実績列"
        ws.Range("J1").Value = "完了実績列"
        ws.Range("K1").Value = "データ開始行"
        ws.Range("L1").Value = "階層モード"
        
        ws.Range("A1:L1").Font.Bold = True
        ws.Range("A1:L1").Interior.Color = RGB(48, 84, 150)
        ws.Range("A1:L1").Font.Color = RGB(255, 255, 255)
    End If
    
    ' 最終行を取得
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If lastRow < 2 Then lastRow = 2
    
    ' 設定を保存
    ws.Cells(lastRow, 1).Value = config.TemplateName
    ws.Cells(lastRow, 2).Value = config.SourceSheetName
    ws.Cells(lastRow, 3).Value = config.WBSColumn
    ws.Cells(lastRow, 4).Value = config.TaskNameColumn
    ws.Cells(lastRow, 5).Value = config.AssigneeColumn
    ws.Cells(lastRow, 6).Value = config.EndPlanColumn
    ws.Cells(lastRow, 7).Value = config.ProgressColumn
    ws.Cells(lastRow, 8).Value = config.StartPlanColumn
    ws.Cells(lastRow, 9).Value = config.StartActualColumn
    ws.Cells(lastRow, 10).Value = config.EndActualColumn
    ws.Cells(lastRow, 11).Value = config.DataStartRow
    ws.Cells(lastRow, 12).Value = config.HierarchyMode
    
    MsgBox "設定を保存しました: " & config.TemplateName, vbInformation, "設定保存"
    Exit Sub
    
ErrorHandler:
    MsgBox "設定保存エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  マッピング設定の読み込み（Excelシート）
' ==========================================
Public Function LoadMappingConfig(ByVal templateName As String) As MappingConfig
    On Error GoTo ErrorHandler
    
    Dim config As MappingConfig
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.SETTINGS_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "設定マスタシートが見つかりません。", vbExclamation, "エラー"
        Exit Function
    End If
    
    ' 設定名で検索
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lastRow
        If ws.Cells(r, 1).Value = templateName Then
            ' 設定を読み込み
            config.TemplateName = ws.Cells(r, 1).Value
            config.SourceSheetName = ws.Cells(r, 2).Value
            config.WBSColumn = ws.Cells(r, 3).Value
            config.TaskNameColumn = ws.Cells(r, 4).Value
            config.AssigneeColumn = ws.Cells(r, 5).Value
            config.EndPlanColumn = ws.Cells(r, 6).Value
            config.ProgressColumn = ws.Cells(r, 7).Value
            config.StartPlanColumn = ws.Cells(r, 8).Value
            config.StartActualColumn = ws.Cells(r, 9).Value
            config.EndActualColumn = ws.Cells(r, 10).Value
            config.DataStartRow = ws.Cells(r, 11).Value
            If ws.Cells(r, 12).Value <> "" Then
                config.HierarchyMode = ws.Cells(r, 12).Value
            Else
                config.HierarchyMode = 0 ' Default to WBS
            End If
            
            LoadMappingConfig = config
            Exit Function
        End If
    Next r
    
    MsgBox "設定 '" & templateName & "' が見つかりません。", vbExclamation, "エラー"
    Exit Function
    
ErrorHandler:
    MsgBox "設定読込エラー: " & Err.Description, vbCritical, "エラー"
End Function

' ==========================================
'  保存済み設定の一覧を取得
' ==========================================
Public Function GetSavedTemplates() As Variant
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.SETTINGS_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        GetSavedTemplates = Array()
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow < 2 Then
        GetSavedTemplates = Array()
        Exit Function
    End If
    
    ' 設定名の一覧を作成
    Dim templates() As String
    ReDim templates(1 To lastRow - 1)
    
    Dim r As Long, idx As Long
    idx = 1
    For r = 2 To lastRow
        templates(idx) = ws.Cells(r, 1).Value
        idx = idx + 1
    Next r
    
    GetSavedTemplates = templates
    Exit Function
    
ErrorHandler:
    GetSavedTemplates = Array()
End Function
