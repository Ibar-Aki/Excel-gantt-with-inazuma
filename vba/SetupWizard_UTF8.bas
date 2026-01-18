Attribute VB_Name = "SetupWizard"
Option Explicit

' ==========================================
'  セットアップウィザードモジュール
' ==========================================
' 対話形式でセットアップを進めるウィザード機能
' ==========================================

' ==========================================
'  サイレントセットアップ（自動テスト用）
' ==========================================
' MsgBoxなしで自動実行。PowerShell等からの呼び出し用。
' 引数: addSampleData - サンプルデータを追加するか
Public Sub SilentSetup(Optional ByVal isAddSampleData As Boolean = True)
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 開始日を計算（14日前の範囲内で最も近い月曜日）
    Dim startDate As Date
    startDate = Date - 14
    ' 月曜日に調整（Weekday: 1=日, 2=月, ..., 7=土）
    Dim dayOffset As Long
    dayOffset = Weekday(startDate, vbMonday) - 1 ' 月曜からのオフセット
    startDate = startDate - dayOffset
    
    ' シート作成（サイレントモード・開始日指定）
    CreateMainSheetSilent Format(startDate, "yy/mm/dd")
    
    ' サンプルデータ追加
    If isAddSampleData Then
        ' startDateを基準にサンプルデータを追加
        AddSampleData startDate
    End If
    
    ' 設定マスタシート作成
    InazumaGantt_v2.EnsureSettingsSheet
    
    ' メインシートをアクティブに（重要：設定マスタではなくメインシートで描画）
    ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME).Activate
    
    ' 階層色分け設定
    HierarchyColor.SetupHierarchyColors
    
    ' ガントチャート描画
    InazumaGantt_v2.RefreshInazumaGantt
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Err.Raise Err.Number, "SilentSetup", Err.Description
End Sub

' ==========================================
'  ウィザード実行
' ==========================================
Sub RunSetupWizard()
    On Error GoTo ErrorHandler
    
    Dim result As VbMsgBoxResult
    
    ' ステップ1: 開始確認
    result = MsgBox("InazumaGantt セットアップウィザードへようこそ！" & vbCrLf & vbCrLf & _
                   "このウィザードでは以下を設定します:" & vbCrLf & _
                   "1. メインシートの作成" & vbCrLf & _
                   "2. 設定マスタ（祝日欄含む）の作成" & vbCrLf & _
                   "3. サンプルデータの追加（任意）" & vbCrLf & vbCrLf & _
                   "続行しますか？", _
                   vbQuestion + vbYesNo, "セットアップウィザード")
    
    If result <> vbYes Then
        MsgBox "セットアップをキャンセルしました。", vbInformation
        Exit Sub
    End If
    
    ' ステップ2: シート作成確認
    result = MsgBox("新しいシート「InazumaGantt_v2」を作成しますか？" & vbCrLf & vbCrLf & _
                   "注意: 同名のシートが既に存在する場合は上書きされません。", _
                   vbQuestion + vbYesNo, "ステップ 1/3: シート作成")
    
    If result = vbYes Then
        CreateMainSheet
    End If
    
    ' ステップ3: サンプルデータ
    result = MsgBox("サンプルデータを追加しますか？" & vbCrLf & vbCrLf & _
                   "サンプルデータには以下が含まれます:" & vbCrLf & _
                   "- 3つのフェーズ（LV1）" & vbCrLf & _
                   "- 各フェーズに2-3個のタスク（LV2-LV3）", _
                   vbQuestion + vbYesNo, "ステップ 2/3: サンプルデータ")
    
    If result = vbYes Then
        AddSampleData
    End If
    
    ' ステップ4: 階層色分けとガント描画を自動実行
    ' まずメインシートをアクティブにする
    ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME).Activate
    
    Application.ScreenUpdating = False
    
    ' v2.2: 設定マスタシートを作成
    InazumaGantt_v2.EnsureSettingsSheet
    
    ' 階層色分けの条件付き書式を設定
    HierarchyColor.SetupHierarchyColors
    
    ' ガントチャートを描画
    InazumaGantt_v2.RefreshInazumaGantt
    
    Application.ScreenUpdating = True
    
    ' ステップ5: 完了
    MsgBox "セットアップウィザードが完了しました！" & vbCrLf & vbCrLf & _
           "以下の設定が完了しました:" & vbCrLf & _
           "- シート作成（メイン、祝日マスタ、設定マスタ）" & vbCrLf & _
           "- 階層色分け（条件付き書式）" & vbCrLf & _
           "- ガントチャート描画" & vbCrLf & vbCrLf & _
           "【シートモジュールの設定】" & vbCrLf & _
           "ダブルクリック完了・折りたたみ機能を使うには、" & vbCrLf & _
           "SheetModule_SJIS.bas をシートモジュールに貼り付けてください。", _
           vbInformation, "セットアップ完了"
    Exit Sub
    
ErrorHandler:
    MsgBox "セットアップ中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub


' ==========================================
'  メインシートの作成
' ==========================================
Private Sub CreateMainSheet()
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = InazumaGantt_v2.MAIN_SHEET_NAME
    End If
    
    ws.Activate
    InazumaGantt_v2.SetupInazumaGantt False, Null
End Sub

' ==========================================
'  メインシートの作成（サイレント版）
' ==========================================
Private Sub CreateMainSheetSilent(ByVal startDateStr As String)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = InazumaGantt_v2.MAIN_SHEET_NAME
    End If
    
    ws.Activate
    InazumaGantt_v2.SetupInazumaGantt True, startDateStr
End Sub

' ==========================================
'  サンプルデータの追加
' ==========================================
' ==========================================
'  サンプルデータの追加（統合版）
' ==========================================
Private Sub AddSampleData(Optional ByVal baseDate As Date = 0)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    
    Dim startRow As Long
    startRow = InazumaGantt_v2.ROW_DATA_START
    
    ' 日付指定がない場合は今日を基準
    If baseDate = 0 Then baseDate = Date
    
    ' フェーズ1: 計画フェーズ（完了フェーズ）
    ws.Cells(startRow, "C").Value = "計画フェーズ"
    ws.Cells(startRow, "H").Value = "完了"
    ws.Cells(startRow, "I").Value = 1
    ws.Cells(startRow, "J").Value = "山田"
    ws.Cells(startRow, "K").Value = GetWorkday(baseDate - 14)
    ws.Cells(startRow, "L").Value = GetWorkday(baseDate - 7)
    ws.Cells(startRow, "M").Value = GetWorkday(baseDate - 14)
    ws.Cells(startRow, "N").Value = GetWorkday(baseDate - 8)
    
    ws.Cells(startRow + 1, "D").Value = "要件定義"
    ws.Cells(startRow + 1, "H").Value = "完了"
    ws.Cells(startRow + 1, "I").Value = 1
    ws.Cells(startRow + 1, "J").Value = "山田"
    ws.Cells(startRow + 1, "K").Value = GetWorkday(baseDate - 14)
    ws.Cells(startRow + 1, "L").Value = GetWorkday(baseDate - 10)
    
    ws.Cells(startRow + 2, "D").Value = "設計書作成"
    ws.Cells(startRow + 2, "H").Value = "完了"
    ws.Cells(startRow + 2, "I").Value = 1
    ws.Cells(startRow + 2, "J").Value = "鈴木"
    ws.Cells(startRow + 2, "K").Value = GetWorkday(baseDate - 10)
    ws.Cells(startRow + 2, "L").Value = GetWorkday(baseDate - 7)
    
    ' フェーズ2: 開発フェーズ（進行中）
    ws.Cells(startRow + 3, "C").Value = "開発フェーズ"
    ws.Cells(startRow + 3, "H").Value = "進行中"
    ws.Cells(startRow + 3, "I").Value = 0.6
    ws.Cells(startRow + 3, "J").Value = "田中"
    ws.Cells(startRow + 3, "K").Value = GetWorkday(baseDate - 7)
    ws.Cells(startRow + 3, "L").Value = GetWorkday(baseDate + 14)
    
    ws.Cells(startRow + 4, "D").Value = "機能開発"
    ws.Cells(startRow + 4, "H").Value = "進行中"
    ws.Cells(startRow + 4, "I").Value = 0.7
    ws.Cells(startRow + 4, "J").Value = "田中"
    ws.Cells(startRow + 4, "K").Value = GetWorkday(baseDate - 7)
    ws.Cells(startRow + 4, "L").Value = GetWorkday(baseDate + 7)
    
    ' フェーズ3: リリースフェーズ（未着手）
    ws.Cells(startRow + 5, "C").Value = "リリースフェーズ"
    ws.Cells(startRow + 5, "H").Value = "未着手"
    ws.Cells(startRow + 5, "I").Value = 0
    ws.Cells(startRow + 5, "J").Value = "山田"
    ws.Cells(startRow + 5, "K").Value = GetWorkday(baseDate + 14)
    ws.Cells(startRow + 5, "L").Value = GetWorkday(baseDate + 21)
    
    ' 階層自動判定
    InazumaGantt_v2.AutoDetectTaskLevel
    Exit Sub
    
ErrorHandler:
    MsgBox "サンプルデータ追加エラー: " & Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  平日（土日を避けた日付）を取得
' ==========================================
Private Function GetWorkday(ByVal targetDate As Date) As Date
    ' 土曜の場合は前の金曜に
    ' 日曜の場合は次の月曜に
    Dim dow As Long
    dow = Weekday(targetDate, vbSunday) ' 1=日, 2=月, ..., 7=土
    
    If dow = 1 Then ' 日曜
        GetWorkday = targetDate + 1 ' 月曜に
    ElseIf dow = 7 Then ' 土曜
        GetWorkday = targetDate - 1 ' 金曜に
    Else
        GetWorkday = targetDate
    End If
End Function

' ==========================================
'  シートモジュール設定手順の表示
' ==========================================
Sub ShowSheetModuleInstructions()
    Dim instructions As String
    
    instructions = "【シートモジュールの設定手順】" & vbCrLf & vbCrLf & _
                  "1. Alt + F11 でVBAエディタを開く" & vbCrLf & _
                  "2. プロジェクトエクスプローラーで" & vbCrLf & _
                  "   「InazumaGantt_v2」シートをダブルクリック" & vbCrLf & _
                  "3. vba/SheetModule_SJIS.bas の内容を" & vbCrLf & _
                  "   コピー＆貼り付け" & vbCrLf & _
                  "4. 保存して閉じる" & vbCrLf & vbCrLf & _
                  "これにより以下の機能が有効になります:" & vbCrLf & _
                  "- タスク入力時の階層自動判定" & vbCrLf & _
                  "- 進捗率変更時の状況自動更新" & vbCrLf & _
                  "- ダブルクリックでタスク完了"
    
    MsgBox instructions, vbInformation, "シートモジュール設定"
End Sub

' ==========================================
'  モジュール存在確認
' ==========================================
Public Function IsModuleInstalled(ByVal moduleName As String) As Boolean
    On Error Resume Next
    Dim vbComp As Object
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If StrComp(vbComp.Name, moduleName, vbTextCompare) = 0 Then
            IsModuleInstalled = True
            Exit Function
        End If
    Next vbComp
    
    IsModuleInstalled = False
End Function

' ==========================================
'  インストール状態の確認
' ==========================================
Sub CheckInstallation()
    Dim status As String
    
    status = "【モジュールインストール状態】" & vbCrLf & vbCrLf
    
    ' 必須モジュール
    status = status & "必須モジュール:" & vbCrLf
    status = status & "  InazumaGantt_v2: " & IIf(IsModuleInstalled("InazumaGantt_v2"), "OK", "未インストール") & vbCrLf
    status = status & "  HierarchyColor: " & IIf(IsModuleInstalled("HierarchyColor"), "OK", "未インストール") & vbCrLf
    
    ' オプションモジュール
    status = status & vbCrLf & "オプションモジュール:" & vbCrLf
    status = status & "  DataMigration: " & IIf(IsModuleInstalled("DataMigration"), "OK", "未インストール") & vbCrLf
    status = status & "  ErrorHandler: " & IIf(IsModuleInstalled("ErrorHandler"), "OK", "未インストール") & vbCrLf
    
    MsgBox status, vbInformation, "インストール状態"
End Sub
