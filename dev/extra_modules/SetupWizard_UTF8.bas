Attribute VB_Name = "SetupWizard"
Option Explicit

' ==========================================
'  セットアップウィザードモジュール
' ==========================================
' 対話形式でセットアップを進めるウィザード機能
' ==========================================

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
                   "2. 祝日マスタシートの作成" & vbCrLf & _
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
    Application.ScreenUpdating = False
    
    ' 階層色分けの条件付き書式を設定
    HierarchyColor.SetupHierarchyColors
    
    ' ガントチャートを描画
    InazumaGantt_v2.RefreshInazumaGantt
    
    Application.ScreenUpdating = True
    
    ' ステップ5: 完了
    MsgBox "セットアップウィザードが完了しました！" & vbCrLf & vbCrLf & _
           "以下の設定が完了しました:" & vbCrLf & _
           "- シート作成" & vbCrLf & _
           "- 階層色分け（条件付き書式）" & vbCrLf & _
           "- ガントチャート描画" & vbCrLf & vbCrLf & _
           "タスクを入力して RefreshInazumaGantt を実行してください。", _
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
    InazumaGantt_v2.SetupInazumaGantt
End Sub

' ==========================================
'  サンプルデータの追加
' ==========================================
Private Sub AddSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    ' ActiveSheetではなく、明示的にメインシートを指定
    Set ws = ThisWorkbook.Worksheets(InazumaGantt_v2.MAIN_SHEET_NAME)
    
    Dim startRow As Long
    startRow = InazumaGantt_v2.ROW_DATA_START
    
    Dim baseDate As Date
    baseDate = Date  ' 今日を基準にする
    
    ' フェーズ1: 計画フェーズ
    ws.Cells(startRow, "C").Value = "計画フェーズ"
    ws.Cells(startRow, "H").Value = "完了"
    ws.Cells(startRow, "I").Value = 1
    ws.Cells(startRow, "J").Value = "山田"
    ws.Cells(startRow, "K").Value = baseDate - 14
    ws.Cells(startRow, "L").Value = baseDate - 7
    ws.Cells(startRow, "M").Value = baseDate - 14
    ws.Cells(startRow, "N").Value = baseDate - 8
    
    ws.Cells(startRow + 1, "D").Value = "要件定義"
    ws.Cells(startRow + 1, "H").Value = "完了"
    ws.Cells(startRow + 1, "I").Value = 1
    ws.Cells(startRow + 1, "J").Value = "山田"
    ws.Cells(startRow + 1, "K").Value = baseDate - 14
    ws.Cells(startRow + 1, "L").Value = baseDate - 10
    ws.Cells(startRow + 1, "M").Value = baseDate - 14
    ws.Cells(startRow + 1, "N").Value = baseDate - 10
    
    ws.Cells(startRow + 2, "D").Value = "設計書作成"
    ws.Cells(startRow + 2, "H").Value = "完了"
    ws.Cells(startRow + 2, "I").Value = 1
    ws.Cells(startRow + 2, "J").Value = "鈴木"
    ws.Cells(startRow + 2, "K").Value = baseDate - 10
    ws.Cells(startRow + 2, "L").Value = baseDate - 7
    ws.Cells(startRow + 2, "M").Value = baseDate - 10
    ws.Cells(startRow + 2, "N").Value = baseDate - 8
    
    ' フェーズ2: 開発フェーズ
    ws.Cells(startRow + 3, "C").Value = "開発フェーズ"
    ws.Cells(startRow + 3, "H").Value = "進行中"
    ws.Cells(startRow + 3, "I").Value = 0.6
    ws.Cells(startRow + 3, "J").Value = "田中"
    ws.Cells(startRow + 3, "K").Value = baseDate - 7
    ws.Cells(startRow + 3, "L").Value = baseDate + 14
    ws.Cells(startRow + 3, "M").Value = baseDate - 7
    
    ws.Cells(startRow + 4, "D").Value = "機能開発"
    ws.Cells(startRow + 4, "H").Value = "進行中"
    ws.Cells(startRow + 4, "I").Value = 0.7
    ws.Cells(startRow + 4, "J").Value = "田中"
    ws.Cells(startRow + 4, "K").Value = baseDate - 7
    ws.Cells(startRow + 4, "L").Value = baseDate + 7
    ws.Cells(startRow + 4, "M").Value = baseDate - 7
    
    ws.Cells(startRow + 5, "E").Value = "機能A開発"
    ws.Cells(startRow + 5, "H").Value = "完了"
    ws.Cells(startRow + 5, "I").Value = 1
    ws.Cells(startRow + 5, "J").Value = "田中"
    ws.Cells(startRow + 5, "K").Value = baseDate - 7
    ws.Cells(startRow + 5, "L").Value = baseDate - 3
    ws.Cells(startRow + 5, "M").Value = baseDate - 7
    ws.Cells(startRow + 5, "N").Value = baseDate - 2
    
    ws.Cells(startRow + 6, "E").Value = "機能B開発"
    ws.Cells(startRow + 6, "H").Value = "進行中"
    ws.Cells(startRow + 6, "I").Value = 0.5
    ws.Cells(startRow + 6, "J").Value = "佐藤"
    ws.Cells(startRow + 6, "K").Value = baseDate - 3
    ws.Cells(startRow + 6, "L").Value = baseDate + 4
    ws.Cells(startRow + 6, "M").Value = baseDate - 3
    
    ws.Cells(startRow + 7, "D").Value = "テスト"
    ws.Cells(startRow + 7, "H").Value = "未着手"
    ws.Cells(startRow + 7, "I").Value = 0
    ws.Cells(startRow + 7, "J").Value = "鈴木"
    ws.Cells(startRow + 7, "K").Value = baseDate + 5
    ws.Cells(startRow + 7, "L").Value = baseDate + 14
    
    ' フェーズ3: リリースフェーズ
    ws.Cells(startRow + 8, "C").Value = "リリースフェーズ"
    ws.Cells(startRow + 8, "H").Value = "未着手"
    ws.Cells(startRow + 8, "I").Value = 0
    ws.Cells(startRow + 8, "J").Value = "山田"
    ws.Cells(startRow + 8, "K").Value = baseDate + 14
    ws.Cells(startRow + 8, "L").Value = baseDate + 21
    
    ws.Cells(startRow + 9, "D").Value = "本番環境構築"
    ws.Cells(startRow + 9, "H").Value = "未着手"
    ws.Cells(startRow + 9, "I").Value = 0
    ws.Cells(startRow + 9, "J").Value = "佐藤"
    ws.Cells(startRow + 9, "K").Value = baseDate + 14
    ws.Cells(startRow + 9, "L").Value = baseDate + 18
    
    ws.Cells(startRow + 10, "D").Value = "リリース作業"
    ws.Cells(startRow + 10, "H").Value = "未着手"
    ws.Cells(startRow + 10, "I").Value = 0
    ws.Cells(startRow + 10, "J").Value = "山田"
    ws.Cells(startRow + 10, "K").Value = baseDate + 19
    ws.Cells(startRow + 10, "L").Value = baseDate + 21
    
    ' 階層自動判定
    InazumaGantt_v2.AutoDetectTaskLevel
    
    Exit Sub
    
ErrorHandler:
    MsgBox "サンプルデータ追加エラー: " & Err.Description, vbCritical, "エラー"
End Sub

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
