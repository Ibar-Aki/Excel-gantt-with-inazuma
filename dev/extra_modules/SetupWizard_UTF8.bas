Attribute VB_Name = "SetupWizard"
Option Explicit

' ==========================================
'  InazumaGantt v2 セットアップウィザード
' ==========================================
' このマクロを実行すると、対話的にセットアップを進められます
' 
' 使い方:
'   Alt + F8 → RunSetupWizard → 実行
' ==========================================

Private Const VERSION As String = "2.0.0"

' ==========================================
'  メインセットアップウィザード
' ==========================================
Sub RunSetupWizard()
    On Error GoTo ErrorHandler
    
    Dim result As VbMsgBoxResult
    
    ' ウェルカムメッセージ
    result = MsgBox("InazumaGantt v" & VERSION & " セットアップウィザードへようこそ！" & vbCrLf & vbCrLf & _
                    "このウィザードでは以下の設定を行います：" & vbCrLf & _
                    "1. シート作成" & vbCrLf & _
                    "2. 初期設定" & vbCrLf & _
                    "3. サンプルデータ追加（任意）" & vbCrLf & vbCrLf & _
                    "セットアップを開始しますか？", _
                    vbYesNo + vbQuestion, "セットアップウィザード")
    
    If result = vbNo Then Exit Sub
    
    ' ステップ1: 必須モジュールチェック
    If Not CheckRequiredModules() Then
        MsgBox "必須モジュールがインポートされていません。" & vbCrLf & vbCrLf & _
               "セットアップガイド（SETUP.md）を参照して、" & vbCrLf & _
               "以下のモジュールをインポートしてください：" & vbCrLf & vbCrLf & _
               "・InazumaGantt_v2_SJIS.bas" & vbCrLf & _
               "・HierarchyColor_SJIS.bas" & vbCrLf & vbCrLf & _
               "インポート後、再度このウィザードを実行してください。", _
               vbExclamation, "モジュール不足"
        Exit Sub
    End If
    
    ' ステップ2: シート作成
    result = MsgBox("ステップ1/3: InazumaGantt_v2 シートを作成します。" & vbCrLf & vbCrLf & _
                    "既存のシートは削除されますが、よろしいですか？", _
                    vbYesNo + vbQuestion, "シート作成")
    
    If result = vbYes Then
        Call InazumaGantt_v2.SetupInazumaGantt
        MsgBox "シートの作成が完了しました！", vbInformation, "完了"
    End If
    
    ' ステップ3: サンプルデータ
    result = MsgBox("ステップ2/3: サンプルデータを追加しますか？" & vbCrLf & vbCrLf & _
                    "サンプルデータを追加すると、すぐに動作を確認できます。", _
                    vbYesNo + vbQuestion, "サンプルデータ")
    
    If result = vbYes Then
        Call AddSampleData
        MsgBox "サンプルデータを追加しました！", vbInformation, "完了"
    End If
    
    ' ステップ4: シートモジュール
    result = MsgBox("ステップ3/3: シートモジュールの設定を行いますか？" & vbCrLf & vbCrLf & _
                    "シートモジュールを設定すると、以下の機能が有効になります：" & vbCrLf & _
                    "・タスク入力時の階層自動判定" & vbCrLf & _
                    "・進捗率入力時の状況自動更新" & vbCrLf & _
                    "・ダブルクリックでタスク完了" & vbCrLf & vbCrLf & _
                    "※手動での設定が必要です", _
                    vbYesNo + vbQuestion, "シートモジュール")
    
    If result = vbYes Then
        Call ShowSheetModuleInstructions
    End If
    
    ' 完了メッセージ
    MsgBox "セットアップウィザードが完了しました！" & vbCrLf & vbCrLf & _
           "次の手順：" & vbCrLf & _
           "1. RefreshInazumaGantt マクロでガント描画" & vbCrLf & _
           "2. ApplyHierarchyColors マクロで色分け" & vbCrLf & vbCrLf & _
           "詳細はREADME.mdを参照してください。", _
           vbInformation, "セットアップ完了"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "セットアップ中にエラーが発生しました：" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "エラー"
End Sub

' ==========================================
'  必須モジュールの存在チェック
' ==========================================
Private Function CheckRequiredModules() As Boolean
    Dim result As Boolean
    result = True

    If Not ModuleExists("InazumaGantt_v2") Then result = False
    If Not ModuleExists("HierarchyColor") Then result = False

    CheckRequiredModules = result
End Function

Private Function ModuleExists(ByVal moduleName As String) As Boolean
    On Error GoTo Fallback

    Dim vbComp As Object
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If StrComp(vbComp.Name, moduleName, vbTextCompare) = 0 Then
            ModuleExists = True
            Exit Function
        End If
    Next vbComp

    ModuleExists = False
    Exit Function

Fallback:
    ' VBProject?????????????????????
    ModuleExists = True
End Function

' ==========================================
'  サンプルデータの追加
' ==========================================
Private Sub AddSampleData()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("InazumaGantt_v2")
    
    Application.ScreenUpdating = False
    
    ' プロジェクト情報
    ws.Range("B2").Value = "サンプルプロジェクト"
    ws.Range("B3").Value = "プロジェクトマネージャー"
    ws.Range("K3").Value = Date
    ws.Range("K4").Value = 1
    ws.Range("M3").Value = Date
    
    Dim startRow As Long
    startRow = 9  ' ROW_DATA_START
    
    ' サンプルタスク
    ' フェーズ1（LV1）
    ws.Cells(startRow, "B").Value = 1
    ws.Cells(startRow, "C").Value = "フェーズ1：計画"
    ws.Cells(startRow, "G").Value = "プロジェクト計画フェーズ"
    ws.Cells(startRow, "H").Value = "完了"
    ws.Cells(startRow, "I").Value = 1
    ws.Cells(startRow, "J").Value = "山田"
    ws.Cells(startRow, "K").Value = Date
    ws.Cells(startRow, "L").Value = Date + 7
    ws.Cells(startRow, "M").Value = Date
    ws.Cells(startRow, "N").Value = Date + 5
    
    ' タスク1-1（LV2）
    ws.Cells(startRow + 1, "B").Value = 2
    ws.Cells(startRow + 1, "D").Value = "要件定義"
    ws.Cells(startRow + 1, "G").Value = "詳細な要件をまとめる"
    ws.Cells(startRow + 1, "H").Value = "完了"
    ws.Cells(startRow + 1, "I").Value = 1
    ws.Cells(startRow + 1, "J").Value = "佐藤"
    ws.Cells(startRow + 1, "K").Value = Date
    ws.Cells(startRow + 1, "L").Value = Date + 3
    ws.Cells(startRow + 1, "M").Value = Date
    ws.Cells(startRow + 1, "N").Value = Date + 3
    
    ' タスク1-2（LV2）
    ws.Cells(startRow + 2, "B").Value = 3
    ws.Cells(startRow + 2, "D").Value = "設計書作成"
    ws.Cells(startRow + 2, "G").Value = "基本設計書"
    ws.Cells(startRow + 2, "H").Value = "進行中"
    ws.Cells(startRow + 2, "I").Value = 0.6
    ws.Cells(startRow + 2, "J").Value = "鈴木"
    ws.Cells(startRow + 2, "K").Value = Date + 3
    ws.Cells(startRow + 2, "L").Value = Date + 7
    ws.Cells(startRow + 2, "M").Value = Date + 3
    
    ' フェーズ2（LV1）
    ws.Cells(startRow + 3, "B").Value = 4
    ws.Cells(startRow + 3, "C").Value = "フェーズ2：開発"
    ws.Cells(startRow + 3, "G").Value = "実装フェーズ"
    ws.Cells(startRow + 3, "H").Value = "未着手"
    ws.Cells(startRow + 3, "I").Value = 0
    ws.Cells(startRow + 3, "J").Value = "田中"
    ws.Cells(startRow + 3, "K").Value = Date + 7
    ws.Cells(startRow + 3, "L").Value = Date + 21
    
    ' 階層レベルを自動判定
    InazumaGantt_v2.AutoDetectTaskLevel 0  ' 全行
    
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "サンプルデータの追加でエラーが発生しました：" & vbCrLf & Err.Description, vbCritical
End Sub

' ==========================================
'  シートモジュール設定の説明表示
' ==========================================
Private Sub ShowSheetModuleInstructions()
    MsgBox "【シートモジュールの設定方法】" & vbCrLf & vbCrLf & _
           "1. Alt + F11 でVBAエディタを開く" & vbCrLf & _
           "2. プロジェクトエクスプローラーで" & vbCrLf & _
           "   「InazumaGantt_v2」シートをダブルクリック" & vbCrLf & _
           "3. 開いたウィンドウに、以下のファイルの内容を貼り付け：" & vbCrLf & _
           "   vba_modules\import\InazumaGantt_v2_SheetModule.bas" & vbCrLf & vbCrLf & _
           "詳細はSETUP.mdを参照してください。", _
           vbInformation, "シートモジュール設定"
End Sub

' ==========================================
'  クイックスタート（既存ユーザー向け）
' ==========================================
Sub QuickStart()
    On Error GoTo ErrorHandler
    
    Dim result As VbMsgBoxResult
    
    result = MsgBox("既に設定済みのユーザー向けクイックスタートです。" & vbCrLf & vbCrLf & _
                    "以下の処理を一括実行します：" & vbCrLf & _
                    "1. ガント更新（RefreshInazumaGantt）" & vbCrLf & _
                    "2. 色分け適用（ApplyHierarchyColors）" & vbCrLf & vbCrLf & _
                    "実行しますか？", _
                    vbYesNo + vbQuestion, "クイックスタート")
    
    If result = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    ' ガント更新
    Call InazumaGantt_v2.RefreshInazumaGantt
    
    ' 色分け適用
    Call HierarchyColor.ApplyHierarchyColors
    
    Application.ScreenUpdating = True
    
    MsgBox "クイックスタートが完了しました！", vbInformation, "完了"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical
End Sub
