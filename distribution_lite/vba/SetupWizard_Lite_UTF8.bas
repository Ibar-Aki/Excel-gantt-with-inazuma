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
    Dim hasShowcaseSample As Boolean

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
        hasShowcaseSample = True
    End If

    ' 設定マスタシート作成
    InazumaGantt_v3.EnsureSettingsSheet

    ' メインシートをアクティブに（重要：設定マスタではなくメインシートで描画）
    ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME).Activate

    ' 階層色分け設定
    HierarchyColor.SetupHierarchyColors

    ' ガントチャート描画
    InazumaGantt_v3.RefreshInazumaGantt

    If hasShowcaseSample Then
        WBSSampleShowcase.FinalizeShowcasePresentation True, WBSSampleShowcase.GetShowcaseReferenceDate(startDate)
    End If

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
    Dim mainSheet As Worksheet
    Dim hasShowcaseSample As Boolean

    ' ステップ1: 開始確認
    result = MsgBox("InazumaGantt Lite セットアップウィザードへようこそ！" & vbCrLf & vbCrLf & _
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
    result = MsgBox("新しいシート「" & InazumaGantt_v3.MAIN_SHEET_NAME & "」を作成しますか？" & vbCrLf & vbCrLf & _
                   "注意: 同名のシートが既に存在する場合は上書きされません。", _
                   vbQuestion + vbYesNo, "ステップ 1/3: シート作成")

    If result = vbYes Then
        CreateMainSheet
        Set mainSheet = GetMainSheet()
    Else
        Set mainSheet = GetMainSheet()
        If mainSheet Is Nothing Then
            MsgBox "メインシートを作成しない場合、セットアップは続行できません。", vbExclamation, "セットアップ"
            Exit Sub
        End If
    End If

    ' ステップ3: サンプルデータ
    result = MsgBox("サンプルデータを追加しますか？" & vbCrLf & vbCrLf & _
                   "サンプルデータには以下が含まれます:" & vbCrLf & _
                   "- 会話力向上アプリ開発をテーマにした12個のフェーズ" & vbCrLf & _
                   "- 約150行の構造化WBS" & vbCrLf & _
                   "- WBSサマリも同時生成" & vbCrLf & _
                   "- 実績列なし / 単線ガント", _
                   vbQuestion + vbYesNo, "ステップ 2/3: サンプルデータ")

    If result = vbYes Then
        AddSampleData
        hasShowcaseSample = True
    End If

    ' ステップ4: 階層色分けとガント描画を自動実行
    ' まずメインシートをアクティブにする
    Set mainSheet = GetMainSheet()
    If mainSheet Is Nothing Then
        MsgBox "メインシートが見つからないため、セットアップを続行できません。", vbCritical, "セットアップ"
        Exit Sub
    End If
    mainSheet.Activate

    Application.ScreenUpdating = False

' v3: 設定マスタシートを作成
    InazumaGantt_v3.EnsureSettingsSheet

    ' 階層色分けの条件付き書式を設定
    HierarchyColor.SetupHierarchyColors

    ' ガントチャートを描画
    InazumaGantt_v3.RefreshInazumaGantt

    If hasShowcaseSample Then
        Dim roadmapReferenceDate As Date
        If IsDate(mainSheet.Range(InazumaGantt_v3.CELL_PROJECT_START).Value) Then
            roadmapReferenceDate = WBSSampleShowcase.GetShowcaseReferenceDate(CDate(mainSheet.Range(InazumaGantt_v3.CELL_PROJECT_START).Value))
        Else
            roadmapReferenceDate = Date
        End If
        WBSSampleShowcase.FinalizeShowcasePresentation True, roadmapReferenceDate
    End If

    Application.ScreenUpdating = True

    ' ステップ5: 完了
    MsgBox "セットアップウィザードが完了しました！" & vbCrLf & vbCrLf & _
           "以下の設定が完了しました:" & vbCrLf & _
           "- シート作成（メイン、設定マスタ）" & vbCrLf & _
           "- 階層色分け（条件付き書式）" & vbCrLf & _
           "- ガントチャート描画" & vbCrLf & vbCrLf & _
           "【シートモジュールの設定】" & vbCrLf & _
           "ダブルクリック完了・折りたたみ機能を使うには、" & vbCrLf & _
           "SheetModule_Lite_UTF8.bas をシートモジュールに貼り付けてください。", _
           vbInformation, "セットアップ完了"
    Exit Sub

ErrorHandler:
    MsgBox "セットアップ中にエラーが発生しました: " & Err.Description, vbCritical, "エラー"
End Sub

Private Function GetMainSheet() As Worksheet
    On Error Resume Next
    Set GetMainSheet = ThisWorkbook.Worksheets(InazumaGantt_v3.MAIN_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function MainSheetNeedsSetup(ByVal ws As Worksheet) As Boolean
    MainSheetNeedsSetup = (Trim$(CStr(ws.Range("A" & InazumaGantt_v3.ROW_HEADER).Value)) <> "LV" Or _
                           Trim$(CStr(ws.Range("B" & InazumaGantt_v3.ROW_HEADER).Value)) <> "No." Or _
                           Trim$(CStr(ws.Range("K2").Value)) <> "開始日：")
End Function

Private Function MainSheetHasTaskData(ByVal ws As Worksheet) As Boolean
    Dim lastRow As Long

    lastRow = InazumaGantt_v3.GetLastDataRow(ws)
    If lastRow < InazumaGantt_v3.ROW_DATA_START Then Exit Function

    MainSheetHasTaskData = (Application.WorksheetFunction.CountA(ws.Range("C" & InazumaGantt_v3.ROW_DATA_START & ":F" & lastRow)) > 0)
End Function


' ==========================================
'  メインシートの作成
' ==========================================
Private Sub CreateMainSheet()
    Dim ws As Worksheet

    Set ws = GetMainSheet()

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = InazumaGantt_v3.MAIN_SHEET_NAME
        ws.Activate
        InazumaGantt_v3.SetupInazumaGantt False, Null
    Else
        ws.Activate
        If MainSheetNeedsSetup(ws) And Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            InazumaGantt_v3.SetupInazumaGantt False, Null
        Else
            MsgBox "既存のメインシートを使用します。" & vbCrLf & _
                   "既存データは上書きせず、そのまま維持します。", vbInformation, "セットアップ"
        End If
    End If
End Sub

' ==========================================
'  メインシートの作成（サイレント版）
' ==========================================
Private Sub CreateMainSheetSilent(ByVal startDateStr As String)
    Dim ws As Worksheet

    Set ws = GetMainSheet()

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = InazumaGantt_v3.MAIN_SHEET_NAME
        ws.Activate
        InazumaGantt_v3.SetupInazumaGantt True, startDateStr
    Else
        ws.Activate
        If MainSheetNeedsSetup(ws) And Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            InazumaGantt_v3.SetupInazumaGantt True, startDateStr
        End If
    End If
End Sub

' ==========================================
'  サンプルデータの追加
' ==========================================
' ==========================================
'  サンプルデータの追加（統合版）
' ==========================================
Private Sub AddSampleData(Optional ByVal baseDate As Date = 0)
    WBSSampleShowcase.CreateShowcaseSampleWBS baseDate, False, False, False
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
                  "   「" & InazumaGantt_v3.MAIN_SHEET_NAME & "」シートをダブルクリック" & vbCrLf & _
                  "3. vba/SheetModule_Lite_UTF8.bas の内容を" & vbCrLf & _
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
    status = status & "  InazumaGantt_v3: " & IIf(IsModuleInstalled("InazumaGantt_v3"), "OK", "未インストール") & vbCrLf
    status = status & "  WBSRoadmapReport: " & IIf(IsModuleInstalled("WBSRoadmapReport"), "OK", "未インストール") & vbCrLf
    status = status & "  WBSSampleShowcase: " & IIf(IsModuleInstalled("WBSSampleShowcase"), "OK", "未インストール") & vbCrLf
    status = status & "  HierarchyColor: " & IIf(IsModuleInstalled("HierarchyColor"), "OK", "未インストール") & vbCrLf
    status = status & "  SetupWizard: " & IIf(IsModuleInstalled("SetupWizard"), "OK", "未インストール") & vbCrLf

    status = status & vbCrLf & "手動設定:" & vbCrLf
    status = status & "  SheetModule: " & InazumaGantt_v3.MAIN_SHEET_NAME & " シートモジュールへ貼り付け" & vbCrLf

    MsgBox status, vbInformation, "インストール状態"
End Sub
