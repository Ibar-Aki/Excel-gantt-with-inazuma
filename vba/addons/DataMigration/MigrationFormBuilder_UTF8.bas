Attribute VB_Name = "MigrationFormBuilder"
Option Explicit

' ==========================================
'  UserForm動的生成マクロ
' ==========================================
' frmMigrationWizard を自動生成します
' 初回セットアップ時に一度だけ実行してください

' ==========================================
'  ウィザードフォームの作成
' ==========================================
Public Sub CreateMigrationWizardForm()
    On Error GoTo ErrorHandler
    
    ' VBComponentsへのアクセス権限チェック
    On Error Resume Next
    Dim testAccess As Object
    Set testAccess = ThisWorkbook.VBProject.VBComponents
    If Err.Number <> 0 Then
        If Application.DisplayAlerts Then
            MsgBox "VBAプロジェクトへのアクセスが拒否されました。" & vbCrLf & vbCrLf & _
                   "【解決方法】" & vbCrLf & _
                   "1. Excelのオプション→トラストセンター→トラストセンターの設定" & vbCrLf & _
                   "2. マクロの設定→VBAプロジェクトオブジェクトモデルへのアクセスを信頼する" & vbCrLf & _
                   "3. にチェックを入れてください", vbExclamation, "アクセス拒否"
        End If
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' 既存のフォームを削除
    Dim i As Long
    For i = ThisWorkbook.VBProject.VBComponents.Count To 1 Step -1
        If ThisWorkbook.VBProject.VBComponents(i).Name = "frmMigrationWizard" Then
            ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(i)
        End If
    Next i
    
    ' 新しいUserFormを追加
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    vbComp.Name = "frmMigrationWizard"
    
    ' フォームのプロパティ設定
    With vbComp.Properties
        .Item("Caption").Value = "データ移管ウィザード"
        .Item("Width").Value = 480
        .Item("Height").Value = 480 ' Item 8: 高さ拡張
    End With
    
    ' コントロールを追加
    Dim ctrl As Object
    Dim yPos As Long
    yPos = 10
    
    ' ========== ステップ表示ラベル ==========
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblStep"
        .Caption = "Step 1: 移管元シート選択"
        .Left = 10
        .Top = yPos
        .Width = 460
        .Height = 20
        .Font.Size = 12
        .Font.Bold = True
    End With
    yPos = yPos + 30
    
    ' ========== Step 1 パネル ==========
    ' 移管元シート選択
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblSourceSheet"
        .Caption = "移管元シート:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboSourceSheet"
        .Left = 120
        .Top = yPos
        .Width = 200
        .Height = 20
    End With
    yPos = yPos + 35
    
    ' 保存済み設定の読み込みボタン
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnLoadConfig"
        .Caption = "保存済み設定を読み込み..."
        .Left = 120
        .Top = yPos
        .Width = 150
        .Height = 24
    End With
    yPos = yPos + 40
    
    ' ========== Step 2 パネル (初期は非表示) ==========
    ' 階層判定モード選択
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblMode"
        .Caption = "判定形式:"
        .Left = 10
        .Top = yPos
        .Width = 60
        .Height = 18
        .Visible = False
    End With

    Set ctrl = vbComp.Designer.Controls.Add("Forms.OptionButton.1")
    With ctrl
        .Name = "optModeWBS"
        .Caption = "WBS番号 (1.1.1)"
        .GroupName = "HierarchyMode"
        .Value = True
        .Left = 80
        .Top = yPos
        .Width = 120
        .Height = 18
        .Visible = False
    End With

    Set ctrl = vbComp.Designer.Controls.Add("Forms.OptionButton.1")
    With ctrl
        .Name = "optModeLevel"
        .Caption = "レベル数値 (1,2...)"
        .GroupName = "HierarchyMode"
        .Left = 210
        .Top = yPos
        .Width = 120
        .Height = 18
        .Visible = False
    End With
    yPos = yPos + 30

    ' 階層列（旧WBS列）
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblWBSColumn"
        .Caption = "階層列 *:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboWBSColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' タスク名列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblTaskColumn"
        .Caption = "タスク名列 *:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboTaskColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' 担当者列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblAssigneeColumn"
        .Caption = "担当者列:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboAssigneeColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' アイテム8: 開始予定列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblStartPlanColumn"
        .Caption = "開始予定列:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboStartPlanColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' 完了予定列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblEndPlanColumn"
        .Caption = "完了予定列:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboEndPlanColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' アイテム8: 開始実績列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblStartActualColumn"
        .Caption = "開始実績列:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboStartActualColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' アイテム8: 完了実績列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblEndActualColumn"
        .Caption = "完了実績列:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboEndActualColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30

    ' 進捗率列
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblProgressColumn"
        .Caption = "進捗率列:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ComboBox.1")
    With ctrl
        .Name = "cboProgressColumn"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' データ開始行
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .Name = "lblDataStartRow"
        .Caption = "データ開始行 *:"
        .Left = 10
        .Top = yPos
        .Width = 100
        .Height = 18
        .Visible = False
    End With
    
    Set ctrl = vbComp.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl
        .Name = "txtDataStartRow"
        .Left = 120
        .Top = yPos
        .Width = 80
        .Height = 20
        .Text = "2"
        .Visible = False
    End With
    yPos = yPos + 35
    
    ' 設定保存チェックボックス
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CheckBox.1")
    With ctrl
        .Name = "chkSaveConfig"
        .Caption = "この設定を保存する"
        .Left = 10
        .Top = yPos
        .Width = 150
        .Height = 20
        .Visible = False
    End With
    yPos = yPos + 30
    
    ' ========== Step 3 パネル (初期は非表示) ==========
    ' プレビューリストボックス
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl
        .Name = "lstPreview"
        .Left = 10
        .Top = 40
        .Width = 460
        .Height = 320 ' 少し拡張
        .Visible = False
    End With
    
    ' ========== ナビゲーションボタン ==========
    yPos = 400 ' 下部に配置
    
    ' 戻るボタン
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnBack"
        .Caption = "< 戻る"
        .Left = 150
        .Top = yPos
        .Width = 80
        .Height = 28
        .Enabled = False
    End With
    
    ' 次へボタン
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnNext"
        .Caption = "次へ >"
        .Left = 240
        .Top = yPos
        .Width = 80
        .Height = 28
    End With
    
    ' キャンセルボタン
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnCancel"
        .Caption = "キャンセル"
        .Left = 330
        .Top = yPos
        .Width = 80
        .Height = 28
    End With
    
    ' 実行ボタン（初期は非表示）
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .Name = "btnExecute"
        .Caption = "実行"
        .Left = 240
        .Top = yPos
        .Width = 80
        .Height = 28
        .Visible = False
    End With
    
    ' ========== フォームのコードを追加 ==========
    AddFormCode vbComp
    
    If Application.DisplayAlerts Then
        MsgBox "ウィザードフォーム (frmMigrationWizard) を作成しました！" & vbCrLf & vbCrLf & _
               "ShowMigrationWizard() マクロでウィザードを起動できます。", vbInformation, "作成完了"
    End If
    Exit Sub
    
    Exit Sub
    
ErrorHandler:
    If Application.DisplayAlerts Then
        MsgBox "フォーム作成エラー: " & Err.Description, vbCritical, "エラー"
    End If
End Sub

' ==========================================
'  フォームのVBAコードを追加
' ==========================================
Private Sub AddFormCode(ByRef vbComp As Object)
    Dim code As String
    
    code = "Option Explicit" & vbCrLf & vbCrLf
    code = code & "Private currentStep As Long" & vbCrLf & vbCrLf
    
    ' Initialize
    code = code & "Private Sub UserForm_Initialize()" & vbCrLf
    code = code & "    currentStep = 1" & vbCrLf
    code = code & "    LoadSheetList" & vbCrLf
    code = code & "    LoadColumnList" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' LoadSheetList
    code = code & "Private Sub LoadSheetList()" & vbCrLf
    code = code & "    Dim ws As Worksheet" & vbCrLf
    code = code & "    For Each ws In ThisWorkbook.Worksheets" & vbCrLf
    code = code & "        If ws.Name <> ""InazumaGantt_v2"" And ws.Name <> ""設定マスタ"" And ws.Name <> ""移管設定"" And ws.Name <> ""祝日マスタ"" And ws.Name <> ""InazumaGantt_説明"" Then" & vbCrLf
    code = code & "            cboSourceSheet.AddItem ws.Name" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    Next ws" & vbCrLf
    code = code & "    If cboSourceSheet.ListCount > 0 Then cboSourceSheet.ListIndex = 0" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' LoadColumnList - Item 7: A-Z limit removed
    code = code & "Private Sub LoadColumnList()" & vbCrLf
    code = code & "    Dim i As Long" & vbCrLf
    code = code & "    Dim colName As String" & vbCrLf
    ' A-Z
    code = code & "    For i = 1 To 26" & vbCrLf
    code = code & "        colName = Chr(64 + i)" & vbCrLf
    code = code & "        AddColumnItem colName" & vbCrLf
    code = code & "    Next i" & vbCrLf
    ' AA-AZ (Item 7)
    code = code & "    For i = 1 To 26" & vbCrLf
    code = code & "        colName = ""A"" & Chr(64 + i)" & vbCrLf
    code = code & "        AddColumnItem colName" & vbCrLf
    code = code & "    Next i" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' AddColumnItem Helper
    code = code & "Private Sub AddColumnItem(colName As String)" & vbCrLf
    code = code & "    cboWBSColumn.AddItem colName" & vbCrLf
    code = code & "    cboTaskColumn.AddItem colName" & vbCrLf
    code = code & "    cboAssigneeColumn.AddItem colName" & vbCrLf
    code = code & "    cboEndPlanColumn.AddItem colName" & vbCrLf
    code = code & "    cboProgressColumn.AddItem colName" & vbCrLf
    code = code & "    cboStartPlanColumn.AddItem colName" & vbCrLf ' Item 8
    code = code & "    cboStartActualColumn.AddItem colName" & vbCrLf ' Item 8
    code = code & "    cboEndActualColumn.AddItem colName" & vbCrLf ' Item 8
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' btnNext_Click
    code = code & "Private Sub btnNext_Click()" & vbCrLf
    code = code & "    If currentStep = 1 Then" & vbCrLf
    code = code & "        If cboSourceSheet.Text = """" Then" & vbCrLf
    code = code & "            MsgBox ""移管元シートを選択してください"", vbExclamation" & vbCrLf
    code = code & "            Exit Sub" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "        ShowStep2" & vbCrLf
    code = code & "    ElseIf currentStep = 2 Then" & vbCrLf
    code = code & "        If cboWBSColumn.Text = """" Or cboTaskColumn.Text = """" Or txtDataStartRow.Text = """" Then" & vbCrLf
    code = code & "            MsgBox ""必須項目(*) を入力してください"", vbExclamation" & vbCrLf
    code = code & "            Exit Sub" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "        ShowStep3" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' btnBack_Click
    code = code & "Private Sub btnBack_Click()" & vbCrLf
    code = code & "    If currentStep = 2 Then" & vbCrLf
    code = code & "        ShowStep1" & vbCrLf
    code = code & "    ElseIf currentStep = 3 Then" & vbCrLf
    code = code & "        ShowStep2" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' btnExecute_Click
    code = code & "Private Sub btnExecute_Click()" & vbCrLf
    code = code & "    Dim config As DataMigrationWizard.MappingConfig" & vbCrLf
    code = code & "    config.SourceSheetName = cboSourceSheet.Text" & vbCrLf
    code = code & "    config.WBSColumn = cboWBSColumn.Text" & vbCrLf
    code = code & "    config.TaskNameColumn = cboTaskColumn.Text" & vbCrLf
    code = code & "    config.AssigneeColumn = cboAssigneeColumn.Text" & vbCrLf
    code = code & "    config.EndPlanColumn = cboEndPlanColumn.Text" & vbCrLf
    code = code & "    config.ProgressColumn = cboProgressColumn.Text" & vbCrLf
    code = code & "    config.StartPlanColumn = cboStartPlanColumn.Text" & vbCrLf ' Item 8
    code = code & "    config.StartActualColumn = cboStartActualColumn.Text" & vbCrLf ' Item 8
    code = code & "    config.EndActualColumn = cboEndActualColumn.Text" & vbCrLf ' Item 8
    code = code & "    config.DataStartRow = CLng(txtDataStartRow.Text)" & vbCrLf
    code = code & "    If optModeWBS.Value Then" & vbCrLf
    code = code & "        config.HierarchyMode = 0 ' WBS" & vbCrLf
    code = code & "    Else" & vbCrLf
    code = code & "        config.HierarchyMode = 1 ' Level" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    If chkSaveConfig.Value Then" & vbCrLf
    code = code & "        config.TemplateName = InputBox(""設定名を入力してください:"", ""設定保存"")" & vbCrLf
    code = code & "        If config.TemplateName <> """" Then" & vbCrLf
    code = code & "            DataMigrationWizard.SaveMappingConfig config" & vbCrLf
    code = code & "        End If" & vbCrLf
    code = code & "    End If" & vbCrLf
    code = code & "    Me.Hide" & vbCrLf
    code = code & "    DataMigrationWizard.ExecuteMigration config" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' btnCancel_Click
    code = code & "Private Sub btnCancel_Click()" & vbCrLf
    code = code & "    Unload Me" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' ShowStep1
    code = code & "Private Sub ShowStep1()" & vbCrLf
    code = code & "    currentStep = 1" & vbCrLf
    code = code & "    lblStep.Caption = ""Step 1: 移管元シート選択""" & vbCrLf
    code = code & "    lblSourceSheet.Visible = True" & vbCrLf
    code = code & "    cboSourceSheet.Visible = True" & vbCrLf
    code = code & "    btnLoadConfig.Visible = True" & vbCrLf
    code = code & "    HideStep2Controls" & vbCrLf
    code = code & "    lstPreview.Visible = False" & vbCrLf
    code = code & "    btnBack.Enabled = False" & vbCrLf
    code = code & "    btnNext.Visible = True" & vbCrLf
    code = code & "    btnExecute.Visible = False" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' ShowStep2
    code = code & "Private Sub ShowStep2()" & vbCrLf
    code = code & "    currentStep = 2" & vbCrLf
    code = code & "    lblStep.Caption = ""Step 2: 列マッピング設定""" & vbCrLf
    code = code & "    lblSourceSheet.Visible = False" & vbCrLf
    code = code & "    cboSourceSheet.Visible = False" & vbCrLf
    code = code & "    btnLoadConfig.Visible = False" & vbCrLf
    code = code & "    ShowStep2Controls" & vbCrLf
    code = code & "    lstPreview.Visible = False" & vbCrLf
    code = code & "    btnBack.Enabled = True" & vbCrLf
    code = code & "    btnNext.Visible = True" & vbCrLf
    code = code & "    btnExecute.Visible = False" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' ShowStep3
    code = code & "Private Sub ShowStep3()" & vbCrLf
    code = code & "    currentStep = 3" & vbCrLf
    code = code & "    lblStep.Caption = ""Step 3: プレビュー確認""" & vbCrLf
    code = code & "    HideStep2Controls" & vbCrLf
    code = code & "    lstPreview.Visible = True" & vbCrLf
    code = code & "    lstPreview.Clear" & vbCrLf
    code = code & "    lstPreview.AddItem ""移管プレビュー（先頭5行のみ表示）""" & vbCrLf
    code = code & "    lstPreview.AddItem ""LV | タスク名 | 担当者 | 完了予定""" & vbCrLf
    code = code & "    lstPreview.AddItem ""--------------------------------------------------""" & vbCrLf
    code = code & "    btnBack.Enabled = True" & vbCrLf
    code = code & "    btnNext.Visible = False" & vbCrLf
    code = code & "    btnExecute.Visible = True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' ShowStep2Controls
    code = code & "Private Sub ShowStep2Controls()" & vbCrLf
    code = code & "    lblMode.Visible = True" & vbCrLf
    code = code & "    optModeWBS.Visible = True" & vbCrLf
    code = code & "    optModeLevel.Visible = True" & vbCrLf
    code = code & "    lblWBSColumn.Visible = True" & vbCrLf
    code = code & "    cboWBSColumn.Visible = True" & vbCrLf
    code = code & "    lblTaskColumn.Visible = True" & vbCrLf
    code = code & "    cboTaskColumn.Visible = True" & vbCrLf
    code = code & "    lblAssigneeColumn.Visible = True" & vbCrLf
    code = code & "    cboAssigneeColumn.Visible = True" & vbCrLf
    code = code & "    lblEndPlanColumn.Visible = True" & vbCrLf
    code = code & "    cboEndPlanColumn.Visible = True" & vbCrLf
    ' Item 8 additions
    code = code & "    lblStartPlanColumn.Visible = True" & vbCrLf
    code = code & "    cboStartPlanColumn.Visible = True" & vbCrLf
    code = code & "    lblStartActualColumn.Visible = True" & vbCrLf
    code = code & "    cboStartActualColumn.Visible = True" & vbCrLf
    code = code & "    lblEndActualColumn.Visible = True" & vbCrLf
    code = code & "    cboEndActualColumn.Visible = True" & vbCrLf
    '
    code = code & "    lblProgressColumn.Visible = True" & vbCrLf
    code = code & "    cboProgressColumn.Visible = True" & vbCrLf
    code = code & "    lblDataStartRow.Visible = True" & vbCrLf
    code = code & "    txtDataStartRow.Visible = True" & vbCrLf
    code = code & "    chkSaveConfig.Visible = True" & vbCrLf
    code = code & "End Sub" & vbCrLf & vbCrLf
    
    ' HideStep2Controls
    code = code & "Private Sub HideStep2Controls()" & vbCrLf
    code = code & "    lblMode.Visible = False" & vbCrLf
    code = code & "    optModeWBS.Visible = False" & vbCrLf
    code = code & "    optModeLevel.Visible = False" & vbCrLf
    code = code & "    lblWBSColumn.Visible = False" & vbCrLf
    code = code & "    cboWBSColumn.Visible = False" & vbCrLf
    code = code & "    lblTaskColumn.Visible = False" & vbCrLf
    code = code & "    cboTaskColumn.Visible = False" & vbCrLf
    code = code & "    lblAssigneeColumn.Visible = False" & vbCrLf
    code = code & "    cboAssigneeColumn.Visible = False" & vbCrLf
    code = code & "    lblEndPlanColumn.Visible = False" & vbCrLf
    code = code & "    cboEndPlanColumn.Visible = False" & vbCrLf
    ' Item 8 additions
    code = code & "    lblStartPlanColumn.Visible = False" & vbCrLf
    code = code & "    cboStartPlanColumn.Visible = False" & vbCrLf
    code = code & "    lblStartActualColumn.Visible = False" & vbCrLf
    code = code & "    cboStartActualColumn.Visible = False" & vbCrLf
    code = code & "    lblEndActualColumn.Visible = False" & vbCrLf
    code = code & "    cboEndActualColumn.Visible = False" & vbCrLf
    '
    code = code & "    lblProgressColumn.Visible = False" & vbCrLf
    code = code & "    cboProgressColumn.Visible = False" & vbCrLf
    code = code & "    lblDataStartRow.Visible = False" & vbCrLf
    code = code & "    txtDataStartRow.Visible = False" & vbCrLf
    code = code & "    chkSaveConfig.Visible = False" & vbCrLf
    code = code & "End Sub" & vbCrLf
    
    ' コードをフォームに追加
    vbComp.CodeModule.AddFromString code
End Sub
