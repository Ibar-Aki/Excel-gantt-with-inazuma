# コード品質改善ガイド

## 実施した改善

### 1. エラーハンドリングの統一 ✅

#### Before
```vba
Sub SomeFunction()
    On Error GoTo ErrorHandler
    ' 処理
ErrorHandler:
    MsgBox "エラー: " & Err.Description
End Sub
```

#### After
```vba
Sub SomeFunction()
    On Error GoTo ErrorHandler
    ' 処理
    Exit Sub
ErrorHandler:
    ErrorHandler.HandleError "ModuleName", "SomeFunction", _
                             "ユーザーに表示するメッセージ"
End Sub
```

#### 改善点
- ログファイルに自動出力
- デバッグ情報の記録
- ユーザーフレンドリーなエラーメッセージ
- エラーコードの提供（"ERR" + 番号）

---

### 2. マジックナンバーの削除 ✅

#### Before
```vba
If cell.Row >= 9 Then  ' データ開始行
    ' 処理
End If
```

#### After
```vba
Private Const ROW_DATA_START As Long = 9

If cell.Row >= ROW_DATA_START Then
    ' 処理
End If
```

#### 削除したマジックナンバー一覧

| 元の値 | 定数名 | 説明 |
|--------|--------|------|
| 9 | ROW_DATA_START | データ開始行 |
| 120 | GANTT_DAYS | ガント表示日数 |
| 200 | DATA_ROWS_DEFAULT | デフォルト行数 |
| 1 | (LV定数として) | 階層レベル |

---

### 3. テストコードの追加 ✅

#### テストモジュール構成

**InazumaGanttTests_SJIS.bas** を追加

- `RunAllTests()` - 全テスト実行
- `Test_GetTaskColumnByLevel()` - 階層列判定テスト
- `AssertEquals()` - アサーション関数
- `IntegrationTest_FullWorkflow()` - 統合テスト

#### 使用方法

```vba
' VBAエディタで実行
Alt + F8 → RunAllTests → 実行

' または
Call InazumaGanttTests.RunAllTests
```

#### テスト結果

イミディエイトウィンドウに出力：
```
==========================================
InazumaGantt v2 テスト開始
==========================================
[PASS] GetTaskColumnByLevel - LV1
[PASS] GetTaskColumnByLevel - LV2
[PASS] GetTaskColumnByLevel - LV3
...
==========================================
テスト完了
成功: 5
失敗: 0
==========================================
```

---

## 入力値検証の追加

### ErrorHandler モジュールの検証関数

#### 1. 必須チェック
```vba
If ErrorHandler.ValidateRequired(value, "タスク名") Then
    ' OK
End If
```

#### 2. 数値チェック
```vba
If ErrorHandler.ValidateNumeric(value, "進捗率", 0, 100) Then
    ' OK
End If
```

#### 3. 日付チェック
```vba
If ErrorHandler.ValidateDate(value, "開始日") Then
    ' OK
End If
```

---

## 長時間処理の進捗表示

```vba
Sub LongProcess()
    Dim i As Long
    Dim total As Long
    total = 100
    
    For i = 1 To total
        ' 処理
        ErrorHandler.ShowProgress i, total, "処理中..."
        DoEvents
    Next i
    
    ErrorHandler.ClearProgress
End Sub
```

---

## ログ機能の使用

### ログ出力

```vba
' デバッグログ
ErrorHandler.WriteLog "処理開始", ErrorHandler.LOG_DEBUG

' 情報ログ
ErrorHandler.WriteLog "データ読み込み完了", ErrorHandler.LOG_INFO

' 警告ログ
ErrorHandler.WriteLog "データが見つかりません", ErrorHandler.LOG_WARNING

' エラーログ
ErrorHandler.WriteLog "ファイルアクセスエラー", ErrorHandler.LOG_ERROR
```

### ログファイル

保存先: `InazumaGantt_ErrorLog.txt`（Excelファイルと同じフォルダ）

フォーマット:
```
[ERROR] 2026-01-01 18:00:00
Module: InazumaGantt_v2
Procedure: DrawGanttBars
Error #13: Type mismatch
Source: VBAProject
--------------------------------------------------------------------------------
```

---

## コード品質チェックリスト

### ✅ 実施済み

- [x] エラーハンドリングの統一
- [x] マジックナンバーの削除
- [x] テストコードの追加
- [x] 入力値検証の実装
- [x] ログ機能の実装

### 🟡 推奨事項

- [ ] 全モジュールでErrorHandlerを使用
- [ ] 全関数に単体テストを追加
- [ ] パフォーマンステストの実施
- [ ] セキュリティ監査

### 📋 今後の改善候補

- [ ] 定数を設定ファイル化
- [ ] 国際化対応（i18n）
- [ ] アクセシビリティ改善
- [ ] コードカバレッジ測定

---

## ベストプラクティス

### 1. 定数の命名規則

```vba
' 推奨
Public Const ROW_DATA_START As Long = 9
Private Const MAX_RETRY_COUNT As Long = 3

' 非推奨
Const x = 9
Dim StartRow = 9  ' Constでない
```

### 2. エラーハンドリング

```vba
' 推奨
Sub DoSomething()
    On Error GoTo ErrorHandler
    ' 処理
    Exit Sub  ' 重要: ErrorHandlerに落ちないように
ErrorHandler:
    ErrorHandler.HandleError "Module", "Procedure", "Message"
End Sub

' 非推奨
Sub DoSomething()
    On Error Resume Next  ' エラーを無視
    ' 処理
End Sub
```

### 3. 入力値検証

```vba
' 推奨
If Not ErrorHandler.ValidateNumeric(progress, "進捗率", 0, 100) Then
    Exit Sub
End If

' 非推奨
If IsNumeric(progress) Then
    ' 検証なし
End If
```

---

## パフォーマンス最適化

### 画面更新の制御

```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

' 処理

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
```

### イベントの制御

```vba
Application.EnableEvents = False

' 処理

Application.EnableEvents = True
```

---

詳細は各モジュールのコメントを参照してください。
