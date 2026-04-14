# Restoreからの作業手順書 Lite

- 作成日: 2026-04-15 01:05 JST
- 作成者: Codex (GPT-5)

## 対象

- bundle ファイル: `bundle_YYMMDD_InazumaGantt_Lite_Distribution_XX.txt`
- 変換器ルート配下の `restore_input` と `restore_output`

## 事前準備

1. 変換器ルートの `restore_input` にある既存の bundle ファイルを退避します。
2. 変換器ルートの `restore_output` に既存の復元結果がある場合は退避します。
3. Lite 版の配布 bundle を `restore_input` に 1 件だけ配置します。

## Restore

1. 変換器ルートで `restore_files.bat` を実行します。
2. 正常終了後、`restore_output\InazumaGantt_Lite_Distribution` が作成されることを確認します。
3. 復元された主な構成を確認します。

```text
InazumaGantt_Lite_Distribution
├─ excel
│  └─ WorkbookPayload.json
├─ scripts
│  ├─ OneClick_CreateLiteWorkbook.ps1
│  ├─ RestoreWorkbookFromPayload.ps1
│  └─ BuildInazumaGantt_Lite_UTF8.ps1
└─ vba
```

## 空のExcelファイル作成

1. Excel を起動します。
2. 空のブックを新規作成します。
3. `restore_output\InazumaGantt_Lite_Distribution\excel\BlankWorkbook.xlsx` として保存します。

## 配布同梱workbookの復元

1. `restore_output\InazumaGantt_Lite_Distribution\scripts\Run_RestoreWorkbookFromPayload.bat` を実行します。
2. `excel` フォルダに `InazumaGantt_Lite_*.xlsm` が生成されることを確認します。
3. 生成された `.xlsm` を開き、以下のシートがあることを確認します。

```text
InazumaGantt_Lite
InazumaGantt_説明
設定マスタ
WBSサマリ
```

## ワンクリック生成

1. `restore_output\InazumaGantt_Lite_Distribution\scripts\Run_OneClick_CreateLiteWorkbook.bat` を実行します。
2. `output` フォルダに `InazumaGantt_Lite_*.xlsm` が新規生成されることを確認します。
3. 生成された `.xlsm` を開き、以下を確認します。

- シート構成が `InazumaGantt_Lite / InazumaGantt_説明 / 設定マスタ / WBSサマリ`
- `InazumaGantt_Lite` にサンプル WBS が入っている
- `WBSサマリ` が作成されている
- 開始実績 / 完了実績列がなく、ガントは単線表示になっている

## 注意点

- 変換器は `.xlsm` を bundle に含めないため、配布同梱 workbook を復元するには `WorkbookPayload.json` からの復元が必要です。
- `restore_input` には bundle ファイルを 1 件だけ置いてください。複数件あると復元は失敗します。
- `restore_output` に同名ファイルが残っていると復元に失敗します。
- `Run_OneClick_CreateLiteWorkbook.bat` は Windows 標準の `powershell.exe` でも動作します。PowerShell 7 は必須ではありません。
- 別PCで workbook 生成に失敗する場合は、Excel の `VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` を有効にしてください。
