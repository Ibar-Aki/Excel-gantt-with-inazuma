# Restoreからの作業手順書

- 作成日: 2026-04-12 10:16 JST
- 作成者: Codex (GPT-5)

## 対象

- bundle ファイル: `bundle_YYMMDD_InazumaGantt_v3_Distribution_XX.txt`
- 変換器: `C:\Work_Codex\ps1,batの変換器`

## 事前準備

1. `C:\Work_Codex\ps1,batの変換器\restore_input` にある既存の bundle ファイルを退避します。
2. `C:\Work_Codex\ps1,batの変換器\restore_output` に既存の復元結果がある場合は退避します。
3. 配布 bundle を `restore_input` に 1 件だけ配置します。

## Restore

1. `C:\Work_Codex\ps1,batの変換器\restore_files.bat` を実行します。
2. 正常終了後、`C:\Work_Codex\ps1,batの変換器\restore_output\InazumaGantt_v3_Distribution` が作成されることを確認します。
3. 復元された主な構成を確認します。

```text
InazumaGantt_v3_Distribution
├─ excel
│  └─ WorkbookPayload.json
├─ scripts
│  ├─ OneClick_CreateLatestWorkbook.ps1
│  ├─ RestoreWorkbookFromPayload.ps1
│  └─ BuildInazumaGantt_UTF8.ps1
└─ vba
```

## 空のExcelファイル作成

1. Excel を起動します。
2. 空のブックを新規作成します。
3. `C:\Work_Codex\ps1,batの変換器\restore_output\InazumaGantt_v3_Distribution\excel\BlankWorkbook.xlsx` として保存します。

## 配布同梱workbookの復元

1. `C:\Work_Codex\ps1,batの変換器\restore_output\InazumaGantt_v3_Distribution\scripts\Run_RestoreWorkbookFromPayload.bat` を実行します。
2. `excel` フォルダに `InazumaGantt_v3_*.xlsm` が生成されることを確認します。
3. 生成された `.xlsm` を開き、以下のシートがあることを確認します。

```text
InazumaGantt_v3
InazumaGantt_説明
設定マスタ
WBSサマリ
```

## ワンクリック生成

1. `C:\Work_Codex\ps1,batの変換器\restore_output\InazumaGantt_v3_Distribution\scripts\Run_OneClick_CreateLatestWorkbook.bat` を実行します。
2. `output` フォルダに `InazumaGantt_v3_*.xlsm` が新規生成されることを確認します。
3. 生成された `.xlsm` を開き、以下を確認します。

- シート構成が `InazumaGantt_v3 / InazumaGantt_説明 / 設定マスタ / WBSサマリ`
- `InazumaGantt_v3` にサンプル WBS が入っている
- `WBSサマリ` が作成されている

## 動作確認結果

- Restore: 正常
- 空の Excel ファイル作成: 正常
- Payload からの `.xlsm` 復元: 正常
- One-click 生成: 正常

## 注意点

- 変換器は `.xlsm` を bundle に含めないため、配布同梱 workbook を復元するには `WorkbookPayload.json` からの復元が必要です。
- `restore_input` には bundle ファイルを 1 件だけ置いてください。複数件あると復元は失敗します。
- `restore_output` に同名ファイルが残っていると復元に失敗します。
