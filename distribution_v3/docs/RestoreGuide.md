# Restoreからの作業手順書

- 作成日: 2026-04-12 10:16 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-04-24

## 対象

- bundle ファイル: `bundle_YYMMDD_InazumaGantt_v3_Distribution_XX.txt`
- 変換器ルート配下の `restore_input` と `restore_output`

## 事前準備

1. 変換器ルートの `restore_input` にある既存の bundle ファイルを退避します。
2. 変換器ルートの `restore_output` に既存の復元結果がある場合は退避します。
3. 配布 bundle を `restore_input` に 1 件だけ配置します。

## Restore

1. 変換器ルートで `restore_files.bat` を実行します。
2. 正常終了後、`restore_output\InazumaGantt_v3_Distribution` が作成されることを確認します。
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
3. `restore_output\InazumaGantt_v3_Distribution\excel\BlankWorkbook.xlsx` として保存します。

## 配布同梱workbookの復元

1. `restore_output\InazumaGantt_v3_Distribution\scripts\Run_RestoreWorkbookFromPayload.bat` を実行します。
2. `excel` フォルダに `InazumaGantt_v3_*.xlsm` が生成されることを確認します。
3. 生成された `.xlsm` を開き、以下のシートがあることを確認します。

```text
InazumaGantt_v3
InazumaGantt_説明
設定マスタ
WBSサマリ
```

`WBS_Backup_v3` や `_WBS_` で始まる一時シートが含まれていないことも確認します。バックアップシートは利用者が `WBS退避` を実行した時点で作成されます。

## ワンクリック生成

1. `restore_output\InazumaGantt_v3_Distribution\scripts\Run_OneClick_CreateLatestWorkbook.bat` を実行します。
2. `output` フォルダに `InazumaGantt_v3_*.xlsm` が新規生成されることを確認します。
3. 生成された `.xlsm` を開き、以下を確認します。

- シート構成が `InazumaGantt_v3 / InazumaGantt_説明 / 設定マスタ / WBSサマリ`
- `InazumaGantt_v3` にサンプル WBS が入っている
- `WBSサマリ` が作成されている
- `WBS_Backup_v3` や `_WBS_` 一時シートが残っていない

## 動作確認結果

- Restore: 正常
- 空の Excel ファイル作成: 正常
- Payload からの `.xlsm` 復元: 正常
- One-click 生成: 正常
- Build smoke 後のバックアップシート削除: 正常

## 補助情報のみ行の仕様

- `C:F` のタスク名を消しても、詳細・状況・進捗率・担当・開発LT・予定日・実績日が残っていれば、その行は有効行として保持されます。
- この場合は `C` 列に `（補助情報のみ）` が表示され、`No.`・親集計・`WBSサマリ` にも反映されます。
- 行を完全に空にしたときだけ、`LV/No./補助表示` がクリアされます。

## 注意点

- 変換器は `.xlsm` を bundle に含めないため、配布同梱 workbook を復元するには `WorkbookPayload.json` からの復元が必要です。
- 配布ブックにはビルド時に作成された WBS バックアップを含めません。復元後に `WBS退避` を押すと、利用者の現在の WBS から `WBS_Backup_v3` が作成されます。
- `バックアップ復元` は復元前の本 WBS を一時退避し、途中失敗時は直前状態へロールバックします。正常完了後、一時シートは削除されます。
- `restore_input` には bundle ファイルを 1 件だけ置いてください。複数件あると復元は失敗します。
- `restore_output` に同名ファイルが残っていると復元に失敗します。
- `Run_OneClick_CreateLatestWorkbook.bat` は Windows 標準の `powershell.exe` でも動作します。PowerShell 7 は必須ではありません。
- 別PCで workbook 生成に失敗する場合は、Excel の `VBA プロジェクト オブジェクト モデルへのアクセスを信頼する` を有効にしてください。
