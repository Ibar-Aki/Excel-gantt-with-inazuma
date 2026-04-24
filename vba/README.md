# VBAモジュール

更新日: 2026-04-24

InazumaGantt v3 / Lite で使用する VBA モジュール一覧です。

## エンコーディング

| サフィックス | エンコーディング | 用途 |
|-------------|-----------------|------|
| `_SJIS.bas` | Shift-JIS (CP932) | Excel にインポート |
| `_UTF8.bas` | UTF-8 (BOM なし) | 編集・手動コピー/貼り付け |

> 標準モジュールを Excel にインポートするときは必ず `_SJIS.bas` を使ってください。
> `*_UTF8.bas` は編集用であり、手動でコードをコピーして貼り付ける場合に使います。

## 必須モジュール

| ファイル | 用途 |
|----------|------|
| `InazumaGantt_v3_SJIS.bas` | メイン機能 |
| `InazumaGantt_Lite_SJIS.bas` | Lite版メイン機能 |
| `WBSRoadmapReport_SJIS.bas` | WBSサマリ |
| `WBSRoadmapReport_Lite_SJIS.bas` | Lite版 WBSサマリ |
| `WBSSampleShowcase_SJIS.bas` | ShowcaseサンプルWBS生成 |
| `WBSSampleShowcase_Lite_SJIS.bas` | Lite版 ShowcaseサンプルWBS生成 |
| `SetupWizard_SJIS.bas` | セットアップウィザード |
| `SetupWizard_Lite_SJIS.bas` | Lite版セットアップウィザード |
| `HierarchyColor_SJIS.bas` | 階層色分け |
| `HierarchyColor_Lite_SJIS.bas` | Lite版階層色分け |
| `SheetModule_SJIS.bas` | シートイベントコード |
| `SheetModule_Lite_SJIS.bas` | Lite版シートイベントコード |

## インポート手順

1. Excel ファイルを開く
2. `Alt + F11` で VBA エディタを開く
3. 標準版は `InazumaGantt_v3_SJIS.bas` `WBSRoadmapReport_SJIS.bas` `WBSSampleShowcase_SJIS.bas` `SetupWizard_SJIS.bas` `HierarchyColor_SJIS.bas` をインポート
4. Lite版は `InazumaGantt_Lite_SJIS.bas` `WBSRoadmapReport_Lite_SJIS.bas` `WBSSampleShowcase_Lite_SJIS.bas` `SetupWizard_Lite_SJIS.bas` `HierarchyColor_Lite_SJIS.bas` をインポート
5. `Alt + F8 -> RunSetupWizard` を実行
6. 標準版は `InazumaGantt_v3` シートモジュールへ `SheetModule_UTF8.bas` の内容を貼り付ける
7. Lite版は `InazumaGantt_Lite` シートモジュールへ `SheetModule_Lite_UTF8.bas` の内容を貼り付ける

## 運用ルール

- 日常編集は `_UTF8.bas` を更新します。
- 配布や手動インポートの前に `FixEncoding.ps1` で `_SJIS.bas` を再生成します。
- 標準モジュールは `_SJIS.bas` をインポートし、シートモジュールは `SheetModule_UTF8.bas` / `SheetModule_Lite_UTF8.bas` の内容を貼り付けます。
- `高速入力 ON` では Excel 標準 `Ctrl+Z` を優先するため、`Worksheet_Change` ベースの `LV` / `No.` / 親集計 / 色分け / ガント更新を止めます。`高速入力 OFF` または `ガント更新` 実行時にまとめて再整合します。
- `高速入力 OFF` では、すべての計算と見た目をその場で更新します。
- `Worksheet_Change` は処理冒頭で `EnableEvents` / `Calculation` / `ScreenUpdating` を退避し、エラー時も退避済み状態へ戻します。高速入力フラグと Excel イベント状態がずれた場合は通常モードへ修復します。
- `CreateWbsBackupSheet` は一時シートへ完全コピーしてから `WBS_Backup_v3` / `WBS_Backup_Lite` へ差し替えます。
- `RestoreWbsFromBackupSheet` は復元前の本 WBS を一時退避し、復元途中の失敗時は直前状態へロールバックします。
- C-F の単一セル文言編集では、タスク有無が変わらない限り全体 `RenumberRowsForWorksheet` を抑制します。複数セル貼り付けやタスク有無の変化では全体再採番します。
- `C:F` のタスク名を消しても、タスク詳細・状況・進捗率・担当・開発LT・予定日・実績日が残っていれば、その行は有効行として扱います。
- 補助情報だけ残った行は `C` 列に `（補助情報のみ）` を表示し、親集計と `WBSサマリ` でも正式な有効行として集計します。
- `LV` は見た目上は空欄ですが、補助情報保持行は直前の階層を内部的に維持して親子関係と集計を継続します。完全空行になった時点で `LV/No./補助表示` をクリアします。
