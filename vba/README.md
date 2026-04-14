# VBAモジュール

更新日: 2026-04-15

InazumaGantt v3 / Lite で使用する VBA モジュール一覧です。

## エンコーディング

| サフィックス | エンコーディング | 用途 |
|-------------|-----------------|------|
| `_SJIS.bas` | Shift-JIS (CP932) | Excel にインポート |
| `_UTF8.bas` | UTF-8 (BOM なし) | 編集・Git 管理 |

> Excel に取り込むときは必ず `_SJIS.bas` を使ってください。
> `_UTF8.bas` を直接インポートすると文字化けします。

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
6. 標準版は `InazumaGantt_v3` シートモジュールへ `SheetModule_SJIS.bas` を貼り付ける
7. Lite版は `InazumaGantt_Lite` シートモジュールへ `SheetModule_Lite_SJIS.bas` を貼り付ける

## 運用ルール

- 日常編集は `_UTF8.bas` を更新します。
- Excel 取り込み前に `FixEncoding.ps1` で `_SJIS.bas` を再生成します。
- `SheetModule_SJIS.bas` は `InazumaGantt_v3` シートモジュール、`SheetModule_Lite_SJIS.bas` は `InazumaGantt_Lite` シートモジュールに貼り付けます。
