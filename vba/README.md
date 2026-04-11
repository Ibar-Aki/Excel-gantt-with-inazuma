# VBAモジュール

更新日: 2026-04-11

InazumaGantt v3 で使用する VBA モジュール一覧です。

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
| `WBSRoadmapReport_SJIS.bas` | WBSロードマップ |
| `WBSSampleShowcase_SJIS.bas` | ShowcaseサンプルWBS生成 |
| `SetupWizard_SJIS.bas` | セットアップウィザード |
| `HierarchyColor_SJIS.bas` | 階層色分け |
| `SheetModule_SJIS.bas` | シートイベントコード |

## インポート手順

1. Excel ファイルを開く
2. `Alt + F11` で VBA エディタを開く
3. `InazumaGantt_v3_SJIS.bas` `WBSRoadmapReport_SJIS.bas` `WBSSampleShowcase_SJIS.bas` `SetupWizard_SJIS.bas` `HierarchyColor_SJIS.bas` をインポート
4. `Alt + F8 -> RunSetupWizard` を実行
5. `InazumaGantt_v3` シートモジュールへ `SheetModule_SJIS.bas` の内容を貼り付ける

## 運用ルール

- 日常編集は `_UTF8.bas` を更新します。
- Excel 取り込み前に `FixEncoding.ps1` で `_SJIS.bas` を再生成します。
- `SheetModule_SJIS.bas` は標準モジュールではなく、`InazumaGantt_v3` シートモジュールに貼り付けます。
