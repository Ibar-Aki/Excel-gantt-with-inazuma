# VBAモジュール

InazumaGantt v2.2 で使用するVBAモジュールです。

## エンコーディングについて

各ファイルは2つのバージョンがあります：

| サフィックス | エンコーディング | 用途 |
|-------------|-----------------|------|
| `_SJIS.bas` | Shift-JIS (CP932) | **Excelにインポート用** |
| `_UTF8.bas` | UTF-8 (BOMなし) | 編集・Git管理用 |

> **重要**: Excelにインポートする場合は必ず `_SJIS.bas` を使用してください。  
> `_UTF8.bas` をインポートすると文字化けします。

## ファイル一覧

### 必須モジュール

| ファイル | 用途 |
|----------|------|
| `InazumaGantt_v2_SJIS.bas` | メイン機能 |
| `HierarchyColor_SJIS.bas` | 階層色分け |
| `SetupWizard_SJIS.bas` | セットアップウィザード（推奨） |
| `SheetModule_SJIS.bas` | シートイベント（※） |
| `DataMigration_SJIS.bas` | データ移管（任意） |

> **※ SheetModule について**  
> このファイルは「標準モジュール」ではなく、シートモジュールに貼り付けます。

## 統合版について

`統合版/` フォルダには、全モジュールを1つに統合したバージョンがあります。

| ファイル | 用途 |
|----------|------|
| `InazumaGantt_Integrated_SJIS.bas` | 統合版メインモジュール |
| `SheetModule_Integrated_SJIS.bas` | 統合版用シートモジュール |

> **注意**: 統合版を使用する場合は `SheetModule_Integrated_SJIS.bas` を使用してください。

## インポート手順

1. Excelファイルを開く
2. `Alt + F11` でVBAエディタを開く
3. ファイル → ファイルのインポート
4. `InazumaGantt_v2_SJIS.bas` を選択
5. `HierarchyColor_SJIS.bas` を選択
6. `SetupWizard_SJIS.bas` を選択
7. （任意）`DataMigration_SJIS.bas` を選択

## シートモジュールの設定

1. VBAエディタで「InazumaGantt_v2」シートをダブルクリック
2. `SheetModule_SJIS.bas` の内容を全てコピー＆貼り付け
3. 保存して閉じる

## 追加モジュール

開発者向けの追加モジュールは `dev/extra_modules/` にあります。
