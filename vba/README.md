# VBAモジュール

InazumaGantt v3 で使用するVBAモジュールです。

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
| `InazumaGantt_v3_SJIS.bas` | メイン機能 |
| `HierarchyColor_SJIS.bas` | 階層色分け |
| `SetupWizard_SJIS.bas` | セットアップウィザード |
| `SheetModule_SJIS.bas` | シートイベント（※） |

> **※ SheetModule について**  
> このファイルは「標準モジュール」ではなく、シートモジュールに貼り付けます。

## インポート手順

1. Excelファイルを開く
2. `Alt + F11` でVBAエディタを開く
3. ファイル → ファイルのインポート
4. `InazumaGantt_v3_SJIS.bas` を選択
5. `HierarchyColor_SJIS.bas` を選択
6. `SetupWizard_SJIS.bas` を選択

## シートモジュールの設定

> **注意**: シートモジュールはセットアップウィザード実行後に設定してください。

1. VBAエディタで「InazumaGantt_v3」シートをダブルクリック
2. `SheetModule_SJIS.bas` の内容を全てコピー＆貼り付け
3. 保存して閉じる

## 追加モジュール

開発者向けの追加モジュールは `dev/extra_modules/` にあります。
