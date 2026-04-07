# InazumaGantt v3 アーキテクチャ

更新日: 2026-04-07

InazumaGantt v3 の実行構成と責務分担をまとめた開発者向けメモです。

## 構成要素

現行の実行系は次の 4 要素です。

| 要素 | 役割 |
|------|------|
| `InazumaGantt_v3` | シート初期化、ガント描画、設定参照、補助マクロ |
| `SetupWizard` | 初回セットアップとサンプルデータ投入 |
| `HierarchyColor` | 階層色分け用の条件付き書式設定 |
| `SheetModule` | シートイベント処理 |

`設定マスタ` は独立モジュールではなく、`InazumaGantt_v3.EnsureSettingsSheet` が作成する補助シートです。祝日入力欄もこのシートに含まれます。

## 主要シート

| シート名 | 用途 |
|----------|------|
| `InazumaGantt_v3` | メイン入力・描画シート |
| `設定マスタ` | ダブルクリック設定と祝日一覧 |
| `InazumaGantt_説明` | 操作ガイド |

## 依存関係

```text
ユーザー操作
  -> SheetModule
     -> InazumaGantt_v3.AutoDetectTaskLevel
     -> InazumaGantt_v3.ValidateProgressInput
     -> InazumaGantt_v3.ValidateDateInput
     -> InazumaGantt_v3.ToggleTaskCollapse

セットアップ
  -> SetupWizard
     -> InazumaGantt_v3.SetupInazumaGantt
     -> InazumaGantt_v3.EnsureSettingsSheet
     -> HierarchyColor.SetupHierarchyColors
     -> InazumaGantt_v3.RefreshInazumaGantt

描画更新
  -> InazumaGantt_v3.RefreshInazumaGantt
     -> RegenerateDateHeaders
     -> ApplyGanttBorders / ApplyWeekendColors / ApplyHolidayColors
     -> DrawGanttBars
```

## 実装上の前提

- 公開マクロは原則として `InazumaGantt_v3` シートを対象に処理します。
- `ShiftDates` は選択セルを使うため、`InazumaGantt_v3` シートを表示した状態で実行します。
- 進捗率は `0.7` `70` `70%` を受け付け、内部的には 0〜1 に正規化します。
- `RefreshInazumaGantt` はヘッダー再生成、休日色反映、ガント再描画をまとめて実行します。

## データの流れ

### タスク入力

1. `SheetModule.Worksheet_Change` が C-F 列の変更を検知
2. `AutoDetectTaskLevel` が LV を再判定
3. B列の No.、H列の状況、I列の進捗率初期値を補完

### 進捗率入力

1. `Worksheet_Change` が I 列の変更を検知
2. `ValidateProgressInput` が入力値を検証して 0〜1 に正規化
3. `UpdateStatusByProgress` が状況列を更新

### 日付入力

1. `Worksheet_Change` が K-L 列の変更を検知
2. `ValidateDateInput` が日付形式と開始終了の前後関係を確認
3. 土日祝の場合のみ確認ダイアログを表示

### ガント更新

1. `RefreshInazumaGantt` が対象シートを解決
2. 日付ヘッダー、罫線、土日祝色を再構築
3. `DrawGanttBars` が予定バー、進捗バー、実績バー、今日線、イナズマ線を描画

## 保守メモ

- 開発時は `_UTF8.bas` を編集し、Excel 取り込み用は `FixEncoding.ps1` で `_SJIS.bas` を再生成します。
- `SheetModule_SJIS.bas` は標準モジュールへインポートせず、`InazumaGantt_v3` シートモジュールへ貼り付けます。
- 廃止中の移管機能は現行アーキテクチャの対象外です。
