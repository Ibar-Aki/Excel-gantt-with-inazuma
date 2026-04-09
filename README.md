# InazumaGantt v3

更新日: 2026-04-10

Excel ベースのイナズマガントチャート管理ツールです。

> [!NOTE]
> Excel ファイル名は自由に変更できます。
> シート名 `InazumaGantt_v3` `設定マスタ` `InazumaGantt_説明` は変更しないでください。

## クイックスタート

### 1. VBA モジュールをインポート

Excel に取り込むのは `vba/` 配下の `_SJIS.bas` です。

```text
Alt + F11 -> ファイル -> ファイルのインポート
```

- `InazumaGantt_v3_SJIS.bas` : メインロジック
- `WBSOverviewReports_SJIS.bas` : 俯瞰レポート生成
- `SetupWizard_SJIS.bas` : セットアップ機能
- `HierarchyColor_SJIS.bas` : 階層色分け機能

`_UTF8.bas` は編集・Git 管理用です。Excel にはインポートしません。

### 2. セットアップウィザードを実行

```text
Alt + F8 -> RunSetupWizard -> 実行
```

ウィザードで以下をまとめて作成します。

- メインシート `InazumaGantt_v3`
- 設定マスタシート `設定マスタ`（祝日入力欄を含む）
- 階層色分け
- ガントチャートの初期描画

### 3. シートモジュールを設定

セットアップ完了後に、`SheetModule_SJIS.bas` を `InazumaGantt_v3` シートモジュールへ貼り付けます。

1. VBA エディタで `InazumaGantt_v3` シートをダブルクリック
2. `vba/SheetModule_SJIS.bas` の内容を貼り付け
3. 保存して閉じる

## 主な機能

| 機能 | 説明 |
|------|------|
| ガントチャート | 予定バー、進捗バー、実績バー、今日線を描画 |
| イナズマ線 | 今日基準で進捗位置を折れ線表示 |
| 階層色分け | LV1-LV4 を条件付き書式で色分け |
| ダブルクリック完了 | B列ダブルクリックで進捗率 100% と完了状態を反映 |
| 折りたたみ | Shift + 右クリックで LV1 配下を表示・非表示 |
| 日付シフト | 選択した日付を営業日単位で一括シフト |
| PDF 出力 | 当月末までのガントを PDF 化 |
| 全体サマリ | LV1 フェーズ単位で進捗・工数・遅延を一覧化 |
| ロードマップ | 月次に圧縮した全体計画シートを生成 |

## よく使うマクロ

| マクロ | 用途 |
|--------|------|
| `RunSetupWizard` | 初回セットアップ |
| `RefreshInazumaGantt` | ガント再描画 |
| `ResetFormatting` | ヘッダー、罫線、休日色の再構築 |
| `ToggleWeekends` | 土日列の表示切り替え |
| `ExportToPDF` | PDF 出力 |
| `CreatePhaseSummarySheet` | フェーズ単位の全体サマリを生成 |
| `CreateRoadmapOverviewSheet` | 月次ロードマップを生成 |

## ファイル構成

```text
vba/   VBA モジュール本体（Excel には _SJIS.bas をインポート）
docs/  利用者向けドキュメント
dev/   開発者向けドキュメント
```

## ドキュメント

### 利用者向け

| ファイル | 内容 |
|----------|------|
| [docs/利用者ガイド.md](docs/利用者ガイド.md) | 操作マニュアル |
| [docs/FEATURES.md](docs/FEATURES.md) | 機能詳細 |
| [docs/CUSTOMIZE.md](docs/CUSTOMIZE.md) | カスタマイズ方法 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 問題解決 |
| [CHANGELOG.md](CHANGELOG.md) | 更新履歴 |

### 開発者向け

| ファイル | 内容 |
|----------|------|
| [dev/docs/SETUP.md](dev/docs/SETUP.md) | セットアップ詳細 |
| [dev/docs/ARCHITECTURE.md](dev/docs/ARCHITECTURE.md) | 構成と責務分担 |
| [vba/README.md](vba/README.md) | VBA モジュール説明 |

## ライセンス

MIT License - [LICENSE](LICENSE)
