# InazumaGantt v3

更新日: 2026-05-05

Excel ベースのイナズマガントチャート管理ツールです。現在は標準版 `v3` と簡易版 `Lite` を用意しています。

> [!NOTE]
> Excel ファイル名は自由に変更できます。
> シート名 `InazumaGantt_v3` / `InazumaGantt_Lite` `設定マスタ` `InazumaGantt_説明` は変更しないでください。

## バージョン構成

| 版 | 用途 |
|----|------|
| `v3` | 開始実績 / 完了実績を持つ標準版 |
| `Lite` | 実績列を持たない簡易版。ガントは中央 1 本のみ |

## 最近の運用改善

- `H:状況` と `I:進捗率` は相互同期します。進捗率を変えると状況が追従し、状況を直接変えると選択した状況を優先して進捗率を補正します。
- `K:開発LT` は内部的には数値で保持し、表示だけ `h` 付きになります。
- `一括編集 ON/OFF` は Ctrl+Z 優先の高速入力モードです。ON 中は Excel 標準 Undo を守るため自動更新を止め、OFF 復帰時にまとめて再整合します。仕様詳細は `docs/高速入力_状況_進捗率仕様レポート.md` を参照してください。
- 高速入力状態は、Excel 再起動や VBA リセット後に `設定マスタ` と `Application.EnableEvents` がずれても、主要マクロ入口とシート変更時に通常モードへ自動修復します。
- `WBS退避` / `バックアップ復元` ボタンを追加し、見出し付き `A:O` 範囲の値・書式・入力規則を最新 1 枚のコピーシートへ退避・復元できるようにしました。
- `WBS退避` は一時シートに完全コピーしてから最新バックアップへ差し替えるため、失敗時に既存バックアップを壊しにくい構成です。
- `バックアップ復元` は更新日時付きの確認ダイアログを出してから本体へ反映します。復元中に失敗した場合は直前の本 WBS へロールバックします。
- `ガント更新` に再描画と基本書式の再構築を統合し、旧 `書式リセット` は互換用マクロとして残します。
- WBSサマリは設定マスタの `WBSサマリ表示階層` で `LV1のみ` / `LV2まで` を切り替えできます。`LV2まで` では LV2 行を折りたたみ可能な詳細行として表示します。
- 通常入力時の親集計と警告更新は、変更行と祖先チェーンだけを再計算して全体走査を避けます。C-F の単一セル文言編集では、タスク有無が変わらない限り全体再採番も抑制します。
- ビルド時の smoke test で作成される `WBS_Backup_*` は保存前に削除されるため、配布ブックには利用者作成前の古いバックアップを含めません。

## クイックスタート

### 1. VBA モジュールをインポート

Excel に取り込むのは `vba/` 配下の `_SJIS.bas` です。

```text
Alt + F11 -> ファイル -> ファイルのインポート
```

- `InazumaGantt_v3_SJIS.bas` : メインロジック
- `InazumaGantt_Lite_SJIS.bas` : Lite 版メインロジック
- `WBSRoadmapReport_SJIS.bas` : WBSサマリ生成
- `WBSRoadmapReport_Lite_SJIS.bas` : Lite 版 WBSサマリ生成
- `WBSSampleShowcase_SJIS.bas` : 大規模サンプルWBS生成
- `WBSSampleShowcase_Lite_SJIS.bas` : Lite 版サンプルWBS生成
- `SetupWizard_SJIS.bas` : セットアップ機能
- `SetupWizard_Lite_SJIS.bas` : Lite 版セットアップ機能
- `HierarchyColor_SJIS.bas` : 階層色分け機能
- `HierarchyColor_Lite_SJIS.bas` : Lite 版階層色分け機能

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

セットアップ完了後に、標準版は `SheetModule_SJIS.bas` を `InazumaGantt_v3` シートモジュールへ、Lite 版は `SheetModule_Lite_SJIS.bas` を `InazumaGantt_Lite` シートモジュールへ貼り付けます。

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
| 高速入力 | Ctrl+Z を優先し、ON 中は自動再計算と再描画を停止。状態不整合は通常モードへ自動修復 |
| WBSバックアップ | 見出し付き `A:O` を最新 1 枚のコピーシートへ退避 / 復元。退避は一時シート差し替え、復元はロールバック付き |
| 日付シフト | 選択した日付を営業日単位で一括シフト |
| PDF 出力 | 当月末までのガントを PDF 化 |
| WBSサマリ | 月次に圧縮した全体計画シートを生成。設定で LV2 詳細行も折りたたみ表示 |
| Showcase Sample | 約150行の見栄え重視サンプルWBSを生成 |
| Lite版 | 実績列なし / 単線ガントの簡易運用版 |

## よく使うマクロ

| マクロ | 用途 |
|--------|------|
| `RunSetupWizard` | 初回セットアップ |
| `RefreshInazumaGantt` | ガント、親集計、警告表示、ヘッダー、罫線、土日祝色をまとめて更新 |
| `ResetFormatting` | 旧手順互換。内部では `RefreshInazumaGantt` を実行 |
| `ToggleBulkEditMode` | Ctrl+Z 優先の高速入力 ON / OFF |
| `CreateWbsBackupSheet` | 見出し付き `A:O` を一時シート経由で最新バックアップシートへ退避 |
| `RestoreWbsFromBackupSheet` | バックアップシートから `A:O` をロールバック付きで復元し、通常モードで再整合 |
| `ToggleWeekends` | 土日列の表示切り替え |
| `ExportToPDF` | PDF 出力 |
| `CreateRoadmapOverviewSheet` | WBSサマリを生成 |
| `CreateShowcaseSampleWBS` | 約150行のサンプルWBSを生成 |

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
| [docs/新規機能一覧.md](docs/新規機能一覧.md) | 直近追加した機能の一覧 |
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
