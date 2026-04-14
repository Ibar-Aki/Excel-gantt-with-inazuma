# Changelog

更新日: 2026-04-15

All notable changes to InazumaGantt will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.1.4] - 2026-04-11

## [3.2.0] - 2026-04-15

### Added

- **Lite版**: 実績列なし / 単線ガントの `InazumaGantt_Lite` を追加
- **Lite版配布**: `CreateDistributionPackage_Lite.ps1` と Lite 用 bundle / restore 手順書を追加

### Changed

- **配布手順書**: 配布物に同梱する restore 手順書を ASCII 名 `RestoreGuide*.md` でも提供

## [3.1.4] - 2026-04-11

### Changed

- **完了行表示**: 完了時は LV と同じ塗り範囲を薄い灰色で上書きする方式へ統一
- **サマリ専用化**: `WBSRoadmapReport` を追加し、参照サンプルに合わせた `WBSサマリ` 自動生成へ整理
- **サンプル出力**: Showcase サンプルはロードマップのみを同時生成する構成へ変更

### Removed

- **旧俯瞰シート群**: `WBS全体サマリ` `WBSダッシュボード` `WBSフィルタビュー` `完了表現ガイド` を現行機能から外し、`vba/archive/過去の検討機能/` へアーカイブ

## [3.1.2] - 2026-04-11

### Added

- **Showcase Sample WBS**: 約150行の大規模サンプルWBS生成マクロを追加

## [3.1.1] - 2026-04-10

### Added

- **WBSダッシュボード**: 全体 KPI、担当負荷、要注意タスクを俯瞰する別シート生成マクロを追加
- **WBSフィルタビュー**: 担当、状態、期限観点でフィルタしやすい一覧シート生成マクロを追加

## [3.1.0] - 2026-04-09

### Added

- **WBS全体サマリ**: LV1 フェーズ単位で進捗率、工数、遅延を集約する別シート生成マクロを追加
- **WBSサマリ**: 月次に圧縮した全体計画・進捗を可視化する別シート生成マクロを追加

### Changed

- **初期入力範囲**: 書式、採番、条件付き書式の既定範囲を No.1000 まで拡張
- **モジュール分離**: 俯瞰シート生成マクロを `WBSOverviewReports` 専用モジュールへ分離

## [3.0.1] - 2026-04-08

### Added

- **開発LT列**: `K` 列に `xh` 形式の工数入力列を追加
- **工数ロールアップ**: Shift+右クリックで配下の末端タスク工数を親行へ集計

### Changed

- **列構成**: 予定・実績日付列を `L:O`、ガント開始列を `P` に移動
- **日付検証**: 新しい日付列位置に追従するよう更新
- **サンプルデータ**: 開発LT列を含む構成へ更新

## [3.0.0] - 2026-04-07

### Changed

- モジュール名を `InazumaGantt_v3` に統一
- 全ドキュメントのバージョン表記を v3 に統一
- `FixEncoding.ps1` / `BuildInazumaGantt_UTF8.ps1` からDataMigration参照を削除

### Removed

- **データ移管ウィザード**: `addons/DataMigration/` 配下の全モジュール（DataMigration, WBSParser, DataMigrationWizard, MigrationFormBuilder）
- **旧バージョンアーカイブ**: `dev/archive/` 配下の全ファイル
- **旧dev/docs**: IMPROVEMENT_REPORT, LLM_BEST_PRACTICES, LLM_FIX_BRIEF, DEPENDENCIES, InazumaGantt_v2_SheetModule_README, ganttマクロ改善メモ
- **outputディレクトリ**: 古いビルド成果物29ファイル
- **BuildInazumaGantt.ps1** (SJIS版): UTF-8版に統一
- **hex_dump.txt**: デバッグ残骸

## [2.2.0] - 2026-01-14

### Added

- **設定マスタシート**: ダブルクリック完了の動作を制御（`EnsureSettingsSheet`）
- **タスク折りたたみ**: Shift+右クリックでLV1配下を非表示/表示（`ToggleTaskCollapse`）
- **No.自動採番**: 行挿入後も連番維持（`RenumberRows`）
- **日付一括シフト**: 祝日マスタ考慮の営業日シフト（`ShiftDates`）
- **PDF出力**: 当月末までのガントを含む出力（`ExportToPDF`）
- **日付ヘッダー再生成**: 開始日変更時に自動更新（`RegenerateDateHeaders`）
- **バリデーション機能**: 日付・進捗率の入力チェック

### Changed

- **RefreshInazumaGantt**: 日付ヘッダー再生成と色クリア処理を追加
- **ResetFormatting**: ガント領域の色をクリアしてから再塗り
- **SetupWizard**: 設定マスタシートを自動作成

### Fixed

- **休日色塗り問題**: 開始日変更後も正しく色塗り

## [2.1.0] - 2026-01-05

### Added

- **利用者ガイド**: マクロ知識不要の操作マニュアル
- **グリッド線非表示**: セットアップ時に目盛線をオフ
- **オートフィルター**: 8行目A-N列に自動設定
- **コントロールボタン**: ガント更新、土日切替、書式リセット
- **No.初期採番**: 1〜400を自動入力

### Changed

- **イナズマ線改善**:
  - 今日線は9行目からスタート
  - 完了済み過去タスクは今日の位置で接続
- **今日の日付赤字表示**: 7行目の今日列が赤字に
- **7行目ガント部太字**: O列以降を太字に変更
- **ガント部縦罫線**: C-D間と同じ極細線を適用

### Fixed

- **罫線パターン**: 詳細な罫線サマリに基づく実装
- **ダブルクリック完了**: A/B列のみに制限、完了済みタスクは変更不可

## [2.0.0] - 2026-01-01

### Added

- **階層別タスク入力機能**: C～F列の入力位置で階層レベル（LV1～4）を自動判定
- **進捗率自動更新機能**: 進捗率（I列）の入力で状況（H列）を自動更新
  - 0% → 未着手
  - 1～99% → 進行中
  - 100% → 完了
- **ダブルクリック完了機能**: タスク行をダブルクリックで即完了
  - 進捗率 → 100%
  - 状況 → 完了
  - 完了実績 → 今日の日付（開始実績がある場合）
- **階層別色分け機能**: タスク入力列からN列まで階層別に色塗り
  - LV1 (C列) → サーモン色
  - LV2 (D列) → 薄い青
  - LV3 (E列) → 薄い緑
  - LV4 (F列) → 薄い黄色
- **データ移管機能**: 既存ガントチャート形式からv2形式への自動移管
- **イナズマガントチャート**: 進捗の遅れを視覚化
- **VBAモジュール**:
  - `InazumaGantt_v2.bas` - メイン機能
  - `HierarchyColor.bas` - 階層色分け
  - `DataMigration.bas` - データ移管
  - `InazumaGantt_v2_SheetModule.bas` - シートイベント処理

### Changed

- **列構成をv2形式に変更**:
  - A列: LV（階層レベル、自動設定）
  - B列: No.（通し番号）
  - C～F列: TASK（階層別入力）
  - G列: タスク詳細
  - H～N列: 状況、進捗率、担当、予定・実績日付
  - O列以降: ガントチャート
- **フォルダ構造を整理**:
  - `docs/` - ドキュメント集約
  - `vba_modules/import/` - インポート用SJIS版
  - `vba_modules/source/` - 開発用UTF-8版
  - `旧モデル/` - 旧バージョンアーカイブ

### Fixed

- **日付行と項目行のズレを修正**: 曜日表示をROW_HEADERに移動
- **8行目のデータがガントに表示されない問題を修正**

### Documentation

- README.md - プロジェクト概要とクイックスタート
- SETUP.md - 詳細なセットアップガイド
- docs/InazumaGantt_README.md - 基本機能説明
- docs/HierarchyColor_README.md - 階層色分け機能説明
- docs/DataMigration_README.md - データ移管方法
- docs/TROUBLESHOOTING.md - トラブルシューティング
- docs/CUSTOMIZE.md - カスタマイズガイド

## [1.0.0] - 2025-12-XX (旧モデル)

### Initial Release

- 基本的なガントチャート機能
- イナズマ線描画
- 条件付き書式による進捗バー表示

---

## Version Naming Convention

- **Major** (X.0.0): 互換性のない変更
- **Minor** (x.X.0): 後方互換性のある機能追加
- **Patch** (x.x.X): 後方互換性のあるバグ修正
