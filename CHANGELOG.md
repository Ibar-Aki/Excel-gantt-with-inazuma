# Changelog

All notable changes to InazumaGantt v2 will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
