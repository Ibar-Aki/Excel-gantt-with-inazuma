# LLM修正指示書（InazumaGantt v2.2）

目的: 既存の不具合/リスクを解消し、配布品質と操作の安全性を上げる。  
注意: この作業ではビルドは実行しない（修正のみ）。  

## 更新メモ（2026-01-14時点）
- `output/` に配布物あり（`output/InazumaGantt_v2.2_20260114_0025.xlsm`）。READMEの「output/ ← ビルド済みExcelファイル」と整合。
- 未解決の課題は下記の修正対象一覧。

## 修正対象一覧（優先度順）

### P1: ガントバーが範囲外開始タスクを描画しない
- 事象: 開始日がガント表示範囲より前だと、タスクが完全に描画されない。
- 参照: `vba/InazumaGantt_v2_UTF8.bas` / `vba/統合版/InazumaGantt_Integrated_UTF8.bas` の `DrawGanttBars`（`startCol >= ganttStartCol` の条件）
- 修正方針:
  - `startCol` が範囲外でも、表示範囲内にかかる部分は描画する。
  - 例: `startCol < ganttStartCol` の場合は `startCol = ganttStartCol` にクランプする。
  - `endCol` も表示範囲にクランプする。

### P1: 既存図形の誤削除リスク
- 事象: `Bar_`/`Today_`/`Inazuma_` プレフィックスの図形を無条件削除。
- 参照: `vba/InazumaGantt_v2_UTF8.bas` / `vba/統合版/InazumaGantt_Integrated_UTF8.bas` の `DrawGanttBars`
- 修正方針:
  - 図形名の完全一致リストで削除対象を限定する、または
  - 生成時に一意なタグを持たせ、タグ付きのみ削除。

### P2: 計算モードの強制変更（元設定が戻らない）
- 事象: `Application.Calculation` を `xlCalculationManual` にした後、常に `xlCalculationAutomatic` へ戻すためユーザー設定が破壊される。
- 参照: `InazumaGantt_v2_UTF8.bas`（`DrawGanttBars`, `RefreshInazumaGantt`, `SetupInazumaGantt`, `ResetFormatting`）、
        `DataMigration_UTF8.bas`、`HierarchyColor_UTF8.bas`、統合版同名処理
- 修正方針:
  - 実行前の状態を保存して、最後に元に戻す。
  - 例: `prevCalc = Application.Calculation` → 処理 → `Application.Calculation = prevCalc`

### P2: ScreenUpdatingの復元漏れ（エラー時）
- 事象: エラー時に `ScreenUpdating` が復元されない。
- 参照: `vba/SetupWizard_UTF8.bas`
- 修正方針:
  - `On Error GoTo` のハンドラで必ず復元する。

### P2: DataMigrationが既存シートに追記する
- 事象: 「新規作成」と表示しつつ、既存シートがあるとそこに追記。
- 参照: `vba/DataMigration_UTF8.bas`
- 修正方針:
  - 既存シートがある場合は「上書き/新規/キャンセル」を選択させる。
  - 上書きならクリアしてから移管。

### P2: 日付シフトがSelection依存
- 事象: `ShiftDates` が選択範囲に依存し、異なるシート・非セル選択で誤動作。
- 参照: `vba/InazumaGantt_v2_UTF8.bas` の `ShiftDates`
- 修正方針:
  - 対象範囲を明示（K:Nなど）または入力ボックスで範囲指定。
  - `Selection` の型/シート検証を追加。

### P3: PDF出力がPageSetupを戻さない
- 事象: 既存印刷設定を破壊。
- 参照: `vba/InazumaGantt_v2_UTF8.bas` / `vba/統合版/InazumaGantt_Integrated_UTF8.bas` の `ExportToPDF`
- 修正方針:
  - 変更前の `PageSetup` を保存 → 復元。

### P3: ドキュメント/ガイド文と実装の不整合
- 事象: ガイド文（シート内ガイド/開発メモ等）に「A列 or B列ダブルクリックで完了」とあるが、実装はB列のみ。
- 参照: `vba/InazumaGantt_v2_UTF8.bas`（ガイド文）、`dev/docs/InazumaGantt_v2_SheetModule_README.md` と `vba/SheetModule_UTF8.bas`
- 修正方針:
  - 実装をガイドに合わせるか、ガイド文を修正。

### P3: 祝日チェックが線形走査で遅い
- 事象: 貼り付け時に行数×祝日数で遅延。
- 参照: `vba/SheetModule_UTF8.bas` の `CheckHoliday`
- 修正方針:
  - 祝日をDictionaryや配列でキャッシュして高速化。

## 修正後の期待結果（最低限）
- ガント描画で「範囲外開始だが範囲内にかかるタスク」が描画される。
- マクロ実行前の `Calculation` / `ScreenUpdating` が復元される。
- 既存図形の誤削除が起こらない。
- DataMigrationがデータ混在を起こさない。

## 関連ファイル
- `README.md`
- `vba/InazumaGantt_v2_UTF8.bas`
- `vba/統合版/InazumaGantt_Integrated_UTF8.bas`
- `vba/SheetModule_UTF8.bas`
- `vba/統合版/SheetModule_Integrated_UTF8.bas`
- `vba/SetupWizard_UTF8.bas`
- `vba/DataMigration_UTF8.bas`
- `vba/HierarchyColor_UTF8.bas`
