# Excel-gantt-with-inazuma 辛口コードレビュー（再レビュー）

対象: `vba/**/*.bas`, `vba/addons/**`, `BuildInazumaGantt*.ps1`, `FixEncoding.ps1`, `README.md`  
レビュー日: 2026-01-17

## 重大/高
- **[高] イベント再入で無限ループ/重複処理の危険**
  - `vba/SheetModule_UTF8.bas:95-130` と `vba/InazumaGantt_v2_UTF8.bas:922-966`
  - `Worksheet_Change` で `Application.EnableEvents = False` にした直後、`AutoDetectTaskLevel` が **勝手に True に戻す** ため、直後のセル更新が再度イベント発火し得ます。大量貼り付け時に再帰的に重くなる/無限ループ化の可能性。
  - 改善案: `AutoDetectTaskLevel` 側でイベント制御をしない（呼び出し側で一括制御）か、引数で制御する。

- **[高] UTF8ビルドでシートモジュールが文字化け**
  - `BuildInazumaGantt_UTF8.ps1:72-79`
  - `Get-Content -Encoding Default` のままなので、UTF-8の日本語が破壊される環境が残っています。
  - 改善案: `-Encoding UTF8` を明示（または `ReadAllText(...,[Text.Encoding]::UTF8)`）。

- **[高] 移管列の列番号がActiveSheet依存**
  - `vba/addons/DataMigration/DataMigrationWizard_UTF8.bas:107-109`
  - `Range(config.WBSColumn & "1")` が未修飾のため、アクティブシートが移管元以外だと **誤列で読み取り** ます。
  - 改善案: `oldSheet.Range(...)` に限定。

## 中
- **[中] ActiveSheet / Columns の未修飾参照が広範囲**
  - 例: `vba/InazumaGantt_v2_UTF8.bas:546,810,1047,1103` / `vba/HierarchyColor_UTF8.bas:32`
  - 影響: 実行中に別シートがアクティブになると、**別シートを破壊** します。
  - 改善案: すべて `ws.Columns`/`ThisWorkbook.Worksheets(...)` に統一。

- **[中] 「祝日マスタ」前提の文言が残存**
  - `vba/SetupWizard_UTF8.bas:69-73` / `README.md:7,41-43`
  - 実装は `設定マスタ` 内の祝日欄なのに、独立した「祝日マスタ」シートを前提に案内しています。ユーザーを確実に迷わせます。
  - 改善案: 文言と実装を一致させる（独立シートを作るか、説明を直す）。

## 低
- **[低] SheetModule に Option Explicit が無い**
  - `vba/SheetModule_UTF8.bas:1-24`
  - ミスタイプ検出不能。保守性が落ちます。

- **[低] 列リストがAA-AZで頭打ち**
  - `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas:470-483`
  - BA列以降のシートは移管設定できません。

- **[低] キャンセル時の無駄な二重クリア**
  - `vba/InazumaGantt_v2_UTF8.bas:150-157`
  - `Clear` が2回連続で呼ばれています（動作上の害は小さいが、品質が雑）。

- **[低] README の記法崩れと構成ミスマッチ**
  - `README.md:19-25,67-69`
  - 箇条書きの記法が壊れており、`vba/統合版` も実在しません。

## パフォーマンス
- **貼り付け時のO(n^2)化**
  - `vba/SheetModule_UTF8.bas:95-130` / `:255-272`
  - セルごとに `GetNextNo` が全走査され、貼り付けで極端に遅くなる。

- **シェイプ乱造による描画コスト**
  - `vba/InazumaGantt_v2_UTF8.bas:568-737`
  - 行数が増えると描画が重く、Excelが不安定化。

## テスト不足
- 自動テストなし。最低限の手動確認:
  1) 大量貼り付け時の無限ループ/重複更新の有無  
  2) UTF8ビルドで日本語が崩れないか  
  3) 移管ウィザードでアクティブシートを変えても正しく移管されるか  

## 確認したい前提
- すべてのマクロは「必ず InazumaGantt_v2 をアクティブにしてから実行」前提ですか？
