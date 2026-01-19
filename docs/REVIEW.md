# Excel-gantt-with-inazuma 辛口徹底コードレビュー（再レビュー）

対象: `vba/**/*.bas`, `vba/addons/**`, `BuildInazumaGantt*.ps1`, `FixEncoding.ps1`, `README.md`, `vba/README.md`  
レビュー日: 2026-01-19

## 重大/高
- 該当なし（現時点）

## 中
- **[中] AutoFilterが日付行に付与され、ヘッダーにフィルタが出ない**  
  `vba/InazumaGantt_v3_UTF8.bas:210`
- **[中] ActiveSheet依存が残存し、誤シート更新の危険が残る**  
  `vba/InazumaGantt_v3_UTF8.bas:68` / `vba/HierarchyColor_UTF8.bas:32` / `vba/addons/DataMigration/DataMigration_UTF8.bas:30`
- **[中] SJISビルドのシートモジュール注入がOS依存（非JP環境で文字化けの恐れ）**  
  `BuildInazumaGantt.ps1:81`

## 低
- **[低] RegenerateDateHeadersが全面的にエラー無視で失敗を隠蔽**  
  `vba/InazumaGantt_v3_UTF8.bas:1229`
- **[低] 移管ウィザードの列リストがAZ止まり**  
  `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas:470`
- **[低] ドキュメントがv3と不整合・記法崩れ**  
  `README.md:7` / `README.md:19` / `README.md:42` / `vba/README.md:3` / `vba/README.md:23`

## パフォーマンス/UX
- **貼り付け時のO(n^2)化（GetNextNoが全走査）**  
  `vba/SheetModule_UTF8.bas:113` / `vba/SheetModule_UTF8.bas:255`
- **シェイプ乱造で描画コスト増大**  
  `vba/InazumaGantt_v3_UTF8.bas:547`
- **土日祝入力の確認がセル単位で大量ポップアップ**  
  `vba/SheetModule_UTF8.bas:168`

## テスト不足
- AutoFilterがヘッダーに表示されるか（A-N列）
- SJISビルドを非日本語OSで実行して文字化けが出ないか
- 100行以上の貼り付けでフリーズしないか
- K/L列に大量貼り付け時の確認ダイアログが許容範囲か

## 確認したい前提
- すべてのマクロは「必ず InazumaGantt_v3 をアクティブにして実行」前提ですか？
- SJISビルドは日本語Windowsのみで運用する想定ですか？
