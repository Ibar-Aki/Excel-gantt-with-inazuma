# Excel-gantt-with-inazuma 辛口徹底コードレビュー（再レビュー）

対象: `vba/**/*.bas`, `vba/addons/**`, `BuildInazumaGantt*.ps1`, `FixEncoding.ps1`, `README.md`, `vba/README.md`  
レビュー日: 2026-01-18

## 重大/高
- **[高] UTF8ビルドでSheetModuleの文字化け/破損リスク**  
  - `BuildInazumaGantt_UTF8.ps1:76`  
  - `Get-Content -Encoding Default` でUTF-8を読んでおり、環境依存で日本語が崩れる可能性。結果としてビルドされた`xlsm`のシートモジュールが破損し得ます。  
  - 対応: `-Encoding UTF8` か `[IO.File]::ReadAllText(...,[Text.Encoding]::UTF8)` に固定。

## 中
- **[中] AutoFilterが日付行に付与され、ヘッダーにフィルタが出ない**  
  - `vba/InazumaGantt_v3_UTF8.bas:208`  
  - `ROW_HEADER(8)` ではなく `ROW_DATE_HEADER(7)` にAutoFilterを設定しているため、A-N列のヘッダーにフィルタが出ません。  
  - 対応: `ROW_HEADER` に変更。

- **[中] 進捗率が文字列のときバー描画が0%扱い**  
  - `vba/InazumaGantt_v3_UTF8.bas:607`  
  - `IsNumeric`のみで判定しているため、`"30%"` など文字列入力は0%扱いになります。`Worksheet_Change` は文字列も許容しているので不整合。  
  - 対応: `%`除去→数値化、もしくは入力時に数値へ正規化。

- **[中] ActiveSheet依存が広範囲で誤シート破壊の恐れ**  
  - 例: `vba/InazumaGantt_v3_UTF8.bas:67,547,812,1079` / `vba/HierarchyColor_UTF8.bas:24` / `vba/addons/DataMigration/DataMigration_UTF8.bas:28`  
  - 実行中に別シートがアクティブになると、別シートへの描画・上書きが起き得ます。  
  - 対応: `ThisWorkbook.Worksheets(...)` か引数`ws`で固定。

- **[中] 移管処理が計算モードを強制的に自動へ戻す**  
  - `vba/addons/DataMigration/DataMigrationWizard_UTF8.bas:99`  
  - `vba/addons/DataMigration/DataMigration_UTF8.bas:55`  
  - 元の計算設定(手動/自動)を破壊します。  
  - 対応: 開始時に保存→終了時に復元。

## 低
- **[低] SheetModuleにOption Explicitなし**  
  - `vba/SheetModule_UTF8.bas:1`  
  - 変数タイポが検出されず、保守性が落ちます。

- **[低] RegenerateDateHeadersが全域 `On Error Resume Next` で失敗を隠蔽**  
  - `vba/InazumaGantt_v3_UTF8.bas:1195`  
  - ヘッダー生成失敗に気づけません。

- **[低] 移管ウィザードの列リストがAZ止まり**  
  - `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas:470`  
  - BA列以降が選べません。

- **[低] ドキュメントがv3と不整合**  
  - `README.md:7` / `README.md:19` / `README.md:42` / `vba/README.md:3` / `vba/README.md:21`  
  - 祝日マスタ表記やv2ファイル名が残存、箇条書きの記法崩れ。

- **[低] ビルド(SJIS)のエンコーディングがOS依存**  
  - `BuildInazumaGantt.ps1:62`  
  - `-Encoding Default` は非日本語環境で崩れる可能性あり。

## パフォーマンス
- **貼り付け時のO(n^2)化**  
  - `vba/SheetModule_UTF8.bas:113` / `vba/SheetModule_UTF8.bas:255`  
  - `GetNextNo` が列全走査で、複数セル貼り付け時に極端に遅くなります。  
  - 対応: 変更範囲で最大Noを1回取得し、インクリメント採番。

- **シェイプ乱造による描画コスト**  
  - `vba/InazumaGantt_v3_UTF8.bas:547`  
  - 行数が増えると描画が重く、ファイル肥大・Excel不安定化の原因になります。

## テスト不足
- 自動テストなし。最低限の手動確認:  
  1) 進捗率を`"30%"`文字列で入力 → バーが期待通り描画されるか  
  2) 100行以上の大量貼り付けで固まらないか  
  3) UTF8ビルドでSheetModuleの日本語が化けないか  
  4) 移管後に計算モードが元に戻るか

## 確認したい前提
- すべてのマクロは「必ず InazumaGantt_v3 をアクティブにして実行」前提ですか？
- 進捗率は数値(0-1)のみ想定ですか？ 文字列入力も許容しますか？
