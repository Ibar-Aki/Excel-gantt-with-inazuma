# Excel-gantt-with-inazuma 辛口コードレビュー

対象: `vba/**/*.bas`, `vba/addons/**`, `BuildInazumaGantt*.ps1`, `FixEncoding.ps1`, `README.md`  
レビュー日: 2026-01-17

## 重大/高
- **[重大] コンパイル不能: 変数の二重宣言**
  - `vba/HierarchyColor_UTF8.bas:37-38,112-113` / `vba/HierarchyColor_SJIS.bas:37-38,112-113`
  - 影響: モジュールの読み込み時点でコンパイルエラーになり、階層色分けが一切動作しません。ビルドスクリプト経由でも失敗します。
  - 改善案: 重複宣言を削除（`Dim lastRow As Long` を1つに統一）。

- **[高] UTF8ビルドでシートモジュールが文字化け**
  - `BuildInazumaGantt_UTF8.ps1:76`
  - 影響: `SheetModule_UTF8.bas` を `-Encoding Default` で読み込み注入しているため、UTF-8前提の日本語が崩れ、環境によってはコード文字列が壊れます。
  - 改善案: `Get-Content -Encoding UTF8 -Raw` に変更（または `ReadAllText` で明示的にUTF-8）。

- **[高] セットアップキャンセル時に計算モードが強制変更**
  - `vba/InazumaGantt_v2_UTF8.bas:148-157`
  - 影響: `prevCalc` を保存しているのに、キャンセル分岐で `xlCalculationAutomatic` を固定的に設定しており、ユーザーの計算モードが意図せず変わります。
  - 改善案: キャンセル分岐も `prevCalc` を復元する。

## 中
- **[中] イナズマ線の座標が未初期化になり得る**
  - `vba/InazumaGantt_v2_UTF8.bas:668-689`
  - 影響: `useTodayPosition=True` かつ「今日がガント範囲外」の場合、`inazumaX` が未設定のまま `0` 扱いになり、不正な位置に線が引かれる可能性。
  - 改善案: 今日が範囲外ならスキップ、または `inazumaX` を安全な位置に初期化。

- **[中] ActiveSheet/未修飾参照が多く、誤シート破壊リスク**
  - 例: `vba/InazumaGantt_v2_UTF8.bas:545,802,372-375,502-505,1038-1042`
  - 影響: 別シートをアクティブにしたまま実行すると、意図しないシートを加工・破壊します。
  - 改善案: `ThisWorkbook.Worksheets(MAIN_SHEET_NAME)` を基準にし、`Columns` などの未修飾参照を `ws.Columns` へ統一。

- **[中] 移管ウィザードの開始行が数値以外で例外**
  - `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas:524-536`
  - 影響: `txtDataStartRow` が数値以外だと `CLng` で落ちます（UI側での入力ガードがない）。
  - 改善案: `IsNumeric` 検証＋エラーメッセージ。

## 低
- **[低] 移管ウィザードの列リストがA-Z/AA-AZで頭打ち**
  - `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas:470-483`
  - 影響: BA列以降のマッピング不可。大規模シートでは機能不足。
  - 改善案: 動的生成（実際の使用列まで）または AAA まで拡張。

- **[低] SheetModuleに Option Explicit が無い**
  - `vba/SheetModule_UTF8.bas` / `vba/SheetModule_SJIS.bas`
  - 影響: タイプミスが静かに混入しやすく、保守性低下。
  - 改善案: `Option Explicit` を追加。

- **[低] 計算モードが元に戻らないケース**
  - 例: `vba/HierarchyColor_UTF8.bas:34-35,84` / `vba/addons/DataMigration/DataMigrationWizard_UTF8.bas:99-100,220`
  - 影響: 手動計算のブックで強制的に自動計算に戻る。
  - 改善案: `prevCalc` を保存・復元。

- **[低] READMEの記載が実態とズレ**
  - `README.md:7,40-42,66-69`
  - 影響: 「祝日マスタ」独立シートや `vba/統合版` など、現状存在しない前提の説明。
  - 改善案: 実装に合わせて更新。

## パフォーマンス
- **貼り付け時のO(n^2)化**
  - `vba/SheetModule_UTF8.bas:95-130` と `:255-272`
  - 影響: まとめ貼り付けで `GetNextNo` が行数分ループ → 体感で大幅に遅くなる。
  - 改善案: 変更行を一括収集し、最大No.は一度だけ算出して使い回す。

- **シェイプ乱造による描画コスト**
  - `vba/InazumaGantt_v2_UTF8.bas:568-737`
  - 影響: 行数が多いほど極端に遅くなる。Excelの安定性にも影響。
  - 改善案: 進捗バーはセル塗り/条件付き書式に寄せ、シェイプは最小限に。

## テスト不足
- 自動テストがありません。最低限、以下を手動確認してください。
  1) 階層色分けがエラーなく適用される（重複宣言修正後）  
  2) UTF8ビルドで日本語が文字化けしない  
  3) 週末非表示が再生成後も保持される  
  4) 大量貼り付け時の入力速度  
  5) 移管ウィザードで数値以外入力時の挙動

## 確認したい前提
- マクロは「必ず InazumaGantt_v2 シート上から実行する」前提ですか？  
  → 前提なら、冒頭でシート名チェックを入れてガードした方が安全です。

## 変更優先度（提案）
1) **HierarchyColor のコンパイルエラー修正**  
2) **UTF8ビルドのエンコーディング修正**  
3) **計算モードの復元漏れ修正**  
4) ActiveSheet依存の削減  
5) パフォーマンス改善（No採番・シェイプ描画）
