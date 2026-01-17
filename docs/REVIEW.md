# Excel-gantt-with-inazuma コードレビュー

対象: `vba/*_UTF8.bas`, `vba/addons/DataMigration/*_UTF8.bas`, `BuildInazumaGantt.ps1`, `FixEncoding.ps1`
レビュー日: 2026-01-17

## 重大/高
- **設定マスタの破壊リスク**: データ移管の保存/読込が `設定マスタ` を使うため、ダブルクリック設定や祝日マスタを上書き/混在させます。結果として祝日色付け・完了処理設定・移管設定が競合し、データ損失や誤動作が起きます。  
  参照: `vba/addons/DataMigration/DataMigrationWizard_UTF8.bas:238,244,299,306,353,358` / `vba/InazumaGantt_v2_UTF8.bas:1241,1244,1320,1323` / `vba/SheetModule_UTF8.bas:189,193`
  - **改善案**: 移管設定専用シート（例: `移管設定`）を新設し、`設定マスタ` から完全分離。既存の `設定マスタ` を検知した場合は上書きしないガードを追加。

## 中
- **UTF8/SJIS 差分によるビルド破綻**: `HOLIDAY_SHEET_NAME` が UTF8 版に定義されておらず、`ShiftDates` で参照されるため `FixEncoding.ps1` でSJIS生成するとコンパイルエラーになります。  
  参照: `vba/InazumaGantt_v2_UTF8.bas:1437,1455`（定義はSJIS側のみ）
  - **改善案**: UTF8側に定数を追加し、SJISと同一内容に同期。ソースはUTF8のみ正とし、SJISは生成物扱いにする運用も検討。

- **祝日マスタ参照の不整合**: 祝日情報は `設定マスタ` に置かれているのに `ShiftDates` は `HOLIDAY_SHEET_NAME`（別シート）を参照します。結果として営業日シフトで祝日が無視されます。  
  参照: `vba/InazumaGantt_v2_UTF8.bas:1241,1289,1442,1455`
  - **改善案**: `ShiftDates` の祝日取得を `設定マスタ`（A13以降）へ統一、または `祝日マスタ` シートを復活させて一貫性を持たせる。

- **最終行検出の欠落**: `GetLastDataRow` が C/G/K/L/M/N 列のみを参照しており、D/E/F（下位タスク）だけ入力された行を見逃します。結果として書式・検証・ガント描画の対象外になる可能性があります。  
  参照: `vba/InazumaGantt_v2_UTF8.bas:266-274`
  - **改善案**: D/E/F 列も最終行判定に追加。

- **土日非表示状態のリセット**: `RegenerateDateHeaders` が全列幅を `3` に上書きするため、`ToggleWeekends` で隠した土日が `Refresh/Reset` 後に復活します。  
  参照: `vba/InazumaGantt_v2_UTF8.bas:1156,1192` / `vba/InazumaGantt_v2_UTF8.bas:1040`
  - **改善案**: 週末列の幅は保持し、再生成時は既存の幅を尊重する。

## 低
- **階層色分けの適用範囲が固定**: `DATA_ROWS_DEFAULT` までしか条件付き書式を適用していないため、行を増やすと色分けが切れます。  
  参照: `vba/HierarchyColor_UTF8.bas:38,108`
  - **改善案**: 実データの最終行に合わせて再設定、または列全体へ適用。

- **移管ウィザードの列選択がA–Z限定**: 大きいシート（AA列以降）を扱えません。  
  参照: `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas:404`
  - **改善案**: A–ZZ まで生成、または実際の使用列から動的生成。

- **移管マッピングの機能未露出**: `StartPlan/StartActual/EndActual` の設定項目があるのに、フォーム側の入力UIがありません。  
  参照: `vba/addons/DataMigration/DataMigrationWizard_UTF8.bas:165-199` / `vba/addons/DataMigration/MigrationFormBuilder_UTF8.bas`（該当コントロール未生成）

## パフォーマンス/最適化提案
- **イベント処理の重複計算**: `Worksheet_Change` 内でセル単位に `AutoDetectTaskLevel` と `GetNextNo` を繰り返すため、貼り付け時に遅くなります。  
  参照: `vba/SheetModule_UTF8.bas:95-147`  
  - **改善案**: 変更行を集合化し、行単位で一度だけ判定。`GetNextNo` は最大値を一回だけ計算して使い回す。

- **ガント描画のシェイプ乱造**: `DrawGanttBars` が行数分のシェイプを毎回生成・削除するため、行数が多いと重くなります。  
  参照: `vba/InazumaGantt_v2_UTF8.bas:546-794`  
  - **改善案**: 進捗バーは条件付き書式 or セル塗りに寄せ、シェイプは「今日線/イナズマ線」など最小限に限定。

## テスト不足
- 自動テストが存在しません。変更後は最低でも以下の手動確認が必要です。  
  1) 週末非表示→更新で維持されるか  
  2) LV2/LV3のみ入力で書式・検証が適用されるか  
  3) 祝日マスタ入力後の色付けと `ShiftDates` の反映  
  4) データ移管ウィザード使用時に `設定マスタ` が壊れないか

---

### 参考: 主要ファイル
- `vba/InazumaGantt_v2_UTF8.bas`
- `vba/SheetModule_UTF8.bas`
- `vba/HierarchyColor_UTF8.bas`
- `vba/addons/DataMigration/*.bas`
- `BuildInazumaGantt.ps1`
- `FixEncoding.ps1`
