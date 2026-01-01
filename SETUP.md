# 🚀 InazumaGantt v2 セットアップガイド

## 📋 目次
1. [必要なファイル](#必要なファイル)
2. [セットアップ手順](#セットアップ手順)
3. [初回起動](#初回起動)
4. [データ移管（既存データがある場合）](#データ移管)

---

## 必要なファイル

### メインファイル
- ✅ `Ganto2026_v2仕様.xlsm` - Excelファイル

### VBAモジュール（インポート用）
すべて `vba_modules/import/` フォルダ内：
- ✅ `InazumaGantt_v2_SJIS.bas` - メイン機能
- ✅ `HierarchyColor_SJIS.bas` - 階層色分け
- ✅ `DataMigration_SJIS.bas` - データ移管（任意）
- ✅ `InazumaGantt_v2_SheetModule.bas` - シートモジュール用

---

## セットアップ手順

### ステップ1: Excelファイルを開く

1. `Ganto2026_v2仕様.xlsm` を開く
2. マクロを有効にする

### ステップ2: VBAモジュールをインポート

1. **VBAエディタを開く**
   ```
   Alt + F11
   ```

2. **標準モジュールとしてインポート**
   
   以下のファイルを順番にインポート：
   
   **File → Import File** （または ファイル → ファイルのインポート）
   
   - ✅ `vba_modules/import/InazumaGantt_v2_SJIS.bas`
   - ✅ `vba_modules/import/HierarchyColor_SJIS.bas`
   - ✅ `vba_modules/import/DataMigration_SJIS.bas` **（任意）**

3. **インポート確認**
   
   VBAエディタの左側「プロジェクトエクスプローラー」に以下が表示されればOK：
   ```
   VBAProject (Ganto2026_v2仕様.xlsm)
   ├─ 標準モジュール
   │  ├─ InazumaGantt_v2
   │  ├─ HierarchyColor
   │  └─ DataMigration （インポートした場合）
   └─ ...
   ```

### ステップ3: シートモジュールを設定

1. **VBAエディタで、プロジェクトエクスプローラーの「InazumaGantt_v2」シートをダブルクリック**

2. **コードウィンドウが開いたら、以下のファイルの内容を全てコピー&貼り付け**
   ```
   vba_modules/import/InazumaGantt_v2_SheetModule.bas
   ```

3. **VBAエディタを閉じる**

### ステップ4: 初回セットアップマクロを実行

1. **Excelに戻る**
2. **マクロを実行**
   ```
   Alt + F8
   ```
3. **「SetupInazumaGantt」を選択 → 実行**
4. **ガントチャートの開始日を入力**（例: 26/01/01）
5. **セットアップ完了！**

---

## 初回起動

### タスクを入力する

1. **C～F列のいずれかにタスク名を入力**
   - C列 → LV1（大項目）
   - D列 → LV2（中項目）
   - E列 → LV3（小項目）
   - F列 → LV4（詳細項目）

2. **G～N列にデータを入力**
   - G列: タスク詳細
   - H列: 状況（自動設定）
   - I列: 進捗率（0～1 または 0～100%）
   - J列: 担当者
   - K列: 開始予定
   - L列: 完了予定
   - M列: 開始実績
   - N列: 完了実績

### ガントチャートを更新

```
Alt + F8 → RefreshInazumaGantt → 実行
```

### 階層色分けを適用（任意）

```
Alt + F8 → ApplyHierarchyColors → 実行
```

---

## データ移管

既存のガントチャートデータがある場合：

1. **既存シートを開く**
2. **移管マクロを実行**
   ```
   Alt + F8 → MigrateToV2Format → 実行
   ```
3. **InazumaGantt_v2シートに自動的にデータが移管されます**

詳細は [`docs/DataMigration_README.md`](docs/DataMigration_README.md) を参照。

---

## 🎯 次のステップ

- 📚 [機能の詳細](docs/InazumaGantt_README.md)
- 🎨 [階層色分け機能](docs/HierarchyColor_README.md)
- 🔧 [カスタマイズ方法](docs/CUSTOMIZE.md)
- ❓ [トラブルシューティング](docs/TROUBLESHOOTING.md)

---

## 💡 ヒント

### 自動機能
- タスク入力時に**LVが自動設定**されます
- 進捗率入力時に**状況が自動更新**されます
  - 0% → 未着手
  - 1～99% → 進行中
  - 100% → 完了

### ダブルクリック完了
タスク行をダブルクリックすると、そのタスクを一発で完了にできます。

### キーボードショートカット
- `Alt + F8`: マクロ一覧を開く
- `Alt + F11`: VBAエディタを開く

---

セットアップが完了したら、[README.md](README.md) に戻って機能を確認してください！
