# 🚀 セットアップガイド

InazumaGantt v2 のセットアップ手順です。

> [!IMPORTANT]
> **Excelファイル名は自由に変更可能です**  
> ただし、**シート名**（`InazumaGantt_v2`、`祝日マスタ`）は変更しないでください。

---

## 必要なファイル

### VBAモジュール（`vba/` フォルダ内）

| ファイル | 用途 | 必須度 |
|----------|------|--------|
| `InazumaGantt_v2.bas` | メイン機能 | ⭐ 必須 |
| `HierarchyColor.bas` | 階層色分け | ⭐ 必須 |
| `SheetModule.bas` | シートイベント | ⭐ 必須 |
| `DataMigration.bas` | データ移管 | 任意 |

---

## セットアップ手順

### ステップ1: VBAモジュールをインポート

1. **Excelファイル（.xlsm）を開く**

2. **VBAエディタを開く**
   ```
   Alt + F11
   ```

3. **モジュールをインポート**
   ```
   ファイル → ファイルのインポート
   ```
   
   以下を順番にインポート：
   - ✅ `vba/InazumaGantt_v2.bas`
   - ✅ `vba/HierarchyColor.bas`

4. **確認**
   
   「標準モジュール」に以下が表示されればOK：
   ```
   標準モジュール
   ├─ InazumaGantt_v2
   └─ HierarchyColor
   ```

### ステップ2: シートモジュールを設定

1. **VBAエディタで「InazumaGantt_v2」シートをダブルクリック**
   
   （まだシートがない場合はステップ3の後に設定）

2. **コードを貼り付け**
   
   `vba/SheetModule.bas` の内容を全てコピー＆貼り付け

3. **保存して閉じる**
   ```
   Ctrl + S → Alt + Q
   ```

### ステップ3: 初回セットアップを実行

1. **Excelに戻る**

2. **マクロを実行**
   ```
   Alt + F8 → SetupInazumaGantt → 実行
   ```

3. **開始日を入力**（例: `26/01/01`）

4. **完了！**

---

## 次のステップ

### タスクを入力

| 列 | 内容 |
|----|------|
| C〜F列 | タスク名（入力位置で階層が決まる） |
| K列 | 開始予定日 |
| L列 | 完了予定日 |
| I列 | 進捗率（0〜1 または 0〜100%） |

### ガントチャートを更新

```
Alt + F8 → RefreshInazumaGantt → 実行
```

### 色分けを適用（任意）

```
Alt + F8 → ApplyHierarchyColors → 実行
```

---

## データ移管（既存データがある場合）

既存のガントチャートからデータを移行する場合：

1. `vba/DataMigration.bas` をインポート
2. 既存シートで `Alt + F8 → MigrateToV2Format → 実行`

---

## トラブルシューティング

問題が発生した場合は [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) を参照。

---

## 開発者向け

追加のモジュールやドキュメントは `dev/` フォルダにあります：

- `dev/extra_modules/` - SetupWizard, ErrorHandler, テスト
- `dev/docs/` - システム構成、コード品質ガイド
