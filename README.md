# 📊 InazumaGantt v2

Excelベースのイナズマガントチャート管理ツール

> [!NOTE]
> **Excelファイル名は自由に変更可能です**  
> シート名（`InazumaGantt_v2`、`祝日マスタ`）は変更しないでください。

---

## クイックスタート

### 1. VBAモジュールをインポート

```
Alt + F11 → ファイル → ファイルのインポート
```

`vba/` フォルダから以下をインポート：
- ✅ `InazumaGantt_v2.bas` （必須）
- ✅ `HierarchyColor.bas` （必須）

### 2. シートモジュールを設定

1. VBAエディタで「InazumaGantt_v2」シートをダブルクリック
2. `vba/SheetModule.bas` の内容を貼り付け
3. 保存して閉じる

### 3. セットアップ実行

```
Alt + F8 → SetupInazumaGantt → 実行
```

👉 **詳細は [SETUP.md](SETUP.md) を参照**

---

## 主な機能

| 機能 | 説明 |
|------|------|
| 📊 ガントチャート | 予定・進捗・実績バーを表示 |
| ⚡ イナズマ線 | 進捗の遅れを視覚化 |
| 🎨 階層色分け | LVに応じた自動色分け |
| 🖱️ ダブルクリック完了 | タスクを即座に完了 |
| 🔄 自動機能 | 階層・状況の自動設定 |

---

## ファイル構成

```
📁 vba/               ← VBAモジュール（インポート用）
📁 docs/              ← ドキュメント
📁 dev/               ← 開発者用（通常は不要）
```

---

## ドキュメント

| ファイル | 内容 |
|----------|------|
| [SETUP.md](SETUP.md) | セットアップ手順 |
| [docs/FEATURES.md](docs/FEATURES.md) | 機能詳細 |
| [docs/CUSTOMIZE.md](docs/CUSTOMIZE.md) | カスタマイズ方法 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 問題解決 |
| [CHANGELOG.md](CHANGELOG.md) | 更新履歴 |

---

## 使い方

### 基本フロー

1. **タスクを入力**（C〜F列）
2. **日付を入力**（K〜N列）
3. **`RefreshInazumaGantt`** でガント更新
4. **`ApplyHierarchyColors`** で色分け（任意）

### よく使うマクロ

| マクロ | 機能 |
|--------|------|
| `SetupInazumaGantt` | 初回セットアップ |
| `RefreshInazumaGantt` | ガント更新 |
| `ApplyHierarchyColors` | 色分け適用 |

---

## ライセンス

MIT License - [LICENSE](LICENSE)
