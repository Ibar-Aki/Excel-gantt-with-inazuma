# 📊 InazumaGantt v2.2

Excelベースのイナズマガントチャート管理ツール

> [!NOTE]
> **Excelファイル名は自由に変更可能です**  
> シート名（`InazumaGantt_v2`、`祝日マスタ`、`設定マスタ`）は変更しないでください。

---

## クイックスタート

### 1. VBAモジュールをインポート

```
Alt + F11 → ファイル → ファイルのインポート
```

`vba/` フォルダから以下をインポート：

- ✅ `InazumaGantt_v2_SJIS.bas` （必須）
- ✅ `HierarchyColor_SJIS.bas` （必須）
- ✅ `SetupWizard_SJIS.bas` （推奨）
- 🔹 `DataMigration_SJIS.bas` （任意）

### 2. シートモジュールを設定

1. VBAエディタで「InazumaGantt_v2」シートをダブルクリック
2. `vba/SheetModule_SJIS.bas` の内容を貼り付け
3. 保存して閉じる

### 3. セットアップ実行

```
Alt + F8 → RunSetupWizard → 実行
```

👉 **詳細は [SETUP.md](SETUP.md) を参照**

---

## 主な機能

| 機能 | 説明 |
|------|------|
| 📊 ガントチャート | 予定バー（薄灰+黒枠）、進捗バー（紺色）、実績バー（緑色） |
| ⚡ イナズマ線 | 今日基準型で進捗の遅れを視覚化（オレンジ） |
| 🎨 階層色分け | 条件付き書式でLVに応じた自動色分け |
| 🖱️ ダブルクリック完了 | タスクを即座に完了（取り消し線・灰色対応） |
| 📁 折りたたみ | Shift+右クリックでLV1配下を非表示 |
| 📄 PDF出力 | 当月末までのガントを含むPDF出力 |
| 🔄 自動機能 | No.・階層・状況・進捗率の自動設定 |

---

## ファイル構成

```
📁 vba/               ← VBAモジュール（_SJIS.bas をインポート）
📁 docs/              ← 利用者向けドキュメント
📁 dev/               ← 開発者用ドキュメント・仕様
```

---

## ドキュメント

### 利用者向け

| ファイル | 内容 |
|----------|------|
| [SETUP.md](SETUP.md) | セットアップ手順 |
| [dev/docs/利用者ガイド.md](dev/docs/利用者ガイド.md) | **操作マニュアル** |
| [docs/FEATURES.md](docs/FEATURES.md) | 機能詳細 |
| [docs/CUSTOMIZE.md](docs/CUSTOMIZE.md) | カスタマイズ方法 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 問題解決 |
| [CHANGELOG.md](CHANGELOG.md) | 更新履歴 |

### 開発者向け

| ファイル | 内容 |
|----------|------|
| [dev/docs/ARCHITECTURE.md](dev/docs/ARCHITECTURE.md) | アーキテクチャ |
| [dev/docs/ganttマクロ改善メモ.md](dev/docs/ganttマクロ改善メモ.md) | 改善仕様 |
| [vba/README.md](vba/README.md) | VBAモジュール説明 |

---

## 使い方

### 基本フロー

1. **タスクを入力**（C〜F列）→ No.・進捗率・状況が自動入力
2. **日付を入力**（K〜N列）
3. **`RefreshInazumaGantt`** でガント更新

### よく使うマクロ

| マクロ | 機能 |
|--------|------|
| `SetupInazumaGantt` | 初回セットアップ |
| `RefreshInazumaGantt` | ガント更新（バー・イナズマ線描画） |
| `SetupHierarchyColors` | 階層色分け（条件付き書式設定） |

---

## ライセンス

MIT License - [LICENSE](LICENSE)
