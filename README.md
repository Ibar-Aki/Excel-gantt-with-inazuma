# 📊 InazumaGantt v3

Excelベースのイナズマガントチャート管理ツール

> [!NOTE]
> **Excelファイル名は自由に変更可能です**  
> シート名（`InazumaGantt_v3`、`祝日マスタ`、`設定マスタ`）は変更しないでください。

---

## クイックスタート

### 1. VBAモジュールをインポート

```
Alt + F11 → ファイル → ファイルのインポート
```

`- **vba/** : 最新のソースコード (UTF-8)

- `InazumaGantt_v3_UTF8.bas` : メインロジック
- `SheetModule_UTF8.bas` : シートイベント制御
- `SetupWizard_UTF8.bas` : セットアップ機能
- `HierarchyColor_UTF8.bas` : 階層色分け機能
- `addons/` : 拡張機能（データ移管など）（任意：旧形式からの移行用）

### 2. シートモジュールを設定

1. VBAエディタで「Sheet1」（または対象シート）をダブルクリック
2. `vba/SheetModule_UTF8.bas` の内容を貼り付け
3. 保存して閉じる

### 3. セットアップウィザードを実行

```
Alt + F8 → RunSetupWizard → 実行
```

> **ウィザードが自動的に以下を設定します:**
>
> - メインシート（InazumaGantt_v3）作成
> - 設定マスタシート作成（祝日欄含む）
> - 設定マスタシート作成
> - 階層色分け（条件付き書式）
> - ガントチャート描画

👉 **詳細は [dev/docs/SETUP.md](dev/docs/SETUP.md) を参照**

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

```text
📁 vba/               ← VBAモジュール（_SJIS.bas をインポート）
📁 docs/              ← 利用者向けドキュメント
📁 dev/               ← 開発者用ドキュメント・仕様
📁 output/            ← ビルド済みExcelファイル
```

---

## ドキュメント

### 利用者向け

| ファイル | 内容 |
|----------|------|
| [docs/利用者ガイド.md](docs/利用者ガイド.md) | **操作マニュアル** |
| [docs/FEATURES.md](docs/FEATURES.md) | 機能詳細 |
| [docs/CUSTOMIZE.md](docs/CUSTOMIZE.md) | カスタマイズ方法 |
| [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | 問題解決 |
| [CHANGELOG.md](CHANGELOG.md) | 更新履歴 |

### 開発者向け

| ファイル | 内容 |
|----------|------|
| [dev/docs/SETUP.md](dev/docs/SETUP.md) | セットアップ詳細手順 |
| [dev/docs/ARCHITECTURE.md](dev/docs/ARCHITECTURE.md) | アーキテクチャ |
| [vba/README.md](vba/README.md) | VBAモジュール説明 |

---

## 使い方

### 基本フロー

1. **セットアップウィザード実行** → `RunSetupWizard`
2. **タスクを入力**（C〜F列）→ No.・進捗率・状況が自動入力
3. **日付を入力**（K〜N列）
4. **ガント更新ボタン** または `RefreshInazumaGantt` でガント更新

### よく使うマクロ

| マクロ | 機能 |
|--------|------|
| `RunSetupWizard` | **初回セットアップ（推奨）** |
| `RefreshInazumaGantt` | ガント更新（バー・イナズマ線描画） |
| `ResetFormatting` | 書式リセット（罫線・色の修復） |
| `ExportToPDF` | PDF出力 |

---

## ライセンス

MIT License - [LICENSE](LICENSE)
