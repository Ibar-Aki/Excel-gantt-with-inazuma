# 📊 InazumaGantt v2 - プロジェクト進捗管理システム

Excelベースの高機能ガントチャート管理ツール

## 🎯 クイックスタート

### 初めて使う方

1. **`Ganto2026_v2仕様.xlsm` を開く**
2. **[SETUP.md](SETUP.md) を読んでセットアップ**
3. **VBAモジュールをインポート（3ファイル）**
4. **シートモジュールを設定**
5. **`SetupInazumaGantt` マクロを実行**

👉 **詳細な手順は [SETUP.md](SETUP.md) を参照してください**

### 既存データを移管する方

1. **`vba_modules/import/DataMigration_SJIS.bas` をインポート**
2. **既存シートで `MigrateToV2Format` マクロを実行**
3. **自動的にv2形式に変換されます**

👉 **詳細は [docs/DataMigration_README.md](docs/DataMigration_README.md) を参照**

---

## 📁 ファイル構成

```
c:\研究所\進捗管理表\
├── 📊 Ganto2026_v2仕様.xlsm        # メインExcelファイル
├── 📖 README.md                    # このファイル
├── 🚀 SETUP.md                     # セットアップガイド（初めての方必読）
│
├── 📂 vba_modules/                 # VBAモジュール
│   ├── import/                     # インポート用（Shift-JIS）
│   │   ├── InazumaGantt_v2_SJIS.bas    # メイン機能
│   │   ├── HierarchyColor_SJIS.bas     # 階層色分け
│   │   ├── DataMigration_SJIS.bas      # データ移管
│   │   └── InazumaGantt_v2_SheetModule.bas # シートモジュール用
│   └── source/                     # 開発用（UTF-8）
│       └── DataMigration.bas
│
├── 📂 docs/                        # ドキュメント
│   ├── InazumaGantt_README.md          # 基本機能説明
│   ├── HierarchyColor_README.md        # 階層色分け機能
│   ├── DataMigration_README.md         # データ移管機能
│   ├── InazumaGantt_v2_SheetModule_README.md # シートモジュール説明
│   ├── CHANGELOG.md                    # 変更履歴
│   ├── CUSTOMIZE.md                    # カスタマイズガイド
│   └── TROUBLESHOOTING.md              # トラブルシューティング
│
└── 📂 旧モデル/                    # 旧バージョン（参照不要）
    └── ...
```

---

## ⭐ 主な機能

### 1. 階層型タスク管理

タスクを入力する列で階層が自動決定：

| 入力列 | 階層レベル | 用途 |
|--------|-----------|------|
| C列 | LV1 | 大項目・フェーズ |
| D列 | LV2 | 中項目 |
| E列 | LV3 | 小項目 |
| F列 | LV4 | 詳細項目 |

**自動機能**: タスク入力時にLV（A列）が自動設定されます

### 2. 進捗管理

- **進捗率（I列）** を入力すると、**状況（H列）** が自動更新
  - 0% → 未着手
  - 1～99% → 進行中  - 100% → 完了

### 3. ダブルクリック完了

タスク行を**ダブルクリック**すると一発で完了にできます：
- 進捗率 → 100%
- 状況 → 完了
- 完了実績 → 今日の日付（開始実績がある場合）

### 4. イナズマガントチャート

- **予定バー** と **実績バー** を表示
- **イナズマ線** で進捗の遅れを視覚化
- **今日線** で現在位置を明示

### 5. 階層別色分け

LVに応じて自動的に色分け：
- LV1（C列） → サーモン色
- LV2（D列） → 薄い青
- LV3（E列） → 薄い緑
- LV4（F列） → 薄い黄色

タスクが入力された列からN列まで色塗りされます。

---

## 🔧 使い方

### 基本的なワークフロー

1. **タスクを入力**（C～F列、階層別に配置）
2. **詳細情報を入力**（G～N列）
3. **`RefreshInazumaGantt` マクロを実行** → ガント描画
4. **`HierarchyColor.ApplyHierarchyColors` マクロを実行** → 色塗り

### よく使うマクロ

| マクロ名 | 機能 | ショートカット |
|---------|------|--------------|
| `SetupInazumaGantt` | 初回セットアップ | Alt+F8 |
| `RefreshInazumaGantt` | ガント更新 | Alt+F8 |
| `HierarchyColor.ApplyHierarchyColors` | 色塗り | Alt+F8 |
| `MigrateToV2Format` | データ移管 | Alt+F8 |

---

## 📚 ドキュメント

### 機能別ガイド

| ドキュメント | 内容 |
|------------|------|
| [SETUP.md](SETUP.md) | **セットアップ手順（必読）** |
| [InazumaGantt_README.md](docs/InazumaGantt_README.md) | 基本機能の詳細 |
| [HierarchyColor_README.md](docs/HierarchyColor_README.md) | 階層色分け機能 |
| [DataMigration_README.md](docs/DataMigration_README.md) | データ移管方法 |

### サポート

| ドキュメント | 内容 |
|------------|------|
| [TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) | よくある問題と解決方法 |
| [CUSTOMIZE.md](docs/CUSTOMIZE.md) | カスタマイズ方法 |
| [CHANGELOG.md](docs/CHANGELOG.md) | 更新履歴 |

---

## 🎨 列構成（v2形式）

| 列 | 項目 | 説明 |
|----|------|------|
| **A** | LV | 階層レベル（自動設定） |
| **B** | No. | 通し番号 |
| **C** | TASK(LV1) | 大項目 |
| **D** | TASK(LV2) | 中項目 |
| **E** | TASK(LV3) | 小項目 |
| **F** | TASK(LV4) | 詳細項目 |
| **G** | タスク詳細 | 補足説明 |
| **H** | 状況 | 未着手/進行中/完了（自動） |
| **I** | 進捗率 | 0～1または0～100% |
| **J** | 担当 | 担当者名 |
| **K** | 開始予定 | 予定開始日 |
| **L** | 完了予定 | 予定完了日 |
| **M** | 開始実績 | 実際の開始日 |
| **N** | 完了実績 | 実際の完了日 |
| **O～** | ガント | ガントチャート領域 |

---

## 💡 ヒント

### 自動機能を活用
- ✅ タスクを入力するだけで階層が自動設定
- ✅ 進捗率を入力すると状況が自動更新
- ✅ ダブルクリックで即完了

### キーボードショートカット
- `Alt + F8` : マクロ一覧
- `Alt + F11` : VBAエディタ
- `Ctrl + Z` : 元に戻す

### データ移管
既存のガントチャートがある場合は、データ移管機能を使うと一瞬で移行できます。

---

## 🔄 更新履歴

最新の更新情報は [docs/CHANGELOG.md](docs/CHANGELOG.md) を参照してください。

---

## ⚙️ 技術情報

- **言語**: VBA (Visual Basic for Applications)
- **動作環境**: Microsoft Excel 2016以降推奨
- **エンコーディング**: 
  - インポート用: Shift-JIS (`vba_modules/import/`)
  - 開発用: UTF-8 (`vba_modules/source/`)

---

## 📞 サポート

問題が発生した場合は [docs/TROUBLESHOOTING.md](docs/TROUBLESHOOTING.md) を参照してください。

---

**Have a nice project management! 🚀**
