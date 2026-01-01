# VBA Modules

このフォルダにはInazumaGantt v2で使用するVBAモジュールが格納されています。

## 📂 フォルダ構成

### import/

**用途**: Excelにインポートするためのファイル  
**エンコーディング**: Shift-JIS (Windows-31J)

このフォルダ内のファイルをVBAエディタでインポートしてください。

#### 含まれるファイル

| ファイル名 | 説明 | インポート先 |
|-----------|------|------------|
| `InazumaGantt_v2_SJIS.bas` | メイン機能モジュール | 標準モジュール |
| `HierarchyColor_SJIS.bas` | 階層別色分け機能 | 標準モジュール |
| `DataMigration_SJIS.bas` | データ移管機能（任意） | 標準モジュール |
| `InazumaGantt_v2_SheetModule.bas` | シートイベント処理 | シートモジュール |

## 📥 インポート方法

### 標準モジュールのインポート

1. Excel で `Alt + F11` を押してVBAエディタを開く
2. メニューから **ファイル → ファイルのインポート**
3. 以下のファイルを順番にインポート：
   - `InazumaGantt_v2_SJIS.bas`
   - `HierarchyColor_SJIS.bas`
   - `DataMigration_SJIS.bas`（必要な場合）

### シートモジュールの設定

1. VBAエディタで「InazumaGantt_v2」シートをダブルクリック
2. `InazumaGantt_v2_SheetModule.bas` の内容を全てコピー
3. シートモジュールに貼り付け

## 🔧 エンコーディングについて

### なぜShift-JISなのか？

Excel VBAは内部でShift-JIS（Windows-31J）エンコーディングを使用しています。
UTF-8ファイルをインポートすると、日本語が文字化けする可能性があります。

### 開発時の注意

VBAコードを編集する場合：
1. エクスポート機能でコードを取り出す
2. UTF-8で編集（バージョン管理用）
3. Shift-JISに変換してインポート

変換コマンド例（PowerShell）:
```powershell
Get-Content "source.bas" -Encoding UTF8 | Out-File "output_SJIS.bas" -Encoding Default
```

## 📝 モジュール間の依存関係

```
InazumaGantt_v2_SJIS.bas (メイン)
├── 独立して動作
└── HierarchyColor の関数を呼び出し可能

HierarchyColor_SJIS.bas
├── InazumaGantt_v2 の定数を参照
└── GetTaskColumnByLevel() を呼び出し

DataMigration_SJIS.bas
└── 独立して動作

InazumaGantt_v2_SheetModule.bas (シート)
├── InazumaGantt_v2 の関数を呼び出し
└── AutoDetectTaskLevel()
└── CompleteTaskByDoubleClick()
```

## ⚠️ トラブルシューティング

### 文字化けが発生する場合

- ファイルが正しくShift-JISでエンコードされているか確認
- `*_SJIS.bas` ファイルを使用しているか確認

### インポートできない場合

- ファイル拡張子が `.bas` であることを確認
- ファイルが破損していないか確認
- Excel のマクロ設定を確認

---

詳細な使用方法は [../SETUP.md](../SETUP.md) を参照してください。
