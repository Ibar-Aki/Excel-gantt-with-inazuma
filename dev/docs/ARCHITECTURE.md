# 🏗️ InazumaGantt v2 システム構成図

このドキュメントでは、InazumaGantt v2がどのように動いているかを、**IT知識が少ない方にもわかりやすく**説明します。

---

## 📚 目次

1. [全体の仕組み](#全体の仕組み)
2. [各モジュールの役割](#各モジュールの役割)
3. [データの流れ](#データの流れ)
4. [モジュール間の関係図](#モジュール間の関係図)
5. [よくある質問](#よくある質問)

---

## 全体の仕組み

InazumaGantt v2は、**6つの部品（モジュール）**が協力して動いています。

### イメージ例

レストランで例えるなら：

| モジュール | レストランの例 |
|-----------|--------------|
| **InazumaGantt_v2** | 料理長（全体を指揮） |
| **HierarchyColor** | デコレーション係（色付け） |
| **DataMigration** | 引っ越し屋さん（データ移動） |
| **ErrorHandler** | 衛生管理（エラー対策） |
| **InazumaGanttTests** | 品質検査（テスト） |
| **SetupWizard** | 案内係（セットアップ） |
| **SheetModule** | ホールスタッフ（お客様対応） |

---

## 各モジュールの役割

### 1️⃣ InazumaGantt_v2（メイン機能）⭐ 一番重要

**何をする？**
- ガントチャートを描く
- タスクの階層を判定する
- 日付を計算する
- データを整理する

**例え話**:
お店の**料理長**です。全体を見渡して、他のスタッフに指示を出します。

**主な機能**:
```
✓ SetupInazumaGantt → お店の開店準備
✓ RefreshInazumaGantt → 料理を作る
✓ AutoDetectTaskLevel → 材料を分類する
✓ DrawGanttBars → 盛り付け
```

**他のモジュールとの関係**:
- HierarchyColorに「色を塗って」と依頼
- ErrorHandlerに「エラーが出たら記録して」と依頼
- SheetModuleから「お客さんが来たよ」と連絡を受ける

---

### 2️⃣ HierarchyColor（階層色分け）

**何をする？**
- タスクのレベル（LV1, LV2...）に応じて色を塗る
- 見やすくする

**例え話**:
**デコレーション係**です。料理長（InazumaGantt_v2）の指示で、料理を綺麗に彩ります。

**主な機能**:
```
✓ ApplyHierarchyColors → 色を塗る
✓ ClearHierarchyColors → 色を消す
```

**色の意味**:
- LV1（大項目）→ サーモン色（🍣）
- LV2（中項目）→ 青色（💧）
- LV3（小項目）→ 緑色（🌿）
- LV4（詳細）→ 黄色（⭐）

---

### 3️⃣ DataMigration（データ移管）

**何をする？**
- 古い形式のガントチャートから、新しい形式（v2）にデータを移す

**例え話**:
**引っ越し屋さん**です。古い家（旧バージョン）から新しい家（v2）に荷物（データ）を運びます。

**主な機能**:
```
✓ MigrateToV2Format → 引っ越しを実行
```

**いつ使う？**:
- 既存のガントチャートがある場合のみ
- 初めて使う人は不要

---

### 4️⃣ ErrorHandler（エラー処理）

**何をする？**
- 問題が起きたときに記録する
- ユーザーにわかりやすく伝える
- ログファイルに保存する

**例え話**:
**衛生管理**です。問題が起きても慌てず、記録して原因を調べます。

**主な機能**:
```
✓ HandleError → 問題を記録
✓ WriteLog → 日報に書く
✓ ValidateNumeric → 数字かチェック
```

**ログファイル**:
`InazumaGantt_ErrorLog.txt`（Excelファイルと同じフォルダ）

---

### 5️⃣ InazumaGanttTests（テスト）

**何をする？**
- システムが正しく動くか確認する

**例え話**:
**品質検査**です。料理が完成したら、味見をして品質を確認します。

**主な機能**:
```
✓ RunAllTests → 全部テスト
✓ IntegrationTest_FullWorkflow → 総合テスト
```

**いつ使う？**:
- 開発者が使用
- 一般ユーザーは不要

---

### 6️⃣ SetupWizard（セットアップ）

**何をする？**
- 初めて使う人を案内する
- サンプルデータを作る

**例え話**:
**案内係**です。初めてのお客さんを席に案内して、メニューを説明します。

**主な機能**:
```
✓ RunSetupWizard → セットアップ案内
✓ QuickStart → 常連さん用メニュー
✓ AddSampleData → お試しセット
```

---

### 7️⃣ SheetModule（シートモジュール）

**何をする？**
- ユーザーの操作を検知する
- ダブルクリック、入力などに反応

**例え話**:
**ホールスタッフ**です。お客さん（ユーザー）の要望を聞いて、料理長に伝えます。

**主な機能**:
```
✓ Worksheet_BeforeDoubleClick → お客さんがダブルクリック
✓ Worksheet_Change → お客さんが入力
✓ UpdateStatusByProgress → メニュー内容を更新
```

---

## データの流れ

### 🎬 シナリオ1: タスクを入力する

```
1. ユーザーがC列に「フェーズ1」と入力
   ↓
2. SheetModule が気づく
   「お客さんが何か入力したよ！」
   ↓
3. InazumaGantt_v2 に連絡
   「階層を判定して」
   ↓
4. InazumaGantt_v2.AutoDetectTaskLevel が実行
   「これはLV1だ！」
   ↓
5. A列に「1」が自動入力される
```

### 🎬 シナリオ2: ガントチャートを更新する

```
1. ユーザーが「RefreshInazumaGantt」を実行
   ↓
2. InazumaGantt_v2 が動き出す
   「よし、ガントを描くぞ！」
   ↓
3. データを読み込む
   「タスク名、日付、進捗率...OK」
   ↓
4. ガントバーを描く
   「予定バー（黄色）、実績バー（オレンジ）」
   ↓
5. イナズマ線を描く
   「進捗の遅れを表示」
   ↓
6. 完了！
```

### 🎬 シナリオ3: 色分けする

```
1. ユーザーが「ApplyHierarchyColors」を実行
   ↓
2. HierarchyColor が動き出す
   「色を塗るよ！」
   ↓
3. InazumaGantt_v2 に質問
   「LV1のタスクはどの列？」
   ↓
4. InazumaGantt_v2 が答える
   「C列だよ」
   ↓
5. HierarchyColor が色塗り
   「C～N列をサーモン色に」
   ↓
6. 完了！
```

---

## モジュール間の関係図

### 📊 依存関係（誰が誰を呼ぶか）

```
ユーザー
  │
  ├→ SheetModule（シート操作を検知）
  │    ├→ InazumaGantt_v2.AutoDetectTaskLevel（階層判定）
  │    └→ InazumaGantt_v2.CompleteTaskByDoubleClick（完了処理）
  │
  ├→ SetupWizard.RunSetupWizard（セットアップ）
  │    ├→ InazumaGantt_v2.SetupInazumaGantt（シート作成）
  │    └→ HierarchyColor（色分け確認）
  │
  └→ InazumaGantt_v2.RefreshInazumaGantt（ガント更新）
       ├→ ErrorHandler.HandleError（エラー処理）
       └→ HierarchyColor.ApplyHierarchyColors（色塗り）
```

### 🔗 簡単な図

```
        ユーザー
          ↓
    ┌─────────────┐
    │SheetModule  │ ←──── Excel操作を監視
    └─────────────┘
          ↓
    ┌─────────────┐
    │InazumaGantt │ ←──── メイン処理
    │     v2      │
    └─────────────┘
       ↙    ↓    ↘
  ┌────┐ ┌────┐ ┌────┐
  │色分け│ │移管│ │エラー│
  └────┘ └────┘ └────┘
```

### 🎯 重要度の順位

| 順位 | モジュール | 必須度 |
|------|-----------|--------|
| 1 | InazumaGantt_v2 | ⭐⭐⭐⭐⭐ 絶対必要 |
| 2 | SheetModule | ⭐⭐⭐⭐ かなり重要 |
| 3 | SetupWizard | ⭐⭐⭐⭐ 初心者には必須 |
| 4 | HierarchyColor | ⭐⭐⭐ あると便利 |
| 5 | ErrorHandler | ⭐⭐⭐ あると安心 |
| 6 | DataMigration | ⭐⭐ 移管時のみ |
| 7 | InazumaGanttTests | ⭐ 開発者用 |

---

## よくある質問

### Q1. どのモジュールから勉強すればいい？

**A**: まず **InazumaGantt_v2** を理解してください。これが中心です。

推奨順序:
1. InazumaGantt_v2（メイン機能）
2. SheetModule（ユーザー操作）
3. HierarchyColor（色分け）
4. その他（必要に応じて）

### Q2. モジュールは削除できる？

**A**: 
- ❌ **削除できない**: InazumaGantt_v2, SheetModule
- ✅ **削除可能**: DataMigration, InazumaGanttTests
- ⚠️ **推奨しない**: HierarchyColor, ErrorHandler, SetupWizard

### Q3. カスタマイズしたい場合は？

**A**: 
1. まず全体の動きを理解する（このドキュメント）
2. 変更したい機能を特定する
3. 該当モジュールのコードを読む
4. 小さな変更から始める
5. テストする

詳細は [CUSTOMIZE.md](CUSTOMIZE.md) を参照

### Q4. エラーが出たらどうする？

**A**:
1. `InazumaGantt_ErrorLog.txt` を確認
2. エラーメッセージをメモ
3. [TROUBLESHOOTING.md](TROUBLESHOOTING.md) で検索
4. それでもダメなら、テストを実行  
   `Alt + F8 → RunAllTests`

---

## 🎓 まとめ

### 覚えておくべき3つのポイント

1. **InazumaGantt_v2が中心**  
   すべてはここから始まります

2. **SheetModuleがユーザーと橋渡し**  
   あなたの操作を検知して、InazumaGantt_v2に伝えます

3. **HierarchyColorが装飾**  
   機能的には必須ではないが、見やすさのために重要

### 次のステップ

- [CODE_QUALITY.md](CODE_QUALITY.md) - コード品質の詳細
- [CUSTOMIZE.md](CUSTOMIZE.md) - カスタマイズ方法
- [TROUBLESHOOTING.md](TROUBLESHOOTING.md) - 問題解決

---

**わかりやすかったでしょうか？ 📖**
