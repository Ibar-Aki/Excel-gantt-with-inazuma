# カスタマイズガイド

InazumaGantt v2.2 の設定を変更する方法です。

---

## 設定マスタシート（v2.2新機能）

「設定マスタ」シートでダブルクリック完了の動作をカスタマイズできます。

| 行 | 設定項目 | デフォルト | 効果 |
|----|----------|------------|------|
| B3 | 機能有効 | TRUE | ダブルクリック完了のON/OFF |
| B4 | 完了日自動入力 | TRUE | 完了実績日に今日を自動入力 |
| B5 | 取り消し線 | TRUE | タスクに取り消し線を適用 |
| B6 | 灰色変更 | TRUE | タスクを灰色に変更 |

> 設定を変更した場合、Excelを再起動するか、シートを再アクティブにしてください。

---

## 色の変更

### 階層別の色

`vba/HierarchyColor_SJIS.bas` の定数を変更：

```vba
' 例: LV1の色をオレンジに変更
Public Const COLOR_LV1 As Long = RGB(255, 200, 100)
```

| 定数 | デフォルト | 説明 |
|------|-----------|------|
| `COLOR_LV1` | サーモン | 大項目の色（C〜N列） |
| `COLOR_LV2` | 薄い青 | 中項目の色（D〜N列） |
| `COLOR_LV3` | 薄い緑 | 小項目の色（E〜N列） |
| `COLOR_LV4` | 薄い黄色 | 詳細項目の色（F〜N列） |

### ガントチャートの色

`vba/InazumaGantt_v2_SJIS.bas` の定数を変更：

| 定数 | デフォルト | 説明 |
|------|-----------|------|
| `COLOR_PLAN` | RGB(245,245,245) | 予定バーの色（薄い灰色） |
| `COLOR_PROGRESS` | RGB(48,84,150) | 進捗バーの色（紺色） |
| `COLOR_ACTUAL` | RGB(0,176,80) | 実績バーの色（緑色） |
| `COLOR_TODAY` | RGB(255,0,0) | 今日線の色（赤） |
| `COLOR_INAZUMA` | RGB(255,165,0) | イナズマ線の色（オレンジ） |
| `COLOR_HOLIDAY` | RGB(70,70,80) | 土日・祝日列の色（濃い灰色） |

---

## 表示期間の変更

`vba/InazumaGantt_v2_SJIS.bas` の定数を変更：

```vba
' ガントチャートの表示日数（デフォルト: 120日）
Public Const GANTT_DAYS As Long = 180
```

---

## バーの高さの変更

`DrawGanttBars` サブルーチン内の値を変更：

```vba
barHeight = 6  ' 予定バー・進捗バーの高さ
actualBarHeight = 6  ' 実績バーの高さ
```

---

## 列位置の変更

> ⚠️ **注意**: 列位置を変更する場合は、すべてのモジュールで整合性を取る必要があります。

`vba/InazumaGantt_v2_SJIS.bas` の列定数：

```vba
Public Const COL_HIERARCHY As String = "A"
Public Const COL_NO As String = "B"
Public Const COL_TASK As String = "C"
' ... 略 ...
Public Const COL_GANTT_START As String = "O"
```

---

## データ開始行の変更

```vba
Public Const ROW_DATA_START As Long = 9
```

> SheetModuleの`ROW_DATA_START`も同じ値に変更してください。

---

## 状況の選択肢を変更

`ApplyDataValidationAndFormats` 関数内のドロップダウン設定を変更：

```vba
.Add Type:=xlValidateList, Formula1:="未着手,進行中,完了,保留,中止"
```

---

## 階層色分けの仕様

階層色分けは**条件付き書式**で実装されています。

- `SetupHierarchyColors` を一度実行すれば、以降は自動的に適用
- A列（LV）の値に応じて、対応する範囲に色が付く

| LV | 色分け範囲 |
|----|-----------|
| 1 | C〜N列 |
| 2 | D〜N列 |
| 3 | E〜N列 |
| 4 | F〜N列 |

---

## 祝日マスタの設定

「祝日マスタ」シートのA列に祝日を登録すると：

- ガントチャート上で祝日列が濃い灰色で表示
- 日付シフト機能（`ShiftDates`）で祝日が自動的にスキップされる

---

## PDF出力のカスタマイズ

`ExportToPDF` 関数を編集して出力範囲を変更できます：

- 現在の設定: A列〜N列 + 当月末までのガント領域
- ファイル名: `[シート名]_[YYYYMMDD].pdf`
