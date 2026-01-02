# カスタマイズガイド

InazumaGantt v2 の設定を変更する方法です。

---

## 色の変更

### 階層別の色

`vba/HierarchyColor.bas` の定数を変更：

```vba
' 例: LV1の色をオレンジに変更
Public Const COLOR_LV1 As Long = RGB(255, 200, 100)
```

| 定数 | デフォルト | 説明 |
|------|-----------|------|
| `COLOR_LV1` | サーモン | 大項目の色 |
| `COLOR_LV2` | 薄い青 | 中項目の色 |
| `COLOR_LV3` | 薄い緑 | 小項目の色 |
| `COLOR_LV4` | 薄い黄色 | 詳細項目の色 |

### ガントチャートの色

`vba/InazumaGantt_v2.bas` の定数を変更：

| 定数 | 説明 |
|------|------|
| `COLOR_PLAN` | 予定バーの色 |
| `COLOR_PROGRESS` | 進捗バーの色 |
| `COLOR_ACTUAL` | 実績バーの色 |
| `COLOR_TODAY` | 今日線の色 |
| `COLOR_INAZUMA` | イナズマ線の色 |

---

## 表示期間の変更

`vba/InazumaGantt_v2.bas` の定数を変更：

```vba
' ガントチャートの表示日数（デフォルト: 120日）
Public Const GANTT_DAYS As Long = 180
```

---

## 列位置の変更

> ⚠️ **注意**: 列位置を変更する場合は、すべてのモジュールで整合性を取る必要があります。

`vba/InazumaGantt_v2.bas` の列定数：

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

---

## 状況の選択肢を変更

`ApplyDataValidationAndFormats` 関数内のドロップダウン設定を変更：

```vba
.Add Type:=xlValidateList, Formula1:="未着手,進行中,完了,保留,中止"
```
