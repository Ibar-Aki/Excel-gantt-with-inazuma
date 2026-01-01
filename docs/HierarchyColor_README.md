# HierarchyColor モジュール説明

InazumaGantt_v2 と連携して、
階層レベルに応じた行の色分けを行います。

## できること
- 階層レベル（1-4）に応じた薄色の塗り分け
- タスク入力列から N 列までの色塗り
- 色塗りの一括クリア

## 使い方
1. HierarchyColor_SJIS.bas を標準モジュールとして取り込み
2. ApplyHierarchyColors を実行して色分け
3. 解除したい場合は ClearHierarchyColors を実行

## 自動更新（任意）
シートモジュールに Worksheet_Change を貼ると
階層列の変更時に自動で色分けできます。
（コードはモジュール末尾のコメント参照）

## 依存関係
- InazumaGantt_v2 の GetTaskColumnByLevel を呼び出します
- ROW_DATA_START は 9 行目（InazumaGantt_v2 に合わせる）

## 設定値（変更する場合）
- COL_HIERARCHY: 階層列（既定 A）
- COL_COLOR_START: 色塗り開始列（既定 B）
- COL_COLOR_END: 色塗り終了列（既定 N）
- ROW_DATA_START: データ開始行（既定 9）

## 色定義
- LV1: 薄いオレンジ/サーモン
- LV2: 薄い青
- LV3: 薄い緑
- LV4: 薄い黄色
- LV5以上: 薄い紫（現行ロジックでは 1-4 のみ適用）

