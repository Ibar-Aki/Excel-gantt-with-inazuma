# InazumaGantt_v2 シートモジュール説明

InazumaGantt_v2 シート用のイベントコードです。

## できること
- ダブルクリックでタスク完了処理を実行
- 入力変更時に階層レベルを自動判定

## 設定手順
1. ExcelでAlt+F11を押してVBAエディタを開く
2. プロジェクトエクスプローラーで InazumaGantt_v2 を開く
3. InazumaGantt_v2_SheetModule.bas を貼り付け
4. VBAエディタを閉じる

## 動作詳細
- Worksheet_BeforeDoubleClick: 完了処理を呼び出し
  InazumaGantt_v2.CompleteTaskByDoubleClick を実行
- Worksheet_Change: C-F 列の変更時に階層判定
  InazumaGantt_v2.AutoDetectTaskLevel を呼び出し
