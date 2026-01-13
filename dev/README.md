# 開発者向けリソース

このフォルダには開発者・メンテナー向けのリソースが含まれています。

---

## フォルダ構成

```
dev/
├── archive/          # 旧バージョンのファイル
├── extra_modules/    # 追加VBAモジュール
└── docs/             # 技術ドキュメント
```

---

## extra_modules/

追加のVBAモジュールです。必要に応じてインポートしてください。

| ファイル | 用途 |
|----------|------|
| `ErrorHandler_SJIS.bas` | 統一エラーハンドリング・ログ機能 |
| `InazumaGanttTests_SJIS.bas` | 単体テスト・統合テスト |
| `SheetModule_Legacy_SJIS.bas` | 旧版シートモジュール（互換性用） |

> **注意**: SetupWizardは `vba/` フォルダに移動しました（必須モジュール）。

---

## docs/

技術ドキュメントです。

| ファイル | 内容 |
|----------|------|
| `SETUP.md` | セットアップ詳細手順（v2.2対応） |
| `ARCHITECTURE.md` | システム構成・モジュール関係図 |
| `DEPENDENCIES.md` | 依存関係・データフロー |
| `CODE_QUALITY.md` | コーディング規約・ベストプラクティス |
| `IMPROVEMENT_REPORT.md` | 改善プロジェクトの記録 |
| `INSTALLATION.md` | 詳細インストール手順 |
| `罫線サマリ_表.md` | 罫線設定の詳細仕様 |

---

## archive/

旧バージョンのバックアップです。通常は使用しません。
