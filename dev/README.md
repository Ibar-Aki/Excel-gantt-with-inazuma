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
| `SetupWizard.bas` | 対話式セットアップウィザード |
| `ErrorHandler.bas` | 統一エラーハンドリング・ログ機能 |
| `InazumaGanttTests.bas` | 単体テスト・統合テスト |
| `SheetModule_Legacy.bas` | 旧版シートモジュール |

---

## docs/

技術ドキュメントです。

| ファイル | 内容 |
|----------|------|
| `ARCHITECTURE.md` | システム構成・モジュール関係図 |
| `DEPENDENCIES.md` | 依存関係・データフロー |
| `CODE_QUALITY.md` | コーディング規約・ベストプラクティス |
| `IMPROVEMENT_REPORT.md` | 改善プロジェクトの記録 |
| `INSTALLATION.md` | 詳細インストール手順 |

---

## archive/

旧バージョンのバックアップです。通常は使用しません。
