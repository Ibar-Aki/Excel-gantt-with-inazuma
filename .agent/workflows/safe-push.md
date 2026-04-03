---
description: 安全ブロックを解除してGitHubにプッシュする
---

# 安全プッシュ ワークフロー

安全メカニズム（Global Safety Block）を解除してGitHubにプッシュします。

## 手順

1. 現在のGit状態を確認

```powershell
git status
```

1. プッシュ待ちのコミットを確認

```powershell
git log --oneline origin/main..HEAD
```

1. 環境変数を設定してプッシュを実行
// turbo

```powershell
$env:ALLOW_PUSH=1; git push origin main
```

## 注意事項

- プッシュ前に必ず変更内容を確認してください
- `ALLOW_PUSH=1` は一時的な環境変数で、現在のセッションのみ有効です
