# GitHub セットアップコマンド

## GitHubリポジトリ作成後に実行するコマンド

### 1. リモートリポジトリの追加
```bash
# あなたのGitHubユーザー名に置き換えてください
git remote add origin https://github.com/YOUR_USERNAME/csv-data-transformer-gas.git
```

### 2. デフォルトブランチ名を設定（必要に応じて）
```bash
git branch -M main
```

### 3. 初回プッシュ
```bash
git push -u origin main
```

## 今後の開発フロー

### 日常的な作業
```bash
# 変更を確認
git status

# ファイルを追加
git add .

# コミット
git commit -m "feat: 新機能追加"

# プッシュ
git push
```

### ブランチを使った開発
```bash
# 機能開発用ブランチ作成
git checkout -b feature/csv-import

# 開発・コミット
git add .
git commit -m "feat: CSVインポート機能実装"

# プッシュ
git push -u origin feature/csv-import

# メインブランチに戻る
git checkout main

# ブランチをマージ
git merge feature/csv-import
```

### タグによるバージョン管理
```bash
# バージョンタグの作成
git tag v1.0.0
git push origin v1.0.0

# すべてのタグをプッシュ
git push --tags
```

## clasp との連携

### 開発フロー例
```bash
# 1. ローカルで開発
code コード.js

# 2. Gitでコミット
git add .
git commit -m "feat: メイン機能実装"

# 3. GASにプッシュ
clasp push

# 4. GASエディタで確認
clasp open

# 5. テスト・デバッグ

# 6. GitHubにプッシュ
git push
```

## 注意事項

- `.clasp.json` は `.gitignore` に含まれているため、GitHubにはプッシュされません
- GASプロジェクトのスクリプトIDは個人情報なので、公開リポジトリでは注意してください
- 機密情報は絶対にコミットしないでください