# CSV処理・データ変換ツール

Google Apps Script (GAS) を使用した3シート連携データ処理ツールです。

## 概要

スプレッドシート上で動作するCSVファイルの：
- **インポート**: CSVファイルをLPcsvシートに読み込み
- **データ転記・加工**: PPformat・MFformatシートへの自動転記
- **設定管理**: 加工設定メニューによるカスタマイズ
- **ログ記録**: 全処理履歴の自動記録

## 主な機能

### 🔄 CSVインポート（LPcsvシート）
- ファイル選択ダイアログからCSVファイルを選択
- UTF-8、Shift_JIS文字コードに対応
- 最大10MB、10,000行まで対応
- ヘッダー行の自動検出・保持

### 🔄 データ転記・加工
- **PPformatシート**: LPcsvデータの加工版（パターン1）
- **MFformatシート**: LPcsvデータの加工版（パターン2）
- 列マッピング、データ変換、フィルタリング対応

### ⚙️ 加工設定メニュー
- 列マッピング設定（どの列をどこに転記するか）
- データ変換ルール（フォーマット変更、計算式適用）
- フィルタリング条件（特定条件の行のみ転記）

### 🎯 HTMLサイドバーUI（ミニマルデザイン）
- 320px固定幅のシンプルデザイン
- CSVインポート・設定・実行・ログ表示
- リアルタイム進捗表示

### 📋 自動ログ記録
- 全処理履歴を「ログ」シートに自動記録
- タイムスタンプ、処理種別、結果、エラー詳細
- 1000行を超えたら古いログを自動削除

## 技術仕様

- **プラットフォーム**: Google Apps Script
- **対応ブラウザ**: Chrome, Firefox, Safari, Edge
- **必要権限**: Google スプレッドシート編集権限
- **制限事項**: 
  - 最大ファイルサイズ: 10MB
  - 最大行数: 10,000行
  - 実行時間制限: 6分

## シート構成

- **LPcsv**: CSVインポート先シート
- **PPformat**: 加工データシート1
- **MFformat**: 加工データシート2
- **ログ**: 処理履歴記録シート（自動作成）

## プロジェクト構造

```
├── コード.js                    # メイン処理・エントリーポイント
├── csvProcessor.gs              # CSVインポート・データ処理
├── dataTransformer.gs           # データ転記・加工機能
├── settingsManager.gs           # 加工設定の管理
├── logger.gs                    # ログ記録・管理
├── utilities.gs                 # 共通ユーティリティ
├── sidebar.html                 # HTMLサイドバーUI
├── appsscript.json             # GAS設定ファイル
├── requirements.md             # 要件定義書
├── technical-design.md         # 技術設計書
└── project-structure.md        # プロジェクト構造設計
```

## セットアップ

### 1. Google Apps Script プロジェクト

スクリプトID: `1G6dnI7Hj_5wOya2W0qlZ-n48tlO22KU4r1UzT30pWOPd4V30CCP6Ujdp`

### 2. ローカル開発環境セットアップ

```bash
# clasp でログイン（初回のみ）
clasp login

# 既存プロジェクトをクローン
clasp clone 1G6dnI7Hj_5wOya2W0qlZ-n48tlO22KU4r1UzT30pWOPd4V30CCP6Ujdp
```

### 3. ファイルのアップロード

```bash
# ローカルファイルをGASにアップロード
clasp push

# GASエディタを開く
clasp open
```

## 使用方法

### 1. スプレッドシートでの準備

1. 必要なシートを作成：LPcsv, PPformat, MFformat
2. 「拡張機能」→「Apps Script」を選択
3. `showSidebar()` 関数を実行

### 2. 基本操作フロー

1. **CSVインポート**: ファイル選択→LPcsvシートに読み込み
2. **加工設定確認**: 列マッピング・変換ルールの確認・調整
3. **データ転記実行**: PPformat・MFformatシートに自動転記・加工
4. **結果確認**: 各シートとログで処理結果を確認

## 開発

### 開発フロー

```bash
# 1. ローカルで開発
code コード.js

# 2. GASにプッシュ
clasp push

# 3. GASエディタで確認
clasp open

# 4. Git管理
git add .
git commit -m "feat: 新機能追加"
git push
```

### 設定管理

- PropertiesServiceによる設定の永続化
- デフォルト設定の自動適用
- 設定値の検証機能

## トラブルシューティング

### よくある問題

**Q: シートが見つからないエラー**
- A: LPcsv, PPformat, MFformatシートが存在するか確認

**Q: 設定が保存されない**
- A: スプレッドシートの編集権限があるか確認

**Q: 処理が途中で止まる**
- A: データ量が制限を超えていないか確認
- A: ログシートでエラー詳細を確認

## ライセンス

MIT License

## 貢献

バグ報告や機能改善の提案は [Issue](https://github.com/make-it-at/csv-data-transformer-gas/issues) でお知らせください。

## 更新履歴

- v1.0.0 (予定): 初期リリース
  - CSVインポート機能
  - データ転記・加工機能
  - 設定管理機能
  - HTMLサイドバーUI
  - ログ記録機能