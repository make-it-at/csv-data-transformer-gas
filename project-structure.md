# プロジェクト構造設計

## GAS プロジェクト構成

### ファイル構成
```
GAS Project Root/
├── コード.js (main.gs)           # メイン処理・エントリーポイント
├── csvProcessor.gs              # CSVインポート・データ処理
├── dataTransformer.gs           # データ転記・加工機能
├── settingsManager.gs           # 加工設定の管理
├── logger.gs                    # ログ記録・管理
├── utilities.gs                 # 共通ユーティリティ
├── sidebar.html                 # HTMLサイドバーUI
├── styles.html                  # CSSスタイル（HTMLに埋め込み）
└── appsscript.json             # GAS設定ファイル
```

### ローカル開発構成
```
Local Project/
├── コード.js                    # → main.gs (GASにプッシュ時)
├── csvProcessor.gs              # CSV処理機能
├── dataTransformer.gs           # データ変換機能
├── settingsManager.gs           # 設定管理機能
├── logger.gs                    # ログ機能
├── utilities.gs                 # ユーティリティ機能
├── sidebar.html                 # サイドバーUI
├── appsscript.json             # GAS設定
├── .clasp.json                 # clasp設定（Gitignore）
├── .claspignore                # clasp除外設定
├── requirements.md             # 要件定義書
├── technical-design.md         # 技術設計書
├── project-structure.md        # このファイル
└── README.md                   # プロジェクト説明
```

## シート構成

### 必須シート
1. **LPcsv** - CSVインポート先シート
   - 目的: 外部CSVファイルの直接インポート先
   - 構造: CSVファイルの内容をそのまま保持
   - ヘッダー: CSVファイルの1行目を自動検出

2. **PPformat** - 加工データシート1
   - 目的: LPcsvデータの加工版（パターン1）
   - 構造: 設定に基づいてLPcsvから転記・加工
   - ヘッダー: 設定で定義された列構成

3. **MFformat** - 加工データシート2
   - 目的: LPcsvデータの加工版（パターン2）
   - 構造: 設定に基づいてLPcsvから転記・加工
   - ヘッダー: 設定で定義された列構成

### 自動作成シート
4. **ログ** - 処理履歴記録シート
   - 目的: 全ての処理履歴を記録
   - 構造: タイムスタンプ | レベル | カテゴリ | メッセージ | 詳細データ
   - 管理: 1000行を超えたら古いログを自動削除

## 設定管理

### PropertiesService使用
```javascript
// 設定の保存・取得にPropertiesServiceを使用
const PROPERTY_KEYS = {
  PP_COLUMN_MAPPING: 'PP_COLUMN_MAPPING',
  MF_COLUMN_MAPPING: 'MF_COLUMN_MAPPING',
  DATA_TRANSFORM_RULES: 'DATA_TRANSFORM_RULES',
  FILTER_CONDITIONS: 'FILTER_CONDITIONS'
};

// 設定保存例
PropertiesService.getScriptProperties().setProperty(
  PROPERTY_KEYS.PP_COLUMN_MAPPING, 
  JSON.stringify(columnMapping)
);
```

### デフォルト設定
```javascript
const DEFAULT_SETTINGS = {
  ppColumnMapping: {
    // LPcsv列名 → PPformat列名
    'id': 'ID',
    'name': '名前',
    'price': '価格'
  },
  mfColumnMapping: {
    // LPcsv列名 → MFformat列名
    'id': '商品ID',
    'name': '商品名',
    'price': '販売価格'
  },
  transformRules: {
    // データ変換ルール
    'price': {
      type: 'number',
      format: '¥#,##0'
    }
  },
  filterConditions: {
    // フィルタリング条件
    'price': {
      operator: '>',
      value: 0
    }
  }
};
```

## UI設計

### HTMLサイドバー構成
```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    /* ミニマルデザインCSS */
  </style>
</head>
<body>
  <!-- 1. CSVインポートセクション -->
  <section id="import-section">
    <h3>CSVインポート</h3>
    <input type="file" id="csv-file" accept=".csv">
    <button id="import-btn">インポート実行</button>
  </section>

  <!-- 2. 加工設定セクション -->
  <section id="settings-section">
    <h3>加工設定</h3>
    <details>
      <summary>PPformat設定</summary>
      <!-- 列マッピング設定UI -->
    </details>
    <details>
      <summary>MFformat設定</summary>
      <!-- 列マッピング設定UI -->
    </details>
  </section>

  <!-- 3. 処理実行セクション -->
  <section id="process-section">
    <button id="process-btn">データ転記・加工実行</button>
    <div id="progress-bar"></div>
  </section>

  <!-- 4. ログ・ステータス表示 -->
  <section id="status-section">
    <h3>処理状況</h3>
    <div id="status-message"></div>
    <div id="log-viewer"></div>
  </section>

  <script>
    /* クライアントサイドJavaScript */
  </script>
</body>
</html>
```

### ミニマルデザイン方針
- **カラー**: 白ベース、アクセントは青系
- **フォント**: システムフォント使用
- **レイアウト**: シンプルなセクション分け
- **インタラクション**: 最小限のアニメーション
- **レスポンシブ**: 320px固定幅

## 開発・デプロイフロー

### 開発フロー
```bash
# 1. ローカルで開発
code コード.js

# 2. GASにプッシュ
clasp push

# 3. GASエディタで確認
clasp open

# 4. スプレッドシートでテスト
# ブラウザでスプレッドシートを開いて動作確認

# 5. ローカルでGit管理
git add .
git commit -m "feat: 新機能追加"
```

### ファイル同期ルール
- **アップロード対象**: .gs ファイル、.html ファイル、appsscript.json
- **除外対象**: .md ファイル、.clasp.json、Git関連ファイル
- **命名規則**: GAS側では日本語ファイル名（コード.js）も使用可能

### バージョン管理
```bash
# clasp でのバージョン管理
clasp version "v1.0.0 - 初期リリース"

# Git でのバージョン管理
git tag v1.0.0
git push origin v1.0.0
```

## セキュリティ・権限設定

### 必要な権限
```json
// appsscript.json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

### データ保護
- 設定データはPropertiesServiceで暗号化保存
- ログデータは定期的に自動削除
- ファイルアップロードは検証後に処理

## パフォーマンス最適化

### 制限値
```javascript
const LIMITS = {
  MAX_FILE_SIZE: 10 * 1024 * 1024,    // 10MB
  MAX_ROWS: 10000,                     // 10,000行
  MAX_COLUMNS: 100,                    // 100列
  BATCH_SIZE: 1000,                    // バッチ処理サイズ
  LOG_RETENTION: 1000,                 // ログ保持行数
  EXECUTION_TIMEOUT: 5 * 60 * 1000     // 5分
};
```

### 最適化戦略
1. **バッチ処理**: 大量データを分割処理
2. **一括API呼び出し**: getRange().setValues()を使用
3. **メモリ管理**: 不要オブジェクトの早期解放
4. **プログレス表示**: ユーザー体験向上