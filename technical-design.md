# CSV処理・データ変換ツール 技術設計書

## アーキテクチャ概要

### システム構成
```
┌─────────────────────────────────────────────────────────────────┐
│                    Google Spreadsheet                          │
│  ┌─────────────────────┐  ┌─────────────────────────────────┐   │
│  │    HTML Sidebar     │  │         Sheets                  │   │
│  │  ┌─────────────────┐│  │  ┌─────────────────────────┐   │   │
│  │  │ CSV Import      ││  │  │      LPcsv              │   │   │
│  │  │ Settings Panel  ││  │  │  (CSV Import Target)    │   │   │
│  │  │ Processing Btn  ││  │  └─────────────────────────┘   │   │
│  │  │ Log Viewer      ││  │  ┌─────────────────────────┐   │   │
│  │  │ Status Display  ││  │  │      PPformat           │   │   │
│  │  └─────────────────┘│  │  │  (Processed Data 1)     │   │   │
│  └─────────────────────┘  │  └─────────────────────────┘   │   │
│                           │  ┌─────────────────────────┐   │   │
│                           │  │      MFformat           │   │   │
│                           │  │  (Processed Data 2)     │   │   │
│                           │  └─────────────────────────┘   │   │
│                           │  ┌─────────────────────────┐   │   │
│                           │  │        ログ             │   │   │
│                           │  │  (Auto-created Log)     │   │   │
│                           │  └─────────────────────────┘   │   │
│  └─────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
                    │
                    ▼
┌─────────────────────────────────────────┐
│        Google Apps Script              │
│  ┌─────────────────────────────────┐   │
│  │        main.gs                  │   │
│  │  - showSidebar()               │   │
│  │  - processCSVImport()          │   │
│  │  - processCSVExport()          │   │
│  └─────────────────────────────────┘   │
│  ┌─────────────────────────────────┐   │
│  │      csvImporter.gs             │   │
│  │  - parseCSVContent()           │   │
│  │  - validateCSVData()           │   │
│  │  - saveToSheet()               │   │
│  └─────────────────────────────────┘   │
│  ┌─────────────────────────────────┐   │
│  │      csvExporter.gs             │   │
│  │  - readSheetData()             │   │
│  │  - convertToCSV()              │   │
│  │  - downloadCSV()               │   │
│  └─────────────────────────────────┘   │
│  ┌─────────────────────────────────┐   │
│  │       utilities.gs              │   │
│  │  - showMessage()               │   │
│  │  - logError()                  │   │
│  │  - validateInput()             │   │
│  └─────────────────────────────────┘   │
└─────────────────────────────────────────┘
```

## ファイル構成と責務

### 1. main.gs (コード.js)
**責務**: メイン制御・UI表示・シート管理
```javascript
// 主要関数
- showSidebar(): HTMLサイドバーの表示
- onOpen(): スプレッドシート開時のメニュー追加
- initializeSheets(): 必要シートの存在確認・作成
- createLogSheet(): ログシートの自動作成
- processCSVImport(fileContent): CSVインポート処理の制御
- processDataTransfer(): データ転記・加工処理の制御
```

### 2. csvProcessor.gs
**責務**: CSVインポート・データ処理
```javascript
// 主要関数
- parseCSVContent(content, options): CSV文字列の解析
- validateCSVData(data): データ検証
- importToLPcsv(data): LPcsvシートへの保存
- detectEncoding(content): 文字コード自動判定
- extractHeaders(data): ヘッダー行の抽出・分析
```

### 3. dataTransformer.gs
**責務**: データ転記・加工機能
```javascript
// 主要関数
- transferToPPformat(sourceData, settings): PPformatシートへの転記
- transferToMFformat(sourceData, settings): MFformatシートへの転記
- applyColumnMapping(data, mapping): 列マッピング適用
- applyDataTransforms(data, rules): データ変換ルール適用
- applyFilters(data, conditions): フィルタリング条件適用
```

### 4. settingsManager.gs
**責務**: 加工設定の管理
```javascript
// 主要関数
- getProcessingSettings(): 現在の加工設定取得
- saveProcessingSettings(settings): 加工設定保存
- getDefaultSettings(): デフォルト設定取得
- validateSettings(settings): 設定値検証
- getColumnMappings(): 列マッピング設定取得
```

### 5. logger.gs
**責務**: ログ記録・管理
```javascript
// 主要関数
- writeLog(level, category, message, data): ログ記録
- getRecentLogs(count): 最新ログ取得
- clearOldLogs(): 古いログの削除
- formatLogEntry(entry): ログエントリのフォーマット
- initializeLogSheet(): ログシート初期化
```

### 6. utilities.gs
**責務**: 共通ユーティリティ
```javascript
// 主要関数
- showMessage(type, message): メッセージ表示
- validateInput(input, rules): 入力値検証
- formatFileSize(bytes): ファイルサイズフォーマット
- getTimestamp(): タイムスタンプ生成
- getSheetSafely(name): シート取得（エラーハンドリング付き）
```

### 7. sidebar.html
**責務**: ミニマルUIサイドバー
```html
<!-- 主要セクション -->
- CSVインポートセクション
- 加工設定パネル（折りたたみ式）
- 処理実行ボタン
- ログ表示エリア
- ステータス・進捗表示
```

## データフロー

### 全体処理フロー
```
1. CSVファイル選択・インポート
   ↓
2. LPcsvシートに保存
   ↓
3. ヘッダー分析・設定確認
   ↓
4. PPformat・MFformatシートに転記・加工
   ↓
5. 処理結果をログシートに記録
```

### 詳細データフロー

#### 1. CSVインポートフロー
```
[HTML Sidebar]
1. ユーザーがファイル選択
   ↓
2. FileReader APIでファイル読み込み
   ↓
[GAS - main.gs]
3. processCSVImport(fileContent)呼び出し
   ↓
[GAS - csvProcessor.gs]
4. parseCSVContent()でCSV解析
5. validateCSVData()でデータ検証
6. extractHeaders()でヘッダー抽出
7. importToLPcsv()でLPcsvシートに保存
   ↓
[GAS - logger.gs]
8. writeLog()で処理ログ記録
   ↓
[HTML Sidebar]
9. 結果表示・次ステップ案内
```

#### 2. データ転記・加工フロー
```
[HTML Sidebar]
1. 加工設定確認・調整
2. 処理実行ボタンクリック
   ↓
[GAS - main.gs]
3. processDataTransfer()呼び出し
   ↓
[GAS - settingsManager.gs]
4. getProcessingSettings()で設定取得
   ↓
[GAS - dataTransformer.gs]
5. LPcsvシートからデータ読み取り
6. applyColumnMapping()で列マッピング
7. applyDataTransforms()でデータ変換
8. applyFilters()でフィルタリング
9. transferToPPformat()でPPformatに転記
10. transferToMFformat()でMFformatに転記
   ↓
[GAS - logger.gs]
11. writeLog()で処理ログ記録
   ↓
[HTML Sidebar]
12. 処理完了・結果表示
```

#### 3. 設定管理フロー
```
[HTML Sidebar]
1. 設定パネルで設定変更
   ↓
[GAS - settingsManager.gs]
2. saveProcessingSettings()で設定保存
3. validateSettings()で設定検証
   ↓
[GAS - main.gs]
4. 設定変更をUI側に反映
```

## エラーハンドリング設計

### エラー分類
1. **入力エラー**: ファイル形式・サイズ・内容の問題
2. **処理エラー**: データ変換・保存時の問題
3. **システムエラー**: GAS制限・権限・ネットワークの問題

### エラー処理方針
```javascript
// 統一エラーハンドリング
function handleError(error, context) {
  // 1. ログ記録
  logError(context, error);
  
  // 2. ユーザー向けメッセージ生成
  const userMessage = generateUserMessage(error);
  
  // 3. UI表示
  showMessage('error', userMessage);
  
  // 4. 必要に応じて復旧処理
  if (isRecoverable(error)) {
    attemptRecovery(error, context);
  }
}
```

## パフォーマンス設計

### 制限値設定
```javascript
const LIMITS = {
  MAX_FILE_SIZE: 10 * 1024 * 1024,    // 10MB
  MAX_ROWS: 10000,                     // 10,000行
  MAX_COLUMNS: 100,                    // 100列
  MAX_CELL_LENGTH: 50000,              // 50,000文字
  BATCH_SIZE: 1000,                    // バッチ処理サイズ
  TIMEOUT_MS: 5 * 60 * 1000           // 5分
};
```

### 最適化戦略
1. **バッチ処理**: 大量データを分割して処理
2. **メモリ管理**: 不要なオブジェクトの早期解放
3. **API呼び出し最小化**: SpreadsheetApp.getRange()の回数削減
4. **プログレス表示**: 長時間処理時のユーザー体験向上

## セキュリティ設計

### 入力検証
```javascript
// CSVファイル検証
function validateCSVFile(file) {
  // ファイル拡張子チェック
  if (!file.name.match(/\.(csv|txt)$/i)) {
    throw new Error('CSVファイルを選択してください');
  }
  
  // ファイルサイズチェック
  if (file.size > LIMITS.MAX_FILE_SIZE) {
    throw new Error(`ファイルサイズが制限を超えています: ${formatFileSize(LIMITS.MAX_FILE_SIZE)}`);
  }
  
  // 内容チェック
  validateCSVContent(file.content);
}
```

### 権限管理
- スプレッドシート編集権限の確認
- 実行権限の適切な設定
- 機密データの適切な処理

## テスト設計

### テストケース分類
1. **正常系テスト**
   - 標準的なCSVファイルのインポート
   - データのエクスポート
   - UI操作

2. **異常系テスト**
   - 不正なファイル形式
   - サイズ制限超過
   - 権限不足

3. **境界値テスト**
   - 最大ファイルサイズ
   - 最大行数・列数
   - 特殊文字を含むデータ

### テスト実装
```javascript
// tests/test.gs
function runAllTests() {
  testCSVImport();
  testCSVExport();
  testErrorHandling();
  testValidation();
}

function testCSVImport() {
  // テストデータ準備
  const testCSV = "name,age,city\nJohn,25,Tokyo\nJane,30,Osaka";
  
  // 実行
  const result = processCSVImport(testCSV);
  
  // 検証
  if (!result.success) {
    throw new Error('CSV import test failed');
  }
}
```

## デプロイメント設計

### 開発フロー
```bash
# 1. ローカル開発
code src/main.gs

# 2. テスト実行
clasp run runAllTests

# 3. GASにアップロード
clasp push

# 4. 本番テスト
# GASエディタまたはスプレッドシートで動作確認

# 5. バージョン管理
git add .
git commit -m "feat: 新機能追加"
git push
```

### 設定管理
```javascript
// 環境別設定
const CONFIG = {
  development: {
    debug: true,
    logLevel: 'DEBUG',
    batchSize: 100
  },
  production: {
    debug: false,
    logLevel: 'INFO',
    batchSize: 1000
  }
};
```

## 運用・保守設計

### ログ設計
```javascript
// ログレベル
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3
};

// ログ出力
function log(level, message, data = null) {
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    level,
    message,
    data
  };
  
  // Console.log出力
  console.log(JSON.stringify(logEntry));
  
  // 必要に応じてスプレッドシートにも記録
  if (level >= LOG_LEVELS.WARN) {
    recordToSheet(logEntry);
  }
}
```

### 監視項目
- 処理時間
- エラー発生率
- ファイルサイズ分布
- 利用頻度

### バックアップ戦略
- 重要なスプレッドシートの定期バックアップ
- 設定データの保存
- エラーログの保持