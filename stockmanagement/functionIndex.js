// functionIndex.gs
// Version 1.0
// Last updated: 2025-03-02
// 関数インデックス - 各ファイルに含まれる関数の一覧

/**
 * 各ファイルの関数一覧
 * 
 * このファイルは、証券管理シートのGASプロジェクト内の
 * 各ファイルに含まれる関数の一覧を管理します。
 */

// ========================================
// mainController.js (メイン制御)
// ========================================
/*
主要な関数:
- updatePF() - 保有株リスト更新
- updateTrendInfo() - トレンド情報更新
- updateReferenceInfo() - _Reference情報更新
- updateReferenceInfoForCodes(codes) - 特定銘柄の_Reference情報更新
- cancelCurrentProcess() - 処理キャンセル

ユーティリティ関数:
- UtilityManager.getStockPriceFormula(code, productType, row)
- UtilityManager.getNumberFormat(productTypeOrIsUSD)
- UtilityManager.applyFormatToCell(sheet, row, col, isUSD)
- ConfigManager.getUpdateConfig(options)
- ConfigManager.getOptimizedConfig(totalItems, complexity)
- ConfigManager.loadSavedConfig()
- ConfigManager.saveConfig(config)
- ConfigManager.handleError(functionName, error, customMessage)
- ConfigManager.validateSheet(sheet, sheetName)
- ConfigManager.validateDataExists(count, dataType)
- ConfigManager.startTimerExt(operationName)
- ConfigManager.checkpointExt(timer, checkpointName)
- ConfigManager.endTimerExt(timer, logResults)
- ConfigManager.analyzeBatchPerformanceExt(batchResults)

キャッシュ管理:
- CacheManager._generateKey(prefix, identifier)
- CacheManager.set(key, value, expirationSeconds)
- CacheManager.setMultiple(keyValueMap, expirationSeconds)
- CacheManager.get(key)
- CacheManager.getMultiple(keys)
- CacheManager.remove(key)
- CacheManager.setStockPrice(code, value)
- CacheManager.setMultipleStockPrices(codeValueMap)
- CacheManager.getStockPrice(code)
- CacheManager.setDailyItem(code, keyword, value)
- CacheManager.getDailyItem(code, keyword)
- CacheManager.setColumnIndices(sheetName, indices)
- CacheManager.getColumnIndices(sheetName)
- CacheManager.clearCache(prefix)

データ取得:
- DataFetcher.getScrapingSettings()
- DataFetcher.fetchStockInfo(code, settings)
- DataFetcher.extractValue(html, selector)
- DataFetcher.processMainInfoBatch(sheet, values, columnIndices, startRow, endRow)

グローバル関数:
- updatePF()
- updateTrendInfo()
- updateTrendInfoOptimized() - トレンド情報更新の最適化処理
- resumeUpdateTrendInfo() - トレンド情報更新の継続処理
- cancelCurrentProcess()
- onOpen()
- updateReferenceInfo()
- updateReferenceInfoForCodes(codes)
- testUpdatePF() - updatePF機能のテスト
*/

// ========================================
// uiManager.js (UI管理)
// ========================================
/*
主要な関数:
- UIManager.showSidebar() - サイドバー表示
- UIManager.createMenu() - メニュー作成
- UIManager.onOpen() - スプレッドシート開時処理
- UIManager.updateProgress(current, total, status, details) - 進捗更新
- UIManager.getProgressLogSheet() - 進捗ログシート取得
- UIManager.getLatestProgress() - 最新進捗取得
- UIManager.clearProgress() - 進捗クリア
- UIManager.showToast(message, title, timeout) - トースト表示
- UIManager.confirmAction(message) - 確認ダイアログ
- UIManager.validateActiveSheet() - アクティブシート検証

グローバル関数:
- onOpen()
- showSidebar()
- updateSidebarProgress(data)
- getLatestProgressFromProperties()
- clearProgress()
*/

// ========================================
// sheetManager.js (シート管理)
// ========================================
/*
主要な関数:
- SheetManager.getUntilBlank(sheet, startRow, startCol, numCols) - 空白まで取得
- SheetManager.getKeywords(settingSheet, frequency) - キーワード取得
- SheetManager.getKeywordsFromRow(sheet, row) - 行からキーワード取得
- SheetManager.getColumnIndices(sheet, keywords) - 列インデックス取得

グローバル関数:
- getKeywordsFromRow(sheet, row)
*/

// ========================================
// mainSheetManager.js (メインシート管理)
// ========================================
/*
主要な関数:
- MainSheetManager.updateSheet(sheetMain, stockData, columnIndices) - シート更新
- MainSheetManager.createMainData(stockData, columnIndices, lastColumn) - メインデータ作成
- MainSheetManager.setRequiredFields(rowData, stock, rowNum, columnIndices) - 必須フィールド設定
- MainSheetManager.writeMainData(sheetMain, mainData, startRow, columnIndices) - データ書き込み
- MainSheetManager.formatRow(sheet, row, columnIndices) - 行フォーマット
*/

// ========================================
// stockDataManager.js (株式データ管理)
// ========================================
/*
主要な関数:
- StockDataManager.processOrders(ordersData) - 注文データ処理
- StockDataManager.validateOrderData(row, columnIndices, rowIndex) - 注文データ検証
- StockDataManager.processStockData(stockData, code, type, num, unitPrice, accountType, productType, accountHolder) - 株式データ処理
- StockDataManager.updateStockData(stockData, code, type, num, unitPrice, accountType, productType, accountHolder) - 株式データ更新
- StockDataManager.validateStockData(stockData) - 株式データ検証
*/

// ========================================
// stockInfoFetcher.js (株式情報取得)
// ========================================
/*
主要な関数:
- StockInfoFetcher.fetchStockInfo(code) - 株式情報取得
- StockInfoFetcher.parseStockInfo(html, code) - 株式情報解析
- StockInfoFetcher.extractValue(html, selector) - 値抽出
*/

// ========================================
// stockPriceUtil.js (株価ユーティリティ)
// ========================================
/*
主要な関数:
- StockPriceUtil.getStockPrice(code, productType) - 株価取得
- StockPriceUtil.getMultipleStockPrices(codes, productType) - 複数株価取得
- StockPriceUtil.updateStockPrices(sheet, codes, productType) - 株価更新
- StockPriceUtil.formatStockPrice(price, productType) - 株価フォーマット
- StockPriceUtil.validateStockCode(code) - 銘柄コード検証
*/

// ========================================
// loggerManager.js (ログ管理)
// ========================================
/*
主要な関数:
- LoggerManager.writeLog(level, message, data) - ログ書き込み
- LoggerManager.getOrCreateLogSheet() - ログシート取得・作成
- LoggerManager.debug(message, data) - デバッグログ
- LoggerManager.info(message, data) - 情報ログ
- LoggerManager.warn(message, data) - 警告ログ
- LoggerManager.error(message, data) - エラーログ
- LoggerManager.clearLogs() - ログクリア

グローバル関数:
- testLogger() - ログテスト
- clearSystemLogs() - システムログクリア
*/

// ========================================
// timeoutManager.js (タイムアウト管理)
// ========================================
/*
主要な関数:
- TimeoutManager.setTimeout(operationName, timeoutMs) - タイムアウト設定
- TimeoutManager.clearTimeout(operationName) - タイムアウトクリア
- TimeoutManager.deleteAllTriggers() - 全トリガー削除
- TimeoutManager.cleanup() - クリーンアップ
*/

// ========================================
// performanceMonitor.js (パフォーマンス監視)
// ========================================
/*
主要な関数:
- PerformanceMonitor.startTimer(operationName) - タイマー開始
- PerformanceMonitor.endTimer(operationName) - タイマー終了
- PerformanceMonitor.getPerformanceStats() - パフォーマンス統計取得
- PerformanceMonitor.logPerformance(operationName, duration) - パフォーマンスログ
*/

// ========================================
// debugUtils.js (デバッグユーティリティ)
// ========================================
/*
主要な関数:
- DebugUtils.logObject(obj, name) - オブジェクトログ
- DebugUtils.logArray(arr, name) - 配列ログ
- DebugUtils.logFunctionCall(functionName, args) - 関数呼び出しログ
- DebugUtils.validateData(data, schema) - データ検証
*/

// ========================================
// dataFetcher.js (データ取得)
// ========================================
/*
主要な関数:
- DataFetcher.fetchUrl(url, options) - URL取得
- DataFetcher.fetchWithRetry(url, options, maxRetries) - リトライ付き取得
- DataFetcher.parseResponse(response) - レスポンス解析
*/

// ========================================
// asyncUtils.js (非同期処理)
// ========================================
/*
主要な関数:
- AsyncUtils.processInBatches(items, processor, maxConcurrent, delayBetweenBatches) - バッチ処理
- AsyncUtils.createBatches(items, batchSize) - バッチ作成
- AsyncUtils.delay(ms) - 遅延処理
- AsyncUtils.shouldCancel() - キャンセル判定
- AsyncUtils.throttle(func, limit) - スロットリング
*/

// ========================================
// ChartURL.js (チャートURL)
// ========================================
/*
主要な関数:
- ChartURL.generateChartUrl(data, options) - チャートURL生成
- ChartURL.getImageUrlChart() - チャート画像URL取得
*/

// ========================================
// logViewer.js (ログ表示)
// ========================================
/*
主要な関数:
- showSystemLogs() - システムログ表示
- showProgressLogs() - 進捗ログ表示
- openLogSheets() - ログシートを開く
- clearAllLogs() - 全ログクリア
*/

// ========================================
// メニューから実行可能な関数
// ========================================
/*
ポートフォリオ管理メニュー:
- 情報更新パネル → showSidebar()
- メイン情報更新 → updateMainInfo() (未実装)
- 保有株リスト更新 → updatePF()
- _Reference情報更新 → updateReferenceInfo()
- 特定銘柄の_Reference情報を更新 → showCodeInputDialog() (未実装)
- ダッシュボード表示 → showDashboard() (未実装)
- チャートURL更新 → getImageUrlChart()

キャッシュ管理サブメニュー:
- すべてのキャッシュをクリア → clearAllCache() (未実装)
- 株価キャッシュをクリア → clearStockCache() (未実装)
- 日次更新項目キャッシュをクリア → clearDailyCache() (未実装)
- 列インデックスキャッシュをクリア → clearColumnsCache() (未実装)

その他:
- 処理をキャンセル → cancelCurrentProcess()

テストツールサブメニュー:
- 単一銘柄株価取得テスト → testStockPrice() (未実装)
- 複数銘柄株価バッチテスト → batchTestStockPrices() (未実装)
- 代替ソースから株価取得 → getStockPriceFromAlternative() (未実装)
*/

/**
 * 関数検索ヘルパー
 * @param {string} functionName - 検索する関数名
 * @returns {string} 関数の場所
 */
function findFunction(functionName) {
  const functionMap = {
    // メイン機能
    'updatePF': 'mainController.js',
    'updateTrendInfo': 'mainController.js',
    'updateReferenceInfo': 'mainController.js',
    'cancelCurrentProcess': 'mainController.js',
    
    // UI機能
    'showSidebar': 'uiManager.js',
    'onOpen': 'uiManager.js',
    'updateProgress': 'uiManager.js',
    
    // ログ機能
    'testLogger': 'loggerManager.js',
    'clearSystemLogs': 'loggerManager.js',
    'showSystemLogs': 'logViewer.js',
    'showProgressLogs': 'logViewer.js',
    'openLogSheets': 'logViewer.js',
    'clearAllLogs': 'logViewer.js',
    
    // データ管理
    'processOrders': 'stockDataManager.js',
    'fetchStockInfo': 'stockInfoFetcher.js',
    'getStockPrice': 'stockPriceUtil.js',
    
    // シート管理
    'getUntilBlank': 'sheetManager.js',
    'getKeywords': 'sheetManager.js',
    'updateSheet': 'mainSheetManager.js'
  };
  
  return functionMap[functionName] || '見つかりません';
}

/**
 * ファイル別関数一覧を表示
 */
function showFunctionIndex() {
  const message = `
証券管理シート - 関数インデックス

【主要ファイルと関数】

1. mainController.js (メイン制御)
   - updatePF() - 保有株リスト更新
   - updateTrendInfo() - トレンド情報更新
   - updateReferenceInfo() - _Reference情報更新
   - cancelCurrentProcess() - 処理キャンセル

2. uiManager.js (UI管理)
   - showSidebar() - サイドバー表示
   - onOpen() - スプレッドシート開時処理
   - updateProgress() - 進捗更新

3. loggerManager.js (ログ管理)
   - testLogger() - ログテスト
   - clearSystemLogs() - システムログクリア

4. logViewer.js (ログ表示)
   - showSystemLogs() - システムログ表示
   - showProgressLogs() - 進捗ログ表示
   - openLogSheets() - ログシートを開く
   - clearAllLogs() - 全ログクリア

5. stockDataManager.js (株式データ管理)
   - processOrders() - 注文データ処理

6. stockPriceUtil.js (株価ユーティリティ)
   - getStockPrice() - 株価取得

【実行方法】
- スプレッドシートのメニュー「ポートフォリオ管理」から実行
- または、スクリプトエディタで直接実行
  `;
  
  SpreadsheetApp.getUi().alert('関数インデックス', message, SpreadsheetApp.getUi().ButtonSet.OK);
} 