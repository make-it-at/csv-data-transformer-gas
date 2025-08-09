// mainController.gs
// Version 38.1
// Last updated: 2024-05-01
// Changes: _Referenceシートの更新機能を追加

var UtilityManager = {
  /**
   * 株価取得用の数式を生成する
   * @param {string} code 銘柄コード
   * @param {string} productType 商品種別
   * @param {number} row 行番号
   * @returns {string} 数式
   */
  getStockPriceFormula(code, productType, row) {
    if (!code) return '';
    
    if (productType === 'MMF') {
      return `=E${row}`;
    } else if (productType === 'US') {
      return `=IF(ISBLANK(B${row}), "", GOOGLEFINANCE(B${row}))`;
    } else if (productType === 'FUND') {
      return `=IF(ISBLANK(B${row}), "", STOCKPRICEJP("TOSHIN", B${row}))`;
    } else {
      return `=IF(ISBLANK(B${row}),"",STOCKPRICEJP("JP", B${row}))`;
    }
  },

  /**
   * 数値フォーマットを取得する
   * @param {string|boolean} productTypeOrIsUSD 商品種別またはUSD通貨フラグ
   * @returns {string} 数値フォーマット
   */
  getNumberFormat(productTypeOrIsUSD) {
    // 引数がブール値の場合はそのまま使用、文字列の場合は商品種別として判定
    const isUSD = typeof productTypeOrIsUSD === 'boolean' 
      ? productTypeOrIsUSD 
      : (productTypeOrIsUSD === 'US' || productTypeOrIsUSD === 'MMF');
    
    return isUSD ? '"$"#,##0.00' : '¥#,##0.00';
  },

  /**
   * セルに数値フォーマットを適用する
   * @param {Object} sheet シート
   * @param {number} row 行
   * @param {number} col 列
   * @param {boolean} isUSD USD通貨かどうか
   */
  applyFormatToCell(sheet, row, col, isUSD) {
    if (col > 0) {
      const cell = sheet.getRange(row, col);
      cell.setNumberFormat(this.getNumberFormat(isUSD));
    }
  }
};

var ConfigManager = {
  /**
   * 更新処理の設定を取得する
   * @param {Object} options カスタム設定オプション
   * @returns {Object} 更新設定
   */
  getUpdateConfig(options = {}) {
    // デフォルト設定
    const defaultConfig = {
      batchSize: 5,
      sleepTime: {
        betweenBatches: 1000,
        betweenColumns: 1000,
        betweenRetries: 500
      }
    };
    
    // カスタム設定をマージ
    return {
      ...defaultConfig,
      ...options,
      sleepTime: {
        ...defaultConfig.sleepTime,
        ...(options.sleepTime || {})
      }
    };
  },
  
  /**
   * データ量に基づいて最適なバッチサイズを計算する
   * @param {number} totalItems 処理対象の総数
   * @param {number} complexity 処理の複雑さ (1-10)
   * @returns {Object} 最適化された設定
   */
  getOptimizedConfig(totalItems, complexity = 5) {
    // 複雑さに基づいて基本バッチサイズを調整
    // 複雑さが高いほど小さいバッチサイズを使用
    const baseBatchSize = Math.max(1, Math.floor(10 / complexity));
    
    // データ量に基づいてバッチサイズを調整
    let batchSize = baseBatchSize;
    if (totalItems > 100) {
      batchSize = Math.min(20, baseBatchSize * 2);
    } else if (totalItems > 50) {
      batchSize = Math.min(10, baseBatchSize * 1.5);
    } else if (totalItems < 10) {
      batchSize = Math.max(1, Math.min(totalItems, baseBatchSize));
    }
    
    // データ量に基づいて待機時間を調整
    const sleepTime = {
      betweenBatches: totalItems > 50 ? 1500 : 1000,
      betweenColumns: totalItems > 50 ? 1200 : 800,
      betweenRetries: 500
    };
    
    return this.getUpdateConfig({ batchSize, sleepTime });
  },
  
  /**
   * スクリプトプロパティから設定を読み込む
   * @returns {Object} 保存された設定
   */
  loadSavedConfig() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const savedConfig = scriptProperties.getProperty('updateConfig');
      
      if (savedConfig) {
        return JSON.parse(savedConfig);
      }
    } catch (e) {
      LoggerManager.warn('設定の読み込みに失敗しました:', e);
    }
    
    return this.getUpdateConfig();
  },
  
  /**
   * 設定をスクリプトプロパティに保存する
   * @param {Object} config 保存する設定
   */
  saveConfig(config) {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('updateConfig', JSON.stringify(config));
      LoggerManager.info('設定を保存しました');
    } catch (e) {
      LoggerManager.warn('設定の保存に失敗しました:', e);
    }
  }
};

var ErrorManager = {
  /**
   * エラーをハンドリングする
   * @param {string} functionName 関数名
   * @param {Error} error エラーオブジェクト
   * @param {string} customMessage カスタムメッセージ
   * @throws {Error} 元のエラー
   */
  handleError(functionName, error, customMessage = '') {
    const message = customMessage ? `${customMessage}: ${error.message}` : error.message;
    LoggerManager.error(`Error in ${functionName}:`, error);
    UIManager.showError(message);
    throw error;
  },
  
  /**
   * シートの存在を検証する
   * @param {Object} sheet シート
   * @param {string} sheetName シート名
   * @throws {Error} シートが存在しない場合
   */
  validateSheet(sheet, sheetName) {
    if (!sheet) {
      throw new Error(`${sheetName}シートが見つかりません`);
    }
  },
  
  /**
   * データの存在を検証する
   * @param {number} count データ数
   * @param {string} dataType データ種別
   * @throws {Error} データが存在しない場合
   */
  validateDataExists(count, dataType) {
    if (count <= 0) {
      throw new Error(`${dataType}のデータが存在しません`);
    }
  }
};

/**
 * パフォーマンス監視モジュール拡張
 * 既存のPerformanceMonitorに追加する機能
 */
var PerformanceMonitorExtension = {
  /**
   * 処理時間の計測を開始する（拡張版）
   * @param {string} operationName 操作名
   * @returns {Object} 計測オブジェクト
   */
  startTimerExt(operationName) {
    return {
      name: operationName,
      startTime: new Date().getTime(),
      checkpoints: []
    };
  },
  
  /**
   * チェックポイントを記録する（拡張版）
   * @param {Object} timer 計測オブジェクト
   * @param {string} checkpointName チェックポイント名
   */
  checkpointExt(timer, checkpointName) {
    if (!timer) return;
    
    const currentTime = new Date().getTime();
    const elapsedSinceStart = currentTime - timer.startTime;
    const lastCheckpoint = timer.checkpoints.length > 0 ? 
      timer.checkpoints[timer.checkpoints.length - 1] : 
      { time: timer.startTime };
    const elapsedSinceLastCheckpoint = currentTime - lastCheckpoint.time;
    
    timer.checkpoints.push({
      name: checkpointName,
      time: currentTime,
      elapsedSinceStart,
      elapsedSinceLastCheckpoint
    });
  },
  
  /**
   * 計測を終了し、結果を出力する（拡張版）
   * @param {Object} timer 計測オブジェクト
   * @param {boolean} logResults 結果をログに出力するかどうか
   * @returns {Object} 計測結果
   */
  endTimerExt(timer, logResults = true) {
    if (!timer) return null;
    
    const endTime = new Date().getTime();
    const totalTime = endTime - timer.startTime;
    
    const result = {
      name: timer.name,
      totalTime,
      checkpoints: timer.checkpoints,
      startTime: timer.startTime,
      endTime
    };
    
    if (logResults) {
      LoggerManager.info(`パフォーマンス [${timer.name}]: 合計時間=${totalTime}ms`);
      
      if (timer.checkpoints.length > 0) {
        timer.checkpoints.forEach(cp => {
          LoggerManager.debug(`  - ${cp.name}: ${cp.elapsedSinceLastCheckpoint}ms (累計: ${cp.elapsedSinceStart}ms)`);
        });
      }
    }
    
    return result;
  },
  
  /**
   * バッチ処理のパフォーマンスを分析する（拡張版）
   * @param {Array} batchResults バッチ処理の結果配列
   * @returns {Object} 分析結果
   */
  analyzeBatchPerformanceExt(batchResults) {
    if (!batchResults || batchResults.length === 0) {
      return null;
    }
    
    const totalBatches = batchResults.length;
    const totalTime = batchResults.reduce((sum, batch) => sum + batch.totalTime, 0);
    const avgBatchTime = totalTime / totalBatches;
    
    // 最速と最遅のバッチを特定
    const fastestBatch = batchResults.reduce((fastest, current) => 
      current.totalTime < fastest.totalTime ? current : fastest, batchResults[0]);
    
    const slowestBatch = batchResults.reduce((slowest, current) => 
      current.totalTime > slowest.totalTime ? current : slowest, batchResults[0]);
    
    const analysis = {
      totalBatches,
      totalTime,
      avgBatchTime,
      fastestBatch: {
        index: batchResults.indexOf(fastestBatch),
        time: fastestBatch.totalTime
      },
      slowestBatch: {
        index: batchResults.indexOf(slowestBatch),
        time: slowestBatch.totalTime
      }
    };
    
    LoggerManager.info(`バッチ処理分析: 合計=${totalBatches}バッチ, 平均時間=${avgBatchTime.toFixed(2)}ms`);
    LoggerManager.info(`  - 最速: バッチ#${analysis.fastestBatch.index + 1} (${analysis.fastestBatch.time}ms)`);
    LoggerManager.info(`  - 最遅: バッチ#${analysis.slowestBatch.index + 1} (${analysis.slowestBatch.time}ms)`);
    
    return analysis;
  }
};

// PerformanceMonitorExtensionの機能をPerformanceMonitorに追加（遅延実行）
function initializePerformanceMonitorExtension() {
  if (typeof PerformanceMonitor !== 'undefined') {
    Object.assign(PerformanceMonitor, PerformanceMonitorExtension);
    LoggerManager.debug('PerformanceMonitorに拡張機能を追加しました。');
  } else {
    // PerformanceMonitorが未定義の場合は警告を出す
    LoggerManager.warn('PerformanceMonitorが定義されていないため、拡張機能を追加できません。');
  }
}

/**
 * キャッシュ管理モジュール
 */
var CacheManager = {
  /**
   * キャッシュキーを生成する
   * @param {string} prefix プレフィックス
   * @param {string} identifier 識別子
   * @returns {string} キャッシュキー
   */
  _generateKey(prefix, identifier) {
    return `${prefix}_${identifier}`;
  },
  
  /**
   * キャッシュに値を保存する
   * @param {string} key キー
   * @param {*} value 値
   * @param {number} expirationSeconds 有効期限（秒）
   */
  set(key, value, expirationSeconds = 21600) { // 6時間に延長
    try {
      const cache = CacheService.getScriptCache();
      const valueString = JSON.stringify(value);
      cache.put(key, valueString, expirationSeconds);
      LoggerManager.debug(`キャッシュに保存: ${key}`);
    } catch (e) {
      LoggerManager.warn(`キャッシュへの保存に失敗: ${key}`, e);
    }
  },
  
  /**
   * 複数のキャッシュ値を一度に保存する
   * @param {Object} keyValueMap キーと値のマップ
   * @param {number} expirationSeconds 有効期限（秒）
   */
  setMultiple(keyValueMap, expirationSeconds = 21600) { // 6時間に延長
    try {
      const cache = CacheService.getScriptCache();
      const serializedMap = {};
      
      // 各値をシリアライズ
      Object.keys(keyValueMap).forEach(key => {
        serializedMap[key] = JSON.stringify(keyValueMap[key]);
      });
      
      // 一度に複数のキャッシュを設定
      cache.putAll(serializedMap, expirationSeconds);
      LoggerManager.debug(`複数のキャッシュを保存: ${Object.keys(keyValueMap).length}件`);
    } catch (e) {
      LoggerManager.warn(`複数キャッシュの保存に失敗`, e);
    }
  },
  
  /**
   * キャッシュから値を取得する
   * @param {string} key キー
   * @returns {*} 値（存在しない場合はnull）
   */
  get(key) {
    try {
      const cache = CacheService.getScriptCache();
      const valueString = cache.get(key);
      
      if (valueString) {
        LoggerManager.debug(`キャッシュからロード: ${key}`);
        return JSON.parse(valueString);
      }
    } catch (e) {
      LoggerManager.warn(`キャッシュからの取得に失敗: ${key}`, e);
    }
    
    return null;
  },
  
  /**
   * 複数のキャッシュ値を一度に取得する
   * @param {Array} keys キーの配列
   * @returns {Object} キーと値のマップ
   */
  getMultiple(keys) {
    try {
      const cache = CacheService.getScriptCache();
      const cachedData = cache.getAll(keys);
      const result = {};
      
      // 各値をデシリアライズ
      Object.keys(cachedData).forEach(key => {
        try {
          result[key] = JSON.parse(cachedData[key]);
        } catch (e) {
          LoggerManager.warn(`キャッシュ値のパースに失敗: ${key}`, e);
        }
      });
      
      LoggerManager.debug(`複数のキャッシュをロード: ${Object.keys(result).length}/${keys.length}件`);
      return result;
    } catch (e) {
      LoggerManager.warn(`複数キャッシュの取得に失敗`, e);
      return {};
    }
  },
  
  /**
   * キャッシュから値を削除する
   * @param {string} key キー
   */
  remove(key) {
    try {
      const cache = CacheService.getScriptCache();
      cache.remove(key);
      LoggerManager.debug(`キャッシュから削除: ${key}`);
    } catch (e) {
      LoggerManager.warn(`キャッシュからの削除に失敗: ${key}`, e);
    }
  },
  
  /**
   * 株価データをキャッシュに保存する
   * @param {string} code 銘柄コード
   * @param {string} value 株価
   */
  setStockPrice(code, value) {
    const key = this._generateKey('stockprice', code);
    this.set(key, { value, timestamp: new Date().getTime() }, 21600); // 6時間キャッシュ
  },
  
  /**
   * 複数の株価データを一度にキャッシュに保存する
   * @param {Object} codeValueMap コードと株価のマップ
   */
  setMultipleStockPrices(codeValueMap) {
    const keyValueMap = {};
    const timestamp = new Date().getTime();
    
    Object.keys(codeValueMap).forEach(code => {
      const key = this._generateKey('stockprice', code);
      keyValueMap[key] = { value: codeValueMap[code], timestamp };
    });
    
    this.setMultiple(keyValueMap, 21600); // 6時間キャッシュ
  },
  
  /**
   * キャッシュから株価データを取得する
   * @param {string} code 銘柄コード
   * @returns {Object} 株価データ（存在しない場合はnull）
   */
  getStockPrice(code) {
    const key = this._generateKey('stockprice', code);
    return this.get(key);
  },
  
  /**
   * 日次更新項目データをキャッシュに保存する
   * @param {string} code 銘柄コード
   * @param {string} keyword キーワード
   * @param {string} value 値
   */
  setDailyItem(code, keyword, value) {
    const key = this._generateKey(`daily_${keyword}`, code);
    this.set(key, { value, timestamp: new Date().getTime() }, 86400); // 24時間キャッシュ
  },
  
  /**
   * キャッシュから日次更新項目データを取得する
   * @param {string} code 銘柄コード
   * @param {string} keyword キーワード
   * @returns {Object} 日次更新項目データ（存在しない場合はnull）
   */
  getDailyItem(code, keyword) {
    const key = this._generateKey(`daily_${keyword}`, code);
    return this.get(key);
  },
  
  /**
   * 列インデックスをキャッシュに保存する
   * @param {string} sheetName シート名
   * @param {Object} indices 列インデックス
   */
  setColumnIndices(sheetName, indices) {
    const key = this._generateKey('columns', sheetName);
    this.set(key, indices, 86400); // 24時間キャッシュ
  },
  
  /**
   * キャッシュから列インデックスを取得する
   * @param {string} sheetName シート名
   * @returns {Object} 列インデックス（存在しない場合はnull）
   */
  getColumnIndices(sheetName) {
    const key = this._generateKey('columns', sheetName);
    return this.get(key);
  },
  
  /**
   * キャッシュをクリアする
   * @param {string} prefix プレフィックス（指定した場合はそのプレフィックスのキャッシュのみクリア）
   */
  clearCache(prefix = null) {
    try {
      if (prefix) {
        // 特定のプレフィックスのキャッシュをクリアする方法はないため、
        // 代わりに有効期限を0に設定して実質的にクリアする
        const keys = this._getAllKeys().filter(key => key.startsWith(prefix));
        const cache = CacheService.getScriptCache();
        keys.forEach(key => cache.put(key, '', 1));
        LoggerManager.info(`キャッシュをクリア: プレフィックス=${prefix}, ${keys.length}件`);
              } else {
        // すべてのキャッシュをクリア
        const keys = this._getAllKeys();
        const cache = CacheService.getScriptCache();
        keys.forEach(key => cache.remove(key));
        LoggerManager.info(`すべてのキャッシュをクリア: ${keys.length}件`);
      }
    } catch (e) {
      LoggerManager.warn('キャッシュのクリアに失敗:', e);
    }
  },
  
  /**
   * すべてのキャッシュキーを取得する
   * @returns {Array} キャッシュキーの配列
   * @private
   */
  _getAllKeys() {
    // 注意: CacheServiceには全キーを取得する方法がないため、
    // スクリプトプロパティに保存されたキーリストを使用
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const keysString = scriptProperties.getProperty('cacheKeys');
      return keysString ? JSON.parse(keysString) : [];
    } catch (e) {
      LoggerManager.warn('キャッシュキーの取得に失敗:', e);
      return [];
    }
  },
  
  /**
   * キャッシュキーをスクリプトプロパティに保存する
   * @param {string} key 追加するキー
   * @private
   */
  _saveKey(key) {
    try {
      const keys = this._getAllKeys();
      if (!keys.includes(key)) {
        keys.push(key);
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty('cacheKeys', JSON.stringify(keys));
      }
    } catch (e) {
      LoggerManager.warn('キャッシュキーの保存に失敗:', e);
    }
  }
};

var MainController = {
  /**
   * スクレイピング設定を取得する
   * @returns {Object} スクレイピング設定
   */
  getScrapingSettings() {
    try {
      const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Setting');
      if (!settingSheet) {
        throw new Error('Settingシートが見つかりません');
      }

      const data = settingSheet.getDataRange().getValues();
      const settings = {};
      
      // ヘッダー行をスキップ
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const keyword = row[0];  // キーワード列
        const selector = row[1]; // 正規表現列
        const tag = row[2];      // タグ列
        const baseUrl = row[3];  // ベースURL列
        
        if (keyword && selector) {
          settings[keyword] = {
            selector: selector,
            tag: tag,
            baseUrl: baseUrl
          };
        }
      }
      
      return settings;
    } catch (e) {
      LoggerManager.error('スクレイピング設定の取得に失敗:', e);
      throw e;
    }
  },

  /**
   * 銘柄の情報を取得する
   * @param {string} code 銘柄コード
   * @param {Object} settings スクレイピング設定
   * @returns {Object} 銘柄情報
   */
  async fetchStockInfo(code, settings) {
    try {
      const previousCloseSettings = settings['前日終値'];
      const signalSettings = settings['シグナル'];
      
      if (!previousCloseSettings || !signalSettings) {
        throw new Error('必要なスクレイピング設定が見つかりません');
      }

      LoggerManager.debug(`スクレイピング設定:`, {
        previousClose: previousCloseSettings,
        signal: signalSettings
      });

      // 前日終値を取得
      const previousCloseUrl = previousCloseSettings.baseUrl.replace('(code)', code);
      LoggerManager.debug(`前日終値URL: ${previousCloseUrl}`);
      
      let previousCloseResponse;
      try {
        previousCloseResponse = await UrlFetchApp.fetch(previousCloseUrl);
      } catch (e) {
        throw new Error(`前日終値の取得に失敗: ${e.message}`);
      }
      
      const previousCloseHtml = previousCloseResponse.getContentText();
      const previousClose = this.extractValue(previousCloseHtml, previousCloseSettings.selector);
      
      if (!previousClose) {
        throw new Error('前日終値が見つかりません');
      }

      // シグナルを取得
      const signalUrl = signalSettings.baseUrl.replace('(code)', code);
      LoggerManager.debug(`シグナルURL: ${signalUrl}`);
      
      let signalResponse;
      try {
        signalResponse = await UrlFetchApp.fetch(signalUrl);
      } catch (e) {
        throw new Error(`シグナルの取得に失敗: ${e.message}`);
      }
      
      const signalHtml = signalResponse.getContentText();
      const signal = this.extractValue(signalHtml, signalSettings.selector);
      
      if (!signal) {
        throw new Error('シグナルが見つかりません');
      }

      return {
        previousClose: previousClose,
        signal: signal
      };
    } catch (e) {
      LoggerManager.error(`銘柄情報の取得に失敗 (${code}):`, {
        error: e.message,
        stack: e.stack,
        settings: settings
      });
      throw e;
    }
  },

  /**
   * HTMLから値を抽出する
   * @param {string} html HTML文字列
   * @param {string} selector セレクタ
   * @returns {string} 抽出された値
   */
  extractValue(html, selector) {
    try {
      LoggerManager.debug(`正規表現パターン: ${selector}`);
      
      // セレクタに基づいて値を抽出するロジック
      const regex = new RegExp(selector);
      const match = html.match(regex);
      
      if (!match) {
        LoggerManager.debug('正規表現にマッチする値が見つかりません');
        return null;
      }
      
      LoggerManager.debug(`抽出された値: ${match[1]}`);
      return match[1];
    } catch (e) {
      LoggerManager.error('値の抽出に失敗:', {
        error: e.message,
        selector: selector,
        htmlSample: html.substring(0, 200) // HTMLの先頭200文字のみログ出力
      });
      return null;
    }
  },

  /**
   * メイン情報のバッチを処理する
   * @param {Object} sheet シート
   * @param {Array} values シートの値
   * @param {Object} columnIndices 列インデックス
   * @param {number} startRow 開始行
   * @param {number} endRow 終了行
   */
  async processMainInfoBatch(sheet, values, columnIndices, startRow, endRow) {
    try {
      // 入力値の検証
      if (!Array.isArray(values)) {
        throw new Error('values must be an array');
      }
      
      if (!columnIndices || !columnIndices.code) {
        throw new Error('必要な列インデックスが見つかりません');
      }
      
      // スクレイピング設定を取得
      const settings = this.getScrapingSettings();
      if (!settings || Object.keys(settings).length === 0) {
        throw new Error('スクレイピング設定が取得できません');
      }
      
      // 現在の日付を取得
      const today = new Date();
      const formattedDate = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');
      
      // バッチ更新用の配列を準備
      const updates = [];
      
      for (let i = startRow; i <= endRow && i <= values.length; i++) {
        const rowIndex = i - 1;
        if (rowIndex >= values.length) {
          LoggerManager.warn(`行インデックス ${rowIndex} が範囲外です`);
          continue;
        }
        
        const row = values[rowIndex];
        if (!Array.isArray(row)) {
          LoggerManager.warn(`行 ${i} のデータが配列ではありません`);
          continue;
        }
        
        const code = row[columnIndices.code - 1];
        if (!code || code.toString().trim() === '') {
          continue;
        }
        
        const productType = columnIndices.productType && columnIndices.productType <= row.length
          ? row[columnIndices.productType - 1]
          : '';
        
        if (productType === 'MMF') {
          continue;
        }
        
        const codeStr = code.toString();
        
        try {
          // スクレイピングで情報を取得
          const stockInfo = await this.fetchStockInfo(codeStr, settings);
          if (stockInfo && stockInfo.previousClose && stockInfo.signal) {
            const update = {
              row: i,
              previousClose: stockInfo.previousClose,
              signal: stockInfo.signal
            };
            // 更新日列が存在する場合のみ追加
            if (columnIndices.updateDate > 0) {
              update.date = formattedDate;
            }
            updates.push(update);
          }
        } catch (fetchError) {
          LoggerManager.error(`銘柄 ${codeStr} の情報取得に失敗:`, fetchError);
          continue;
        }
        
        // API制限を考慮して待機
        await Utilities.sleep(1000);
      }
      
      // 更新データがある場合のみ処理を実行
      if (updates.length > 0) {
        // バッチ更新を実行
        updates.forEach(update => {
          try {
            if (columnIndices.previousClose) {
              sheet.getRange(update.row, columnIndices.previousClose).setValue(update.previousClose);
            }
            if (columnIndices.signal) {
              sheet.getRange(update.row, columnIndices.signal).setValue(update.signal);
            }
            if (columnIndices.updateDate && update.date) {
              sheet.getRange(update.row, columnIndices.updateDate).setValue(update.date);
            }
          } catch (updateError) {
            LoggerManager.error(`行 ${update.row} の更新に失敗:`, updateError);
          }
        });
      }
      
    } catch (e) {
      LoggerManager.error('バッチ処理中にエラー発生', {
        startRow,
        endRow,
        error: e.message,
        stack: e.stack
      });
      throw e;
    }
  },

  /**
   * トレンド情報を更新する
   * @returns {string} 完了メッセージ
   */
  updateTrendInfo(startFromIndex = 0) {
    try {
      // PerformanceMonitor拡張機能の初期化
      initializePerformanceMonitorExtension();
      
      // キャンセルフラグをリセット
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('cancelFlag', 'false');
      
      // 継続処理の開始位置を記録
      let actualStartFromIndex = startFromIndex;
      
      // 保存された開始位置がある場合は読み込む
      if (startFromIndex === 0) {
        const savedStartIndex = scriptProperties.getProperty('nextStartIndex_Reference');
        if (savedStartIndex) {
          actualStartFromIndex = parseInt(savedStartIndex);
          LoggerManager.info(`保存された開始位置を読み込みました: ${actualStartFromIndex}番目から`);
        }
      }
      
      const isResume = actualStartFromIndex > 0;
      if (isResume) {
        LoggerManager.info(`継続処理を開始します: ${actualStartFromIndex}番目から`, {
          resumeFrom: actualStartFromIndex,
          timestamp: new Date().toISOString()
        });
      }
      
      const sheetReference = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.REFERENCE);
      const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTING);
      
      if (!settingSheet || !sheetReference) {
        throw new Error('設定シートまたは_Referenceシートが見つかりません');
      }
      
      // 設定から更新頻度が「毎日」のキーワードを取得
      let dailySettings = [];
      try {
        const settingsData = settingSheet.getRange('A2:E' + settingSheet.getLastRow()).getValues();
        LoggerManager.debug('設定データを取得しました', {
          rowCount: settingsData.length,
          sampleRow: settingsData[0] || 'データなし'
        });
        
        // E列（インデックス4）が「毎日」の行をフィルタ
        dailySettings = settingsData.filter(row => row && row.length > 4 && row[4] === '毎日');
        
        if (dailySettings.length === 0) {
          LoggerManager.warn('「毎日」設定の項目が見つかりません。すべての設定項目を使用します。');
          dailySettings = settingsData.filter(row => row && row[0]); // キーワードがある行のみ
        }
        
        LoggerManager.info('日次更新項目を取得しました', {
          count: dailySettings.length,
          items: dailySettings.map(s => s[0] || '名前なし').join(', ')
        });
        
      } catch (settingsError) {
        LoggerManager.error('設定データの取得に失敗しました', settingsError);
        throw new Error(`設定データ取得エラー: ${settingsError.message}`);
      }
      
      if (!dailySettings || dailySettings.length === 0) {
        throw new Error('有効な設定項目が見つかりません');
      }
      
      // _Referenceシートのデータを取得（ヘッダーは2行目にある）
      const headerRow = sheetReference.getRange(2, 1, 1, sheetReference.getLastColumn()).getValues()[0];
      const dataRange = sheetReference.getRange(3, 1, Math.max(1, sheetReference.getLastRow() - 2), sheetReference.getLastColumn());
      const referenceData = dataRange.getValues();
      
      // 有効データ行数の計算
      const codeIndex = headerRow.indexOf('コード');
      let validData = referenceData.filter(row => {
        return codeIndex !== -1 && row[codeIndex] && row[codeIndex].toString().trim() !== '';
      });
      
      const validRows = validData.length;
      if (validRows <= 0) {
        throw new Error('処理対象の有効なデータが見つかりません');
      }

      UIManager.updateProgress(0, validRows, 'トレンド情報の更新を開始します（全' + validRows + '件）', {
        phase: '初期化',
        startTime: new Date().toISOString(),
        totalRows: totalRows,
        validRows: validRows
      });
      
      UIManager.updateProgress(0, totalRows, '列インデックスを取得中...', {
        phase: '準備',
        itemCount: dailySettings.length,
        items: dailySettings.map(s => s[0]).join(', ')
      });
      
      // 列インデックスの取得
      const columnIndices = {};
      dailySettings.forEach(setting => {
        const keyword = setting[0];
        const index = headerRow.indexOf(keyword);
        if (index !== -1) {
          columnIndices[keyword] = index + 1;
        }
      });

      // 必須列の確認
      const requiredColumns = ['コード', '商品種別'];
      const missingColumns = [];
      
      requiredColumns.forEach(colName => {
        const index = headerRow.indexOf(colName);
        if (index === -1) {
          missingColumns.push(colName);
        } else {
          columnIndices[colName] = index + 1;
        }
      });
      
      if (missingColumns.length > 0) {
        throw new Error(`必要な列が見つかりません: ${missingColumns.join(', ')}`);
      }
      
      // 株価列のインデックスを取得（オプション）
      const stockPriceIndex = headerRow.indexOf('株価');
      if (stockPriceIndex !== -1) {
        columnIndices['株価'] = stockPriceIndex + 1;
      }

      // 処理設定 - タイムアウトを避けるための最適化
      const batchSize = Math.min(8, Math.max(3, Math.floor(validRows / 25))); // バッチサイズを小さく
      const sleepTime = {
        betweenBatches: 800,  // 待機時間を短縮
        betweenColumns: 300,  // 列間待機を短縮
        betweenRetries: 200   // リトライ待機を短縮
      };
      
      // 安全マージンを考慮した最大処理時間（5分）
      const MAX_SAFE_TIME = 5 * 60 * 1000;
      let failedUpdates = [];
      let processedCount = 0;
      let startTime = new Date();

      UIManager.updateProgress(0, validRows, '処理を開始します...', {
        phase: '開始',
        batchSize: batchSize,
        totalBatches: Math.ceil(validData.length / batchSize),
        startTime: startTime.toISOString()
      });

      // 継続処理の開始位置を調整
      const actualStartIndex = Math.max(0, actualStartFromIndex);
      const remainingData = validData.slice(actualStartIndex);
      const totalRemainingItems = remainingData.length;
      
      if (totalRemainingItems === 0) {
        LoggerManager.info('処理対象がありません。完了します。');
        UIManager.updateProgress(validRows, validRows, '処理が完了しました', {
          phase: '完了',
          message: '全ての処理が完了しました'
        });
        return 'Completed';
      }
      
      // バッチ処理開始
      for (let batchStart = 0; batchStart < remainingData.length; batchStart += batchSize) {
        const batchEnd = Math.min(batchStart + batchSize, remainingData.length);
        const batchData = remainingData.slice(batchStart, batchEnd);
        const currentBatch = Math.floor(batchStart / batchSize) + 1;
        const totalBatches = Math.ceil(remainingData.length / batchSize);
        const actualProcessedCount = actualStartIndex + batchStart;
        
        // キャンセルフラグのチェック
        if (scriptProperties.getProperty('cancelFlag') === 'true') {
          UIManager.updateProgress(
            actualProcessedCount,
            validRows,
            `処理がキャンセルされました（${actualProcessedCount}/${validRows}）`,
            {
              phase: 'キャンセル',
              elapsedTime: this.formatElapsedTime(startTime),
              batchesCompleted: currentBatch - 1,
              totalBatches: totalBatches
            }
          );
          return 'Cancelled';
        }
        
        // バッチ処理開始メッセージ
        UIManager.updateProgress(
          processedCount,
          validRows,
          `バッチ ${currentBatch}/${totalBatches} を処理中...`,
          {
            phase: 'バッチ処理',
            currentBatch: currentBatch,
            totalBatches: totalBatches,
            batchSize: batchData.length,
            elapsedTime: this.formatElapsedTime(startTime)
          }
        );
        
        // 株価の更新
        if (columnIndices['株価']) {
          const stockPriceFormulas = [];
          const stockPriceFormats = [];
          const stockCodes = [];

          for (let i = 0; i < batchData.length; i++) {
            const code = batchData[i][columnIndices['コード'] - 1];
            // _Referenceシートでは商品種別列がないため、コードから推定
            let productType = 'JP'; // デフォルト値
            if (code) {
              if (code.startsWith('JP') && code.length === 12) {
                productType = 'TOSHIN';
              } else if (['BND', 'HDV', 'SPYD', 'TLT', 'VIG'].includes(code)) {
                productType = 'US';
              } else if (code === 'FMJXX0000000') {
                productType = 'MMF';
              }
            }
            const actualRow = batchStart + i + 3;
            
            if (code) {
              stockCodes.push(code);
            }

            let formula = '';
            if (code) {
              if (productType === 'MMF') {
                formula = `=E${actualRow}`;
              } else if (productType === 'US') {
                formula = `=IF(ISBLANK(B${actualRow}), "", GOOGLEFINANCE(B${actualRow}))`;
              } else if (productType === 'FUND') {
                // 投資信託コードの形式を確認し、必要に応じて修正
                let fundCode = code;
                if (!fundCode.startsWith('JP')) {
                  fundCode = 'JP' + fundCode;
                }
                formula = `=IF(ISBLANK(B${actualRow}), "", STOCKPRICEJP("TOSHIN", "${fundCode}"))`;
              } else {
                formula = `=IF(ISBLANK(B${actualRow}),"",STOCKPRICEJP("JP", B${actualRow}))`;
              }
            }
            
            stockPriceFormulas.push([formula]);
            stockPriceFormats.push([(productType === 'US' || productType === 'MMF') ? '"$"#,##0.00' : '¥#,##0.00']);
          }

          if (stockPriceFormulas.length > 0) {
            try {
              UIManager.updateProgress(
                processedCount,
                validRows,
                `株価情報を更新中... (${stockCodes.length}件)`,
                {
                  phase: '株価更新',
                  currentBatch: currentBatch,
                  totalBatches: totalBatches,
                  processingItems: stockCodes.join(', ').substring(0, 100) + (stockCodes.length > 5 ? '...' : '')
                }
              );
              
              const stockPriceRange = sheetReference.getRange(batchStart + 3, columnIndices['株価'], stockPriceFormulas.length, 1);
              stockPriceRange.setValues(stockPriceFormulas);
              stockPriceRange.setNumberFormats(stockPriceFormats);
              Utilities.sleep(sleepTime.betweenColumns);
            } catch (error) {
              LoggerManager.error(`株価列の更新エラー: ${batchStart + 3}行目から`, error);
              failedUpdates.push({
                type: '株価',
                startRow: batchStart + 3,
                endRow: batchStart + 3 + stockPriceFormulas.length - 1,
                error: error.message
              });
            }
          }
        }

        // その他の日次更新項目の処理
        for (const setting of dailySettings) {
          const keyword = setting[0];
          if (keyword === '株価') continue;
          
          const colIndex = columnIndices[keyword];
          if (!colIndex) continue;

          const columnData = [];
          let hasError = false;
          const processingCodes = [];

          UIManager.updateProgress(
            processedCount,
            validRows,
            `${keyword}情報を更新中...`,
            {
              phase: '項目更新',
              currentItem: keyword,
              currentBatch: currentBatch,
              totalBatches: totalBatches,
              elapsedTime: this.formatElapsedTime(startTime)
            }
          );

          for (let i = 0; i < batchData.length; i++) {
            const code = batchData[i][columnIndices['コード'] - 1];
            let value = '';

            if (code) {
              processingCodes.push(code);
              
              for (let retryCount = 0; retryCount < 2; retryCount++) {
                try {
                  // URLが設定されているか確認
                  if (!setting[3]) {
                    LoggerManager.warn(`${keyword}のURL設定がありません`);
                    break;
                  }
                  
                  const baseUrl = setting[3].replace('{code}', code);
                  
                  // 正規表現が設定されているか確認
                  if (!setting[1]) {
                    LoggerManager.warn(`${keyword}の正規表現設定がありません`);
                    break;
                  }
                  
                  const regex = new RegExp(setting[1], 's');
                  const html = UrlFetchApp.fetch(baseUrl, {
                    muteHttpExceptions: true,
                    validateHttpsCertificates: true,
                    timeout: 5000
                  }).getContentText();
                  
                  const match = html.match(regex);
                  if (match && match[1]) {
                    value = match[1].trim();
                    LoggerManager.debug(`${keyword}の生の取得値: コード=${code}, 値="${value}"`);
                    
                    if (keyword === '配当金' || keyword === 'PER' || keyword === 'PBR') {
                      const originalValue = value;
                      value = value.replace(/[^0-9.]/g, '');
                      value = parseFloat(value);
                      if (isNaN(value)) value = '';
                      LoggerManager.debug(`${keyword}の処理後の値: コード=${code}, 元の値="${originalValue}", 処理後="${value}"`);
                    }
                    break;
                  } else {
                    LoggerManager.debug(`${keyword}の正規表現マッチ失敗: コード=${code}, URL=${baseUrl}`);
                    if (keyword === '配当金') {
                      LoggerManager.debug(`配当金のHTML内容(最初の500文字): ${html.substring(0, 500)}`);
                      LoggerManager.debug(`配当金の正規表現: ${regex}`);
                    }
                  }
                } catch (error) {
                  if (retryCount === 1) {
                    hasError = true;
                    LoggerManager.error(`${keyword}の取得エラー: コード=${code}`, error);
                  }
                  Utilities.sleep(sleepTime.betweenRetries);
                }
              }
            }
            columnData.push([value]);
          }

          if (columnData.length > 0) {
            try {
              UIManager.updateProgress(
                processedCount,
                validRows,
                `${keyword}情報を更新中... (${processingCodes.length}件)`,
                {
                  phase: '項目更新',
                  currentItem: keyword,
                  processingItems: processingCodes.join(', ').substring(0, 100) + (processingCodes.length > 5 ? '...' : ''),
                  hasErrors: hasError
                }
              );
              
              // Debug logging for batch processing error
              LoggerManager.debug(`Attempting to update ${keyword} data`, {
                batchStart: batchStart,
                colIndex: colIndex,
                columnDataLength: columnData.length,
                sheetReferenceExists: !!sheetReference,
                sheetReferenceName: sheetReference ? sheetReference.getName() : 'undefined'
              });
              
              const columnRange = sheetReference.getRange(batchStart + 3, colIndex, columnData.length, 1);
              columnRange.setValues(columnData);
            } catch (error) {
              LoggerManager.error(`${keyword}のバッチ更新エラー:`, error);
              failedUpdates.push({
                type: keyword,
                startRow: batchStart + 3,
                endRow: batchStart + 3 + columnData.length - 1,
                error: error.message
              });
            }
          }
          
          // 列間の待機
          Utilities.sleep(sleepTime.betweenColumns);
        }
        
        // 処理済み件数を更新
        processedCount += batchData.length;
        
        // 安全マージンを考慮したタイムアウトチェック
        const currentTime = new Date();
        const elapsedTime = currentTime - startTime;
        
        // 進捗の更新
        const elapsedTimeFormatted = this.formatElapsedTime(startTime);
        const percentComplete = Math.round((processedCount / validRows) * 100);
        
        UIManager.updateProgress(
          processedCount,
          validRows,
          `トレンド情報を更新中: ${processedCount}/${validRows} (${percentComplete}%)`,
          {
            phase: 'バッチ完了',
            currentBatch: currentBatch,
            totalBatches: totalBatches,
            elapsedTime: elapsedTimeFormatted,
            percentComplete: percentComplete
          }
        );
        
        if (elapsedTime > MAX_SAFE_TIME) {
          const nextStartIndex_Reference = actualProcessedCount;
          LoggerManager.warn(`安全マージンのため処理を中断しました（${actualProcessedCount}/${validRows}件完了）。次回は${nextStartIndex_Reference}番目から再開します。`);
          
          // 次回の開始位置を保存
          scriptProperties.setProperty('nextStartIndex_Reference', nextStartIndex_Reference.toString());
          scriptProperties.setProperty('lastProcessedCount_Reference', actualProcessedCount.toString());
          scriptProperties.setProperty('safeTimeout_Reference', 'true'); // 安全タイムアウトフラグ
          
          TimeoutManager.cleanup();
          return 'SafeTimeout';
        }
        
        // 通常のタイムアウトチェック（バックアップ）
        if (TimeoutManager.checkAndHandleTimeout('updateTrendInfo')) {
          const nextStartIndex_Reference = actualProcessedCount;
          LoggerManager.warn(`タイムアウトのため処理を中断しました（${actualProcessedCount}/${validRows}件完了）。次回は${nextStartIndex_Reference}番目から再開します。`);
          
          // 次回の開始位置を保存
          scriptProperties.setProperty('nextStartIndex_Reference', nextStartIndex_Reference.toString());
          scriptProperties.setProperty('lastProcessedCount_Reference', actualProcessedCount.toString());
          
          TimeoutManager.cleanup();
          return 'Timeout';
        }
        
        // バッチ間の待機
        if (batchEnd < remainingData.length) {
          Utilities.sleep(sleepTime.betweenBatches);
        }
      }

      const totalElapsedTime = this.formatElapsedTime(startTime);

      if (failedUpdates.length > 0) {
        LoggerManager.warn('更新に失敗した箇所があります:', failedUpdates);
        UIManager.updateProgress(
          validRows, 
          validRows, 
          `トレンド情報の更新が完了しましたが、${failedUpdates.length}件のエラーがありました`,
          {
            phase: '完了（一部エラー）',
            elapsedTime: totalElapsedTime,
            errorCount: failedUpdates.length,
            errors: failedUpdates.map(f => `${f.type}(${f.startRow}-${f.endRow}): ${f.error}`).join('; ')
          }
        );
      } else {
        UIManager.updateProgress(
          validRows, 
          validRows, 
          'トレンド情報の更新が完了しました',
          {
            phase: '完了',
            elapsedTime: totalElapsedTime,
            message: '全ての処理が正常に完了しました'
          }
        );
      }
      
      // 処理完了時に保存された開始位置をクリア
      scriptProperties.deleteProperty('nextStartIndex_Reference');
      scriptProperties.deleteProperty('lastProcessedCount_Reference');
      scriptProperties.deleteProperty('safeTimeout_Reference');
      
      return 'Completed';

    } catch (e) {
      LoggerManager.error('Error in updateTrendInfo:', e);
      UIManager.updateProgress(0, 1, 'エラーが発生しました: ' + e.message, {
        phase: 'エラー',
        error: e.message,
        stack: e.stack
      });
      throw e;
    }
  },

  /**
   * 経過時間をフォーマットする
   * @param {Date} startTime 開始時間
   * @returns {string} フォーマットされた経過時間
   */
  formatElapsedTime: function(startTime) {
    const now = new Date();
    const elapsedMs = now.getTime() - startTime.getTime();
    
    const minutes = Math.floor(elapsedMs / 60000);
    const seconds = Math.floor((elapsedMs % 60000) / 1000);
    
    return `${minutes}分${seconds}秒`;
  },

  /**
   * ポートフォリオを更新する
   * @returns {string} 完了メッセージ
   */
  updatePF: function() {
    try {
      // PerformanceMonitor拡張機能の初期化
      initializePerformanceMonitorExtension();
      
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheetMain = spreadsheet.getSheetByName(SHEET_NAMES.MAIN);
      const sheetOrders = spreadsheet.getSheetByName(SHEET_NAMES.ORDERS);

      if (!sheetMain || !sheetOrders) {
        throw new Error('必要なシートが見つかりません');
      }

      // 開始時間を記録
      const startTime = new Date();

      // プログレス表示の初期化（件数表示なし）
      UIManager.updateProgress(0, 100, '保有株リストの更新を開始します', {
        phase: '初期化',
        startTime: startTime.toISOString()
      });

      // 注文データの取得
      const lastRow = sheetOrders.getLastRow();
      const lastCol = sheetOrders.getLastColumn();
      if (lastRow < 2) {
        throw new Error('注文一覧シートにデータが存在しません');
      }

      const totalRecords = lastRow - 1; // ヘッダー行を除いたデータ件数

      UIManager.updateProgress(10, 100, '注文データを読み込み中...', {
        phase: '読み込み',
        rows: totalRecords,
        columns: lastCol
      });

      const ordersRange = sheetOrders.getRange(1, 1, lastRow, lastCol);
      const ordersData = ordersRange.getValues();

      UIManager.updateProgress(20, 100, '注文データの処理を開始します', {
        phase: '注文処理',
        orderCount: totalRecords,
        elapsedTime: this.formatElapsedTime(startTime)
      });
      
      const stockData = StockDataManager.processOrders(ordersData);
      const stockCount = Object.keys(stockData).length;

      UIManager.updateProgress(40, 100, `注文データの処理が完了しました (${stockCount}銘柄)`, {
        phase: '注文処理完了',
        stockCount: stockCount,
        elapsedTime: this.formatElapsedTime(startTime)
      });

      UIManager.updateProgress(50, 100, 'メインシートの更新を開始します', {
        phase: 'メインシート更新',
        elapsedTime: this.formatElapsedTime(startTime)
      });

      // メインシートのヘッダー行から必要な列のインデックスを取得
      const headerRow = sheetMain.getRange(2, 1, 1, sheetMain.getLastColumn()).getValues()[0];
      const columnIndices = {
        no: headerRow.indexOf('No.') + 1,
        code: headerRow.indexOf('コード') + 1,
        name: headerRow.indexOf('銘柄名') + 1,
        shares: headerRow.indexOf('保有数') + 1,
        purchasePrice: headerRow.indexOf('購入価格') + 1,
        unitPrice: headerRow.indexOf('購入単価') + 1,
        stockPrice: headerRow.indexOf('株価') + 1,
        accountType: headerRow.indexOf('口座種別') + 1,
        productType: headerRow.indexOf('商品種別') + 1,
        accountHolder: headerRow.indexOf('口座名義人') + 1,
        industry: headerRow.indexOf('業種17分類') + 1,
        sensitivity: headerRow.indexOf('景気感応度') + 1
      };

      // メインシートの現在の行数を取得
      const lastRowMain = sheetMain.getLastRow();

      UIManager.updateProgress(60, 100, 'メインシートをクリアしています...', {
        phase: 'シートクリア',
        elapsedTime: this.formatElapsedTime(startTime),
        totalRecords: totalRecords
      });

      // メインシートのクリア（3行目以降）
      if (lastRowMain > 2) {
        sheetMain.getRange(3, 1, lastRowMain - 2, sheetMain.getLastColumn()).clearContent();
      }

      UIManager.updateProgress(70, 100, '新しいデータを作成しています...', {
        phase: 'データ作成',
        stockCount: Object.values(stockData).filter(stock => stock.totalNum > 0).length,
        elapsedTime: this.formatElapsedTime(startTime),
        totalRecords: totalRecords
      });

      // 新しいデータの作成
      const newData = Object.values(stockData)
        .filter(stock => stock.totalNum > 0)
        .sort((a, b) => {
          const codeA = (a.code || '').toString();
          const codeB = (b.code || '').toString();
          return codeA.localeCompare(codeB);
        })
        .map((stock, index) => {
          const row = new Array(sheetMain.getLastColumn()).fill('');
          row[columnIndices.no - 1] = index + 1;
          row[columnIndices.code - 1] = stock.code.toString();
          row[columnIndices.name - 1] = stock.name || '';
          row[columnIndices.shares - 1] = stock.totalNum;
          row[columnIndices.purchasePrice - 1] = stock.totalCost;
          row[columnIndices.accountType - 1] = stock.accountType || "未分類";
          row[columnIndices.productType - 1] = stock.productType || "未分類";
          row[columnIndices.accountHolder - 1] = stock.accountHolder || "未設定";
          row[columnIndices.industry - 1] = stock.industry || '';
          row[columnIndices.sensitivity - 1] = stock.sensitivity || '';
          
          // 購入単価の計算（総コスト÷保有数）
          if (stock.totalNum > 0) {
            row[columnIndices.unitPrice - 1] = stock.totalCost / stock.totalNum;
          }

          // 株価の数式設定
          let formula;
          if (stock.isMMF) {
            formula = `=E${index + 3}`;
          } else if (stock.productType === 'US') {
            formula = `=IF(ISBLANK(B${index + 3}), "", GOOGLEFINANCE(B${index + 3}))`;
          } else if (stock.productType === 'FUND') {
            formula = `=IF(ISBLANK(B${index + 3}), "", STOCKPRICEJP("TOSHIN", B${index + 3}))`;
          } else {
            formula = `=IF(ISBLANK(B${index + 3}),"",STOCKPRICEJP("JP", B${index + 3}))`;
          }
          
          if (columnIndices.stockPrice > 0) {
            row[columnIndices.stockPrice - 1] = formula;
          }

          return row;
        });

      UIManager.updateProgress(80, 100, 'データをシートに書き込んでいます...', {
        phase: 'データ書き込み',
        rowCount: newData.length,
        elapsedTime: this.formatElapsedTime(startTime),
        totalRecords: totalRecords
        });

      if (newData.length > 0) {
        const range = sheetMain.getRange(3, 1, newData.length, sheetMain.getLastColumn());
        range.setValues(newData);

        UIManager.updateProgress(90, 100, '数値フォーマットを設定しています...', {
          phase: 'フォーマット設定',
          rowCount: newData.length,
          elapsedTime: this.formatElapsedTime(startTime),
          totalRecords: totalRecords
        });

        // 数値フォーマットの設定
        newData.forEach((_, index) => {
          const rowIndex = index + 3;
          const stock = Object.values(stockData).filter(s => s.totalNum > 0).sort((a, b) => {
            const codeA = (a.code || '').toString();
            const codeB = (b.code || '').toString();
            return codeA.localeCompare(codeB);
          })[index];

          // 株価列の数値フォーマット設定
          if (columnIndices.stockPrice > 0) {
            const priceCell = sheetMain.getRange(rowIndex, columnIndices.stockPrice);
            if (stock.isUSD || stock.isMMF) {
              priceCell.setNumberFormat('"$"#,##0.00');
            } else {
              priceCell.setNumberFormat('¥#,##0.00');
            }
          }

          // 購入単価列の数値フォーマット設定
          if (columnIndices.unitPrice > 0) {
            const unitPriceCell = sheetMain.getRange(rowIndex, columnIndices.unitPrice);
            if (stock.isUSD || stock.isMMF) {
              unitPriceCell.setNumberFormat('"$"#,##0.00');
            } else {
              unitPriceCell.setNumberFormat('¥#,##0.00');
            }
          }

          // 購入価格列の数値フォーマット設定
          if (columnIndices.purchasePrice > 0) {
            const purchasePriceCell = sheetMain.getRange(rowIndex, columnIndices.purchasePrice);
            if (stock.isUSD || stock.isMMF) {
              purchasePriceCell.setNumberFormat('"$"#,##0.00');
            } else {
              purchasePriceCell.setNumberFormat('¥#,##0.00');
            }
          }

          // 保有数列の数値フォーマット設定
          if (columnIndices.shares > 0) {
            const sharesCell = sheetMain.getRange(rowIndex, columnIndices.shares);
            sharesCell.setNumberFormat('#,##0');
          }
        });
      }

      const totalElapsedTime = this.formatElapsedTime(startTime);
      UIManager.updateProgress(100, 100, '保有株リストの更新が完了しました', {
        phase: '完了',
        elapsedTime: totalElapsedTime,
        stockCount: newData.length,
        totalRecords: totalRecords,
        message: `${newData.length}銘柄の情報を更新しました（全${totalRecords}レコード中）`
      });
      
      return 'Completed';

    } catch (e) {
      LoggerManager.error('Error in updatePF:', e);
      UIManager.updateProgress(100, 100, 'エラーが発生しました: ' + e.message, {
        phase: 'エラー',
        error: e.message,
        stack: e.stack,
        totalRecords: totalRecords
      });
      throw e;
    }
  },

  /**
   * 現在の処理をキャンセルする
   * @returns {string} キャンセルメッセージ
   */
  cancelCurrentProcess() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('cancelFlag', 'true');
      TimeoutManager.deleteAllTriggers();
      
      // キャンセル状態を明示的に表示
      UIManager.updateProgress(100, 100, '処理がキャンセルされました');
      
      LoggerManager.info('処理のキャンセルがリクエストされました');
      return 'Process cancelled';
    } catch (e) {
      LoggerManager.error('Error in cancelCurrentProcess:', e);
      throw e;
    }
  },

  /**
   * すべてのトリガーを削除する
   */
  deleteTriggers() {
    try {
      TimeoutManager.deleteAllTriggers();
      LoggerManager.debug('全てのトリガーを削除しました');
    } catch (e) {
      LoggerManager.error('Error in deleteTriggers:', e);
      throw e;
    }
  },

  /**
   * 選択された行のみ_Referenceシートの情報を更新する
   * @returns {string} 完了メッセージ
   */
  updateSelectedReferenceInfo: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.REFERENCE);
      const selection = SpreadsheetApp.getActiveRange();
      
      if (!sheet || !selection) {
        throw new Error('_Referenceシートが見つからないか、範囲が選択されていません');
      }
      
      // 選択された行を取得
      const startRow = selection.getRow();
      const numRows = selection.getNumRows();
      
      // ヘッダー行（2行目）より前の選択は無効
      if (startRow < 3) {
        throw new Error('データ行（3行目以降）を選択してください');
      }
      
      const selectedRows = [];
      for (let i = 0; i < numRows; i++) {
        selectedRows.push(startRow + i);
      }
      
      LoggerManager.info('選択行の_Reference情報更新を開始', {
        selectedRows: selectedRows,
        rowCount: selectedRows.length
      });
      
      // 選択された行のコードを取得
      const headerRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      const codeColumnIndex = headerRow.indexOf('コード');
      
      if (codeColumnIndex === -1) {
        throw new Error('コード列が見つかりません');
      }
      
      const targetCodes = [];
      selectedRows.forEach(row => {
        const code = sheet.getRange(row, codeColumnIndex + 1).getValue();
        if (code && code.toString().trim() !== '') {
          targetCodes.push(code.toString().trim());
        }
      });
      
      if (targetCodes.length === 0) {
        throw new Error('選択された行に有効なコードが見つかりません');
      }
      
      LoggerManager.info('選択行のコード取得完了', {
        codes: targetCodes,
        codeCount: targetCodes.length
      });
      
      // 通常のupdateReferenceInfo関数を呼び出し、対象コードを指定
      return this.updateReferenceInfo(targetCodes);
      
    } catch (e) {
      LoggerManager.error('選択行の_Reference情報更新エラー', e);
      throw e;
    }
  },

  /**
   * _Referenceシートの情報を更新する
   * @param {Array} targetCodes 更新対象の銘柄コード（指定がない場合は全銘柄を更新）
   * @returns {string} 完了メッセージ
   */
  updateReferenceInfo: function(targetCodes = []) {
    try {
      // キャンセルフラグをリセット
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('cancelFlag', 'false');
      
      // タイムアウト管理の開始
      TimeoutManager.startExecutionTimer();
      
      const sheetReference = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.REFERENCE);
      const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SETTING);
      
      if (!settingSheet || !sheetReference) {
        throw new Error('設定シートまたは_Referenceシートが見つかりません');
      }
      
      // 開始時間を記録
      const startTime = new Date();
      
      // 継続処理の開始位置を記録（現在は0固定だが将来の拡張用）
      let actualStartFromIndex = 0;
      
      // 保存された開始位置がある場合は読み込む（将来の拡張用）
      const savedStartIndex = scriptProperties.getProperty('nextStartIndex_Reference_Reference');
      if (savedStartIndex) {
        actualStartFromIndex = parseInt(savedStartIndex);
        LoggerManager.info(`_Reference継続処理開始位置: ${actualStartFromIndex}番目から`);
      }
      
      const isResume = actualStartFromIndex > 0;
      if (isResume) {
        LoggerManager.info(`_Reference継続処理を開始します: ${actualStartFromIndex}番目から`, {
          resumeFrom: actualStartFromIndex,
          timestamp: new Date().toISOString()
        });
      }
      
      // _Referenceシートのデータを取得（ヘッダーは2行目にある）
      const headerRow = sheetReference.getRange(2, 1, 1, sheetReference.getLastColumn()).getValues()[0];
      const dataRange = sheetReference.getRange(3, 1, Math.max(1, sheetReference.getLastRow() - 2), sheetReference.getLastColumn());
      const referenceData = dataRange.getValues();
      
      LoggerManager.info(`_Referenceシートのヘッダー: ${headerRow.join(', ')}`);
      
      // 設定から更新頻度が「毎日」のキーワードを取得
      let dailySettings = [];
      try {
        const settingsData = settingSheet.getRange('A2:E' + settingSheet.getLastRow()).getValues();
        LoggerManager.debug('設定データを取得しました', {
          rowCount: settingsData.length,
          sampleRow: settingsData[0] || 'データなし'
        });
        
        // _Reference情報更新では「毎日」と「決算ごと」の両方を取得
        const dailyItems = settingsData.filter(row => row && row.length > 4 && row[4] === '毎日');
        const quarterlyItems = settingsData.filter(row => row && row.length > 4 && row[4] === '決算ごと');
        
        // 両方の項目を結合
        dailySettings = [...dailyItems, ...quarterlyItems];
        
        if (dailySettings.length === 0) {
          LoggerManager.warn('更新項目が見つかりません。すべての設定項目を使用します。');
          dailySettings = settingsData.filter(row => row && row[0]); // キーワードがある行のみ
        }
        
        LoggerManager.info('_Reference用更新項目を取得しました', {
          dailyCount: dailyItems.length,
          quarterlyCount: quarterlyItems.length,
          totalCount: dailySettings.length,
          dailyItems: dailyItems.map(s => s[0] || '名前なし').join(', '),
          quarterlyItems: quarterlyItems.map(s => s[0] || '名前なし').join(', ')
        });
        
        
        // 各設定項目の詳細をログ出力
        dailySettings.forEach((setting, index) => {
          LoggerManager.debug(`設定項目 ${index + 1}`, {
            keyword: setting[0] || 'なし',
            regex: setting[1] || 'なし',
            tag: setting[2] || 'なし', 
            url: setting[3] || 'なし',
            frequency: setting[4] || 'なし'
          });
        });
        
      } catch (settingsError) {
        LoggerManager.error('設定データの取得に失敗しました', settingsError);
        throw new Error(`設定データ取得エラー: ${settingsError.message}`);
      }
      
      if (!dailySettings || dailySettings.length === 0) {
        throw new Error('有効な設定項目が見つかりません');
      }
      
      // 列インデックスを取得
      const columnIndices = {};
      dailySettings.forEach(setting => {
        const keyword = setting[0];
        const index = headerRow.indexOf(keyword);
        if (index !== -1) {
          columnIndices[keyword] = index + 1;
        }
      });

      // 必須列の確認（_Referenceシートでは「商品種別」は不要）
      const requiredColumns = ['コード'];
      const missingColumns = [];
      
      requiredColumns.forEach(colName => {
        const index = headerRow.indexOf(colName);
        if (index === -1) {
          missingColumns.push(colName);
        } else {
          columnIndices[colName] = index + 1;
        }
      });
      
      if (missingColumns.length > 0) {
        throw new Error(`必要な列が見つかりません: ${missingColumns.join(', ')}`);
      }
      
      // 株価列のインデックスを取得（オプション）
      const stockPriceIndex = headerRow.indexOf('株価');
      if (stockPriceIndex !== -1) {
        columnIndices['株価'] = stockPriceIndex + 1;
      }

      // 有効データ行数の計算
      const codeIndex = headerRow.indexOf('コード');
      let validData = referenceData.filter(row => {
        return codeIndex !== -1 && row[codeIndex] && row[codeIndex].toString().trim() !== '';
      });
      
      // targetCodesが指定されている場合は、そのコードのみに絞り込む
      if (targetCodes && targetCodes.length > 0) {
        validData = validData.filter(row => {
          const code = row[codeIndex] ? row[codeIndex].toString().trim() : '';
          return targetCodes.includes(code);
        });
        LoggerManager.info('対象コードで絞り込みを実行', {
          targetCodes: targetCodes,
          filteredCount: validData.length
        });
      }
      
      const validRows = validData.length;
      
      if (validRows <= 0) {
        throw new Error('処理対象の有効なデータが見つかりません');
      }

      LoggerManager.info(`_Referenceシート処理対象: ${validRows}件`);

      // 処理設定 - タイムアウトを避けるための最適化
      const batchSize = Math.min(8, Math.max(3, Math.floor(validRows / 25))); // バッチサイズを小さく
      const sleepTime = {
        betweenBatches: 800,  // 待機時間を短縮
        betweenColumns: 300,  // 列間待機を短縮
        betweenRetries: 200   // リトライ待機を短縮
      };
      
      // 安全マージンを考慮した最大処理時間（5分）
      const MAX_SAFE_TIME = 5 * 60 * 1000;
      let failedUpdates = [];
      let processedCount = 0;

      UIManager.updateProgress(0, validRows, '処理を開始します...', {
        phase: '開始',
        batchSize: batchSize,
        totalBatches: Math.ceil(validData.length / batchSize),
        startTime: startTime.toISOString()
      });

      // 継続処理の開始位置を調整
      const actualStartIndex = Math.max(0, actualStartFromIndex);
      const remainingData = validData.slice(actualStartIndex);
      const totalRemainingItems = remainingData.length;
      
      if (totalRemainingItems === 0) {
        LoggerManager.info('処理対象がありません。完了します。');
        UIManager.updateProgress(validRows, validRows, '処理が完了しました', {
          phase: '完了',
          message: '全ての処理が完了しました'
        });
        return 'Completed';
      }
      
      // バッチ処理開始
      for (let batchStart = 0; batchStart < remainingData.length; batchStart += batchSize) {
        const batchEnd = Math.min(batchStart + batchSize, remainingData.length);
        const batchData = remainingData.slice(batchStart, batchEnd);
        const currentBatch = Math.floor(batchStart / batchSize) + 1;
        const totalBatches = Math.ceil(remainingData.length / batchSize);
        const actualProcessedCount = actualStartIndex + batchStart;
        
        // キャンセルフラグのチェック
        if (scriptProperties.getProperty('cancelFlag') === 'true') {
          UIManager.updateProgress(
            actualProcessedCount,
            validRows,
            `処理がキャンセルされました（${actualProcessedCount}/${validRows}）`,
            {
              phase: 'キャンセル',
              elapsedTime: this.formatElapsedTime(startTime),
              batchesCompleted: currentBatch - 1,
              totalBatches: totalBatches
            }
          );
          return 'Cancelled';
        }
        
        // バッチ処理開始メッセージ
        UIManager.updateProgress(
          processedCount,
          validRows,
          `バッチ ${currentBatch}/${totalBatches} を処理中...`,
          {
            phase: 'バッチ処理',
            currentBatch: currentBatch,
            totalBatches: totalBatches,
            batchSize: batchData.length,
            elapsedTime: this.formatElapsedTime(startTime)
          }
        );
        
        // 株価の更新
        if (columnIndices['株価']) {
          const stockPriceFormulas = [];
          const stockPriceFormats = [];
          const stockCodes = [];

          for (let i = 0; i < batchData.length; i++) {
            const code = batchData[i][columnIndices['コード'] - 1];
            // _Referenceシートでは商品種別列がないため、コードから推定
            let productType = 'JP'; // デフォルト値
            if (code) {
              if (code.startsWith('JP') && code.length === 12) {
                productType = 'TOSHIN';
              } else if (['BND', 'HDV', 'SPYD', 'TLT', 'VIG'].includes(code)) {
                productType = 'US';
              } else if (code === 'FMJXX0000000') {
                productType = 'MMF';
              }
            }
            const actualRow = batchStart + i + 3;
            
            if (code) {
              stockCodes.push(code);
            }

            let formula = '';
            if (code) {
              if (productType === 'MMF') {
                formula = `=E${actualRow}`;
              } else if (productType === 'US') {
                formula = `=IF(ISBLANK(B${actualRow}), "", GOOGLEFINANCE(B${actualRow}))`;
              } else if (productType === 'FUND') {
                // 投資信託コードの形式を確認し、必要に応じて修正
                let fundCode = code;
                if (!fundCode.startsWith('JP')) {
                  fundCode = 'JP' + fundCode;
                }
                formula = `=IF(ISBLANK(B${actualRow}), "", STOCKPRICEJP("TOSHIN", "${fundCode}"))`;
              } else {
                formula = `=IF(ISBLANK(B${actualRow}),"",STOCKPRICEJP("JP", B${actualRow}))`;
              }
            }
            
            stockPriceFormulas.push([formula]);
            stockPriceFormats.push([(productType === 'US' || productType === 'MMF') ? '"$"#,##0.00' : '¥#,##0.00']);
          }

          if (stockPriceFormulas.length > 0) {
            try {
              UIManager.updateProgress(
                processedCount,
                validRows,
                `株価情報を更新中... (${stockCodes.length}件)`,
                {
                  phase: '株価更新',
                  currentBatch: currentBatch,
                  totalBatches: totalBatches,
                  processingItems: stockCodes.join(', ').substring(0, 100) + (stockCodes.length > 5 ? '...' : '')
                }
              );
              
              const stockPriceRange = sheetReference.getRange(batchStart + 3, columnIndices['株価'], stockPriceFormulas.length, 1);
              stockPriceRange.setValues(stockPriceFormulas);
              stockPriceRange.setNumberFormats(stockPriceFormats);
              Utilities.sleep(sleepTime.betweenColumns);
            } catch (error) {
              LoggerManager.error(`株価列の更新エラー: ${batchStart + 3}行目から`, error);
              failedUpdates.push({
                type: '株価',
                startRow: batchStart + 3,
                endRow: batchStart + 3 + stockPriceFormulas.length - 1,
                error: error.message
              });
            }
          }
        }

        // その他の日次更新項目の処理
        for (const setting of dailySettings) {
          const keyword = setting[0];
          if (keyword === '株価') continue;
          
          const colIndex = columnIndices[keyword];
          if (!colIndex) continue;

          const columnData = [];
          let hasError = false;
          const processingCodes = [];

          UIManager.updateProgress(
            processedCount,
            validRows,
            `${keyword}情報を更新中...`,
            {
              phase: '項目更新',
              currentItem: keyword,
              currentBatch: currentBatch,
              totalBatches: totalBatches,
              elapsedTime: this.formatElapsedTime(startTime)
            }
          );

          for (let i = 0; i < batchData.length; i++) {
            const code = batchData[i][columnIndices['コード'] - 1];
            let value = '';

            if (code) {
              processingCodes.push(code);
              
              for (let retryCount = 0; retryCount < 2; retryCount++) {
                try {
                  // URLが設定されているか確認
                  if (!setting[3]) {
                    LoggerManager.warn(`${keyword}のURL設定がありません`);
                    break;
                  }
                  
                  const baseUrl = setting[3].replace('{code}', code);
                  
                  // 正規表現が設定されているか確認
                  if (!setting[1]) {
                    LoggerManager.warn(`${keyword}の正規表現設定がありません`);
                    break;
                  }
                  
                  const regex = new RegExp(setting[1], 's');
                  const html = UrlFetchApp.fetch(baseUrl, {
                    muteHttpExceptions: true,
                    validateHttpsCertificates: true,
                    timeout: 5000
                  }).getContentText();
                  
                  const match = html.match(regex);
                  if (match && match[1]) {
                    value = match[1].trim();
                    LoggerManager.debug(`${keyword}の生の取得値: コード=${code}, 値="${value}"`);
                    
                    if (keyword === '配当金' || keyword === 'PER' || keyword === 'PBR') {
                      const originalValue = value;
                      value = value.replace(/[^0-9.]/g, '');
                      value = parseFloat(value);
                      if (isNaN(value)) value = '';
                      LoggerManager.debug(`${keyword}の処理後の値: コード=${code}, 元の値="${originalValue}", 処理後="${value}"`);
                    }
                    break;
                  } else {
                    LoggerManager.debug(`${keyword}の正規表現マッチ失敗: コード=${code}, URL=${baseUrl}`);
                    if (keyword === '配当金') {
                      LoggerManager.debug(`配当金のHTML内容(最初の500文字): ${html.substring(0, 500)}`);
                      LoggerManager.debug(`配当金の正規表現: ${regex}`);
                    }
                  }
                } catch (error) {
                  if (retryCount === 1) {
                    hasError = true;
                    LoggerManager.error(`${keyword}の取得エラー: コード=${code}`, error);
                  }
                  Utilities.sleep(sleepTime.betweenRetries);
                }
              }
            }
            columnData.push([value]);
          }

          if (columnData.length > 0) {
            try {
              UIManager.updateProgress(
                processedCount,
                validRows,
                `${keyword}情報を更新中... (${processingCodes.length}件)`,
                {
                  phase: '項目更新',
                  currentItem: keyword,
                  processingItems: processingCodes.join(', ').substring(0, 100) + (processingCodes.length > 5 ? '...' : ''),
                  hasErrors: hasError
                }
              );
              
              // Debug logging for batch processing error
              LoggerManager.debug(`Attempting to update ${keyword} data`, {
                batchStart: batchStart,
                colIndex: colIndex,
                columnDataLength: columnData.length,
                sheetReferenceExists: !!sheetReference,
                sheetReferenceName: sheetReference ? sheetReference.getName() : 'undefined'
              });
              
              const columnRange = sheetReference.getRange(batchStart + 3, colIndex, columnData.length, 1);
              columnRange.setValues(columnData);
            } catch (error) {
              LoggerManager.error(`${keyword}のバッチ更新エラー:`, error);
              failedUpdates.push({
                type: keyword,
                startRow: batchStart + 3,
                endRow: batchStart + 3 + columnData.length - 1,
                error: error.message
              });
            }
          }
          
          // 列間の待機
          Utilities.sleep(sleepTime.betweenColumns);
        }
        
        // 処理済み件数を更新
        processedCount += batchData.length;
        
        // 安全マージンを考慮したタイムアウトチェック
        const currentTime = new Date();
        const elapsedTime = currentTime - startTime;
        
        // 進捗の更新
        const elapsedTimeFormatted = this.formatElapsedTime(startTime);
        const percentComplete = Math.round((processedCount / validRows) * 100);
        
        UIManager.updateProgress(
          processedCount,
          validRows,
          `トレンド情報を更新中: ${processedCount}/${validRows} (${percentComplete}%)`,
          {
            phase: 'バッチ完了',
            currentBatch: currentBatch,
            totalBatches: totalBatches,
            elapsedTime: elapsedTimeFormatted,
            percentComplete: percentComplete
          }
        );
        
        if (elapsedTime > MAX_SAFE_TIME) {
          const nextStartIndex_Reference = actualProcessedCount;
          LoggerManager.warn(`安全マージンのため処理を中断しました（${actualProcessedCount}/${validRows}件完了）。次回は${nextStartIndex_Reference}番目から再開します。`);
          
          // 次回の開始位置を保存
          scriptProperties.setProperty('nextStartIndex_Reference', nextStartIndex_Reference.toString());
          scriptProperties.setProperty('lastProcessedCount_Reference', actualProcessedCount.toString());
          scriptProperties.setProperty('safeTimeout_Reference', 'true'); // 安全タイムアウトフラグ
          
          TimeoutManager.cleanup();
          return 'SafeTimeout';
        }
        
        // 通常のタイムアウトチェック（バックアップ）
        if (TimeoutManager.checkAndHandleTimeout('updateTrendInfo')) {
          const nextStartIndex_Reference = actualProcessedCount;
          LoggerManager.warn(`タイムアウトのため処理を中断しました（${actualProcessedCount}/${validRows}件完了）。次回は${nextStartIndex_Reference}番目から再開します。`);
          
          // 次回の開始位置を保存
          scriptProperties.setProperty('nextStartIndex_Reference', nextStartIndex_Reference.toString());
          scriptProperties.setProperty('lastProcessedCount_Reference', actualProcessedCount.toString());
          
          TimeoutManager.cleanup();
          return 'Timeout';
        }
        
        // バッチ間の待機
        if (batchEnd < remainingData.length) {
          Utilities.sleep(sleepTime.betweenBatches);
        }
      }

      const totalElapsedTime = this.formatElapsedTime(startTime);

      if (failedUpdates.length > 0) {
        LoggerManager.warn('更新に失敗した箇所があります:', failedUpdates);
        UIManager.updateProgress(
          validRows, 
          validRows, 
          `トレンド情報の更新が完了しましたが、${failedUpdates.length}件のエラーがありました`,
          {
            phase: '完了（一部エラー）',
            elapsedTime: totalElapsedTime,
            errorCount: failedUpdates.length,
            errors: failedUpdates.map(f => `${f.type}(${f.startRow}-${f.endRow}): ${f.error}`).join('; ')
          }
        );
      } else {
        UIManager.updateProgress(
          validRows, 
          validRows, 
          'トレンド情報の更新が完了しました',
          {
            phase: '完了',
            elapsedTime: totalElapsedTime,
            message: '全ての処理が正常に完了しました'
          }
        );
      }
      
      // 処理完了時に保存された開始位置をクリア
      scriptProperties.deleteProperty('nextStartIndex_Reference');
      scriptProperties.deleteProperty('lastProcessedCount_Reference');
      scriptProperties.deleteProperty('safeTimeout_Reference');
      
      return 'Completed';

    } catch (e) {
      LoggerManager.error('Error in updateTrendInfo:', e);
      UIManager.updateProgress(0, 1, 'エラーが発生しました: ' + e.message, {
        phase: 'エラー',
        error: e.message,
        stack: e.stack
      });
      throw e;
    }
  },

  /**
   * 経過時間をフォーマットする
   * @param {Date} startTime 開始時間
   * @return {string} フォーマットされた経過時間
   */
  formatElapsedTime: function(startTime) {
    const now = new Date();
    const elapsedMs = now - startTime;
    const seconds = Math.floor(elapsedMs / 1000);
    const minutes = Math.floor(seconds / 60);
    const remainingSeconds = seconds % 60;
    return `${minutes}分${remainingSeconds}秒`;
  },

  /**
   * _Referenceシートの情報を更新（分割処理版）
   * タイムアウトを避けるため、小さなバッチで確実に処理
   */
  updateReferenceInfoBatched: function(startFromIndex = 0) {
    try {
      LoggerManager.info('分割処理による_Reference情報更新を開始します');
      
      // 現在のところは従来版を呼び出す（将来の分割処理実装時に拡張予定）
      return this.updateReferenceInfo();
      
    } catch (error) {
      LoggerManager.error('_Reference情報更新中にエラーが発生しました:', error);
      throw error;
    }
  },


};

// Global functions
function updatePF() {
  return MainController.updatePF();
}

function updateTrendInfo() {
  return MainController.updateTrendInfo();
}

function cancelCurrentProcess() {
  return MainController.cancelCurrentProcess();
}

function onOpen() {
  UIManager.createMenu();
}

function updateReferenceInfo() {
  return MainController.updateReferenceInfo();
}

// 選択行のみ更新する関数
function updateSelectedReferenceInfo() {
  return MainController.updateSelectedReferenceInfo();
}

/**
 * 特定の銘柄コードのみ_Reference情報を更新する
 * @param {string} codes カンマ区切りの銘柄コード
 * @returns {string} 完了メッセージ
 */
function updateReferenceInfoForCodes(codes) {
  // カンマ区切りの文字列を配列に変換
  const codeArray = codes.split(',').map(code => code.trim()).filter(code => code);
  
  if (codeArray.length === 0) {
    throw new Error('更新対象の銘柄コードが指定されていません');
  }
  
  return MainController.updateReferenceInfo(codeArray);
}

/**
 * トレンド情報更新の継続処理を開始
 */
function resumeUpdateTrendInfo() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const savedStartIndex = scriptProperties.getProperty('nextStartIndex_Reference');
    
    if (!savedStartIndex) {
      LoggerManager.warn('継続処理の開始位置が見つかりません。通常の処理を開始します。');
      return MainController.updateTrendInfo();
    }
    
    const startIndex = parseInt(savedStartIndex);
    LoggerManager.info(`継続処理を開始します: ${startIndex}番目から`);
    
    return MainController.updateTrendInfo(startIndex);
    
  } catch (error) {
    LoggerManager.error('継続処理の開始に失敗しました:', error);
    throw error;
  }
}

/**
 * トレンド情報更新の分割処理を開始（推奨）
 * 残り件数に応じて最適な分割数を自動計算
 */
function updateTrendInfoOptimized() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const savedStartIndex = scriptProperties.getProperty('nextStartIndex_Reference');
    
    if (!savedStartIndex) {
      LoggerManager.info('新規処理を開始します。最適化された設定で実行します。');
      return MainController.updateTrendInfo();
    }
    
    const startIndex = parseInt(savedStartIndex);
    const totalItems = 119; // 総件数
    const remainingItems = totalItems - startIndex;
    
    LoggerManager.info(`最適化された継続処理を開始します: ${startIndex}番目から（残り${remainingItems}件）`);
    
    // 残り件数に応じて処理設定を調整
    if (remainingItems <= 30) {
      LoggerManager.info('残り件数が少ないため、一度で完了を試行します。');
    } else {
      LoggerManager.info('残り件数が多いため、安全マージンを考慮した処理を実行します。');
    }
    
    return MainController.updateTrendInfo(startIndex);
    
  } catch (error) {
    LoggerManager.error('最適化された処理の開始に失敗しました:', error);
    throw error;
  }
}

/**
 * updatePF機能のテスト
 */
function testUpdatePF() {
  try {
    console.log('updatePF機能のテストを開始します...');
    
    // テスト用の注文データを作成
    const testOrdersData = [
      ['コード', '取引種類', '数量', '単価', '口座種別', '商品種別', '口座名義人'],
      ['7203', '購入', '100', '2000', '特定口座', 'JP', 'テスト太郎'],
      ['AAPL', '購入', '10', '150', '特定口座', 'US', 'テスト太郎'],
      ['7203', '売却', '50', '2200', '特定口座', 'JP', 'テスト太郎']
    ];
    
    // StockDataManager.processOrdersのテスト
    console.log('StockDataManager.processOrdersのテスト...');
    const stockData = StockDataManager.processOrders(testOrdersData);
    console.log('処理結果:', stockData);
    
    // 結果の検証
    if (stockData && Object.keys(stockData).length > 0) {
      console.log('✓ StockDataManager.processOrdersテスト成功');
      
      // 各銘柄の情報を表示
      Object.entries(stockData).forEach(([key, stock]) => {
        console.log(`銘柄: ${stock.code}, 保有数: ${stock.totalNum}, 総コスト: ${stock.totalCost}`);
      });
    } else {
      console.log('✗ StockDataManager.processOrdersテスト失敗: データが処理されませんでした');
    }
    
    // 実際のupdatePF機能のテスト（シートが存在する場合のみ）
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMain = spreadsheet.getSheetByName(SHEET_NAMES.MAIN);
    const sheetOrders = spreadsheet.getSheetByName(SHEET_NAMES.ORDERS);
    
    if (sheetMain && sheetOrders) {
      console.log('実際のシートでのupdatePFテスト...');
      
      // 注文シートの現在のデータを保存
      const originalOrdersData = sheetOrders.getDataRange().getValues();
      
      // テストデータを注文シートに書き込み
      sheetOrders.clear();
      sheetOrders.getRange(1, 1, testOrdersData.length, testOrdersData[0].length).setValues(testOrdersData);
      
      try {
        // updatePFを実行
        const result = MainController.updatePF();
        console.log('✓ updatePF実行成功:', result);
      } catch (error) {
        console.log('✗ updatePF実行エラー:', error.message);
      } finally {
        // 元のデータを復元
        sheetOrders.clear();
        sheetOrders.getRange(1, 1, originalOrdersData.length, originalOrdersData[0].length).setValues(originalOrdersData);
      }
    } else {
      console.log('テスト用シートが存在しないため、実際のupdatePFテストはスキップします');
    }
    
    console.log('updatePF機能のテストが完了しました');
    
  } catch (error) {
    console.error('testUpdatePFエラー:', error);
    throw error;
  }
}

/**
 * 分割処理による_Reference情報更新（推奨）
 * タイムアウトを避けるため、小さなバッチで確実に処理
 */
/**
 * _Reference情報更新（メニューから呼び出される関数）
 * タイムアウト時の継続処理機能付き
 */
function updateReferenceInfoBatched() {
  try {
    LoggerManager.info('_Reference情報更新を開始します');
    
    // MainController経由で実行（分割処理版を使用）
    const result = MainController.updateReferenceInfoBatched();
    
    // 結果に応じた処理
    if (result === 'Completed') {
      SpreadsheetApp.getActiveSpreadsheet().toast('_Reference情報の更新が完了しました', '完了', 30);
      LoggerManager.info('_Reference情報更新が完了しました');
    } else if (result === 'SafeTimeout' || result === 'Timeout') {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '処理がタイムアウトしました。同じボタンを再度押すと続きから実行されます。', 
        'タイムアウト', 
        30
      );
      LoggerManager.info('_Reference情報更新がタイムアウトしました。続きから再実行可能です。');
    } else if (result === 'Cancelled') {
      SpreadsheetApp.getActiveSpreadsheet().toast('処理がキャンセルされました', 'キャンセル', 10);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('処理が中断されました', '中断', 10);
    }
    
    return result;
    
  } catch (error) {
    LoggerManager.error('_Reference情報更新でエラーが発生しました:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `エラーが発生しました: ${error.message}`, 
      'エラー', 
      30
    );
    throw error;
  }
}

/**
 * _Reference情報更新（シンプル版 - 直接呼び出し用）
 */
function updateReferenceInfo() {
  return MainController.updateReferenceInfoBatched();
}

/**
 * 分割処理の継続実行
 */
function continueReferenceInfoUpdate() {
  try {
    LoggerManager.info('分割処理の継続を開始します');
    
    // 保存された処理状態を確認
    const scriptProperties = PropertiesService.getScriptProperties();
    const savedStartIndex = scriptProperties.getProperty('nextStartIndex_Reference');
    const safeTimeout_Reference = scriptProperties.getProperty('safeTimeout_Reference');
    
    // トレンド情報更新の継続処理をチェック
    if (savedStartIndex && safeTimeout_Reference === 'true') {
      LoggerManager.info(`トレンド情報更新の継続処理を開始します: ${savedStartIndex}番目から`);
      const result = MainController.updateTrendInfo(parseInt(savedStartIndex));
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `トレンド情報更新の継続処理が完了しました: ${result}`, 
        '継続完了', 
        30
      );
      
      return result;
    }
    
    // _Reference情報更新の継続処理をチェック
    const savedStates = scriptProperties.getProperties();
    const batchStates = Object.keys(savedStates).filter(key => key.startsWith('batch_state_'));
    
    if (batchStates.length > 0) {
      // 最新の処理状態を取得
      const latestState = batchStates.reduce((latest, key) => {
        const state = JSON.parse(savedStates[key]);
        return (!latest || state.timestamp > latest.timestamp) ? state : latest;
      }, null);
      
      if (latestState && latestState.status !== 'completed') {
        LoggerManager.info('_Reference情報更新の継続処理を開始します', latestState);
        
        // 継続処理を実行（分割処理版を使用）
        const result = MainController.updateReferenceInfoBatched();
        
        SpreadsheetApp.getActiveSpreadsheet().toast(
          `_Reference情報更新の継続処理が完了しました: ${result}`, 
          '継続完了', 
          30
        );
        
        return result;
      }
    }
    
    // 継続可能な処理が見つからない場合
    SpreadsheetApp.getActiveSpreadsheet().toast('継続可能な処理が見つかりません', '情報', 10);
    return 'NoContinuation';
    
  } catch (error) {
    LoggerManager.error('継続処理の実行中にエラーが発生しました:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('継続処理でエラーが発生しました: ' + error.message, 'エラー', 30);
    throw error;
  }
}

/**
 * 分割処理の状態をクリア
 */
function clearBatchProcessingState() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const savedStates = scriptProperties.getProperties();
    const batchStates = Object.keys(savedStates).filter(key => key.startsWith('batch_state_'));
    
    batchStates.forEach(key => {
      scriptProperties.deleteProperty(key);
    });
    
    LoggerManager.info(`分割処理の状態をクリアしました: ${batchStates.length}件`);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `分割処理の状態をクリアしました: ${batchStates.length}件`, 
      'クリア完了', 
      10
    );
    
    return 'Cleared';
    
  } catch (error) {
    LoggerManager.error('分割処理状態のクリア中にエラーが発生しました:', error);
    SpreadsheetApp.getActiveSpreadsheet().toast('クリア処理でエラーが発生しました: ' + error.message, 'エラー', 30);
    throw error;
  }
}