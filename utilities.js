/**
 * 共通ユーティリティ機能
 * 
 * 各モジュールで共通して使用される関数群
 */

/**
 * 指定した名前のシートを安全に取得する
 * 
 * @param {string} sheetName - シート名
 * @return {Sheet} シートオブジェクト
 * @throws {Error} シートが見つからない場合
 */
function getSheetSafely(sheetName) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シートが見つかりません: ${sheetName}`);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log(`[utilities.gs] シート取得エラー: ${error.message}`);
    throw error;
  }
}

/**
 * ファイルサイズを人間が読みやすい形式にフォーマット
 * 
 * @param {number} bytes - バイト数
 * @return {string} フォーマットされたファイルサイズ
 */
function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.floor(Math.log(bytes) / Math.log(1024));
  const size = (bytes / Math.pow(1024, index)).toFixed(2);
  
  return `${size} ${units[index]}`;
}

/**
 * タイムスタンプを生成
 * 
 * @param {Date} date - 日付オブジェクト（省略時は現在時刻）
 * @return {string} フォーマットされたタイムスタンプ
 */
function getTimestamp(date = null) {
  const targetDate = date || new Date();
  
  const year = targetDate.getFullYear();
  const month = String(targetDate.getMonth() + 1).padStart(2, '0');
  const day = String(targetDate.getDate()).padStart(2, '0');
  const hours = String(targetDate.getHours()).padStart(2, '0');
  const minutes = String(targetDate.getMinutes()).padStart(2, '0');
  const seconds = String(targetDate.getSeconds()).padStart(2, '0');
  
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * 入力値の検証
 * 
 * @param {any} input - 検証対象の値
 * @param {Object} rules - 検証ルール
 * @return {Object} 検証結果
 */
function validateInput(input, rules = {}) {
  const result = {
    isValid: true,
    errors: []
  };
  
  try {
    // 必須チェック
    if (rules.required && (input === null || input === undefined || input === '')) {
      result.isValid = false;
      result.errors.push('必須項目です');
      return result;
    }
    
    // 型チェック
    if (rules.type && input !== null && input !== undefined) {
      const actualType = typeof input;
      if (actualType !== rules.type) {
        result.isValid = false;
        result.errors.push(`型が正しくありません。期待値: ${rules.type}, 実際: ${actualType}`);
      }
    }
    
    // 最小値チェック
    if (rules.min !== undefined && input < rules.min) {
      result.isValid = false;
      result.errors.push(`最小値 ${rules.min} を下回っています`);
    }
    
    // 最大値チェック
    if (rules.max !== undefined && input > rules.max) {
      result.isValid = false;
      result.errors.push(`最大値 ${rules.max} を超えています`);
    }
    
    // 最小長チェック
    if (rules.minLength !== undefined && input.length < rules.minLength) {
      result.isValid = false;
      result.errors.push(`最小長 ${rules.minLength} を下回っています`);
    }
    
    // 最大長チェック
    if (rules.maxLength !== undefined && input.length > rules.maxLength) {
      result.isValid = false;
      result.errors.push(`最大長 ${rules.maxLength} を超えています`);
    }
    
    // パターンマッチング
    if (rules.pattern && !rules.pattern.test(input)) {
      result.isValid = false;
      result.errors.push('形式が正しくありません');
    }
    
  } catch (error) {
    result.isValid = false;
    result.errors.push(`検証エラー: ${error.message}`);
  }
  
  return result;
}

/**
 * メッセージを表示する（将来的にはHTMLサイドバーに表示）
 * 
 * @param {string} type - メッセージタイプ (success, warning, error, info)
 * @param {string} message - メッセージ内容
 * @param {Object} options - 表示オプション
 */
function showMessage(type, message, options = {}) {
  try {
    const timestamp = getTimestamp();
    const logMessage = `[${type.toUpperCase()}] ${message}`;
    
    // コンソールログに出力
    Logger.log(`[utilities.gs] ${timestamp} - ${logMessage}`);
    
    // 将来的にはHTMLサイドバーにも表示
    // 現在はログのみ
    
    // エラーの場合はアラートも表示
    if (type === 'error' && options.showAlert !== false) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('エラー', message, ui.ButtonSet.OK);
    }
    
  } catch (error) {
    Logger.log(`[utilities.gs] メッセージ表示エラー: ${error.message}`);
  }
}

/**
 * 安全なJSON.stringify（循環参照対応）
 * 
 * @param {any} obj - JSON化するオブジェクト
 * @param {number} maxDepth - 最大深度
 * @return {string} JSON文字列
 */
function safeStringify(obj, maxDepth = 3) {
  const seen = new WeakSet();
  
  function replacer(key, value) {
    if (typeof value === 'object' && value !== null) {
      if (seen.has(value)) {
        return '[Circular Reference]';
      }
      seen.add(value);
    }
    return value;
  }
  
  try {
    return JSON.stringify(obj, replacer, 2);
  } catch (error) {
    return `[JSON化エラー: ${error.message}]`;
  }
}

/**
 * 配列を指定サイズのチャンクに分割
 * 
 * @param {Array} array - 分割する配列
 * @param {number} chunkSize - チャンクサイズ
 * @return {Array<Array>} 分割された配列
 */
function chunkArray(array, chunkSize) {
  if (!Array.isArray(array) || chunkSize <= 0) {
    return [];
  }
  
  const chunks = [];
  for (let i = 0; i < array.length; i += chunkSize) {
    chunks.push(array.slice(i, i + chunkSize));
  }
  
  return chunks;
}

/**
 * 処理時間を測定するヘルパー
 * 
 * @param {Function} func - 実行する関数
 * @param {string} label - ラベル
 * @return {any} 関数の戻り値
 */
function measureExecutionTime(func, label = 'Function') {
  const startTime = new Date().getTime();
  
  try {
    const result = func();
    const endTime = new Date().getTime();
    const executionTime = endTime - startTime;
    
    Logger.log(`[utilities.gs] ${label} 実行時間: ${executionTime}ms`);
    
    return result;
    
  } catch (error) {
    const endTime = new Date().getTime();
    const executionTime = endTime - startTime;
    
    Logger.log(`[utilities.gs] ${label} エラー発生 (実行時間: ${executionTime}ms): ${error.message}`);
    throw error;
  }
}

/**
 * 実行時間制限チェック
 * 
 * @param {number} startTime - 開始時刻（ミリ秒）
 * @param {number} maxTime - 最大実行時間（ミリ秒）
 * @return {boolean} 制限を超えているかどうか
 */
function isTimeoutReached(startTime, maxTime = MAX_EXECUTION_TIME) {
  const currentTime = new Date().getTime();
  return (currentTime - startTime) > maxTime;
}

/**
 * エスケープHTML
 * 
 * @param {string} text - エスケープするテキスト
 * @return {string} エスケープされたテキスト
 */
function escapeHtml(text) {
  if (typeof text !== 'string') {
    return text;
  }
  
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  };
  
  return text.replace(/[&<>"']/g, function(m) {
    return map[m];
  });
}

/**
 * 数値の安全な変換
 * 
 * @param {any} value - 変換する値
 * @param {number} defaultValue - デフォルト値
 * @return {number} 変換された数値
 */
function safeParseNumber(value, defaultValue = 0) {
  if (typeof value === 'number' && !isNaN(value)) {
    return value;
  }
  
  if (typeof value === 'string') {
    const parsed = parseFloat(value);
    if (!isNaN(parsed)) {
      return parsed;
    }
  }
  
  return defaultValue;
}

/**
 * 文字列の安全なトリム
 * 
 * @param {any} value - トリムする値
 * @param {string} defaultValue - デフォルト値
 * @return {string} トリムされた文字列
 */
function safeTrim(value, defaultValue = '') {
  if (typeof value === 'string') {
    return value.trim();
  }
  
  if (value !== null && value !== undefined) {
    return String(value).trim();
  }
  
  return defaultValue;
}