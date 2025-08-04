/**
 * ログ記録・管理機能
 * 
 * 全ての処理履歴をログシートに記録・管理
 */

// ログレベル定義
const LOG_LEVELS = {
  DEBUG: 0,
  INFO: 1,
  WARN: 2,
  ERROR: 3
};

// ログ設定
const LOG_CONFIG = {
  MAX_ENTRIES: 1000,        // 最大ログエントリ数
  CLEANUP_BATCH: 100,       // クリーンアップ時の削除数
  RETENTION_DAYS: 30        // ログ保持日数
};

/**
 * ログをスプレッドシートに記録する
 * 
 * @param {string} level - ログレベル (DEBUG, INFO, WARN, ERROR)
 * @param {string} category - カテゴリ
 * @param {string} message - ログメッセージ
 * @param {Object} data - 追加データ
 */
function writeLog(level, category, message, data = null) {
  try {
    // ログレベルの検証
    if (!Object.keys(LOG_LEVELS).includes(level)) {
      level = 'INFO';
    }
    
    const timestamp = new Date();
    const dataString = data ? safeStringify(data) : '';
    
    // コンソールログにも出力
    const consoleMessage = `[${level}][${category}] ${message}`;
    Logger.log(consoleMessage);
    
    // ログシートに記録
    const logSheet = getLogSheet();
    if (logSheet) {
      logSheet.appendRow([timestamp, level, category, message, dataString]);
      
      // ログローテーション
      manageLogRotation(logSheet);
    }
    
  } catch (error) {
    // ログ記録でエラーが発生した場合はコンソールのみに出力
    Logger.log(`[logger.gs] ログ記録エラー: ${error.message}`);
    Logger.log(`[logger.gs] 元のログ: [${level}][${category}] ${message}`);
  }
}

/**
 * ログシートを取得（存在しない場合は作成）
 * 
 * @return {Sheet} ログシート
 */
function getLogSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = spreadsheet.getSheetByName(SHEET_NAMES.LOG);
    
    if (!logSheet) {
      // ログシートが存在しない場合は作成
      logSheet = spreadsheet.insertSheet(SHEET_NAMES.LOG);
      initializeLogSheet(logSheet);
      Logger.log('[logger.gs] ログシートを作成しました');
    }
    
    return logSheet;
    
  } catch (error) {
    Logger.log(`[logger.gs] ログシート取得エラー: ${error.message}`);
    return null;
  }
}

/**
 * ログローテーション管理
 * 
 * @param {Sheet} logSheet - ログシート
 */
function manageLogRotation(logSheet) {
  try {
    const lastRow = logSheet.getLastRow();
    
    // 最大エントリ数を超えた場合、古いログを削除
    if (lastRow > LOG_CONFIG.MAX_ENTRIES + 1) { // +1 はヘッダー行
      const deleteCount = LOG_CONFIG.CLEANUP_BATCH;
      const startRow = 2; // ヘッダーの次の行から
      
      logSheet.deleteRows(startRow, deleteCount);
      
      Logger.log(`[logger.gs] 古いログを削除: ${deleteCount}件`);
    }
    
  } catch (error) {
    Logger.log(`[logger.gs] ログローテーションエラー: ${error.message}`);
  }
}

/**
 * 最新のログエントリを取得
 * 
 * @param {number} count - 取得するエントリ数
 * @return {Array<Object>} ログエントリ配列
 */
function getRecentLogs(count = 50) {
  try {
    const logSheet = getLogSheet();
    if (!logSheet) {
      return [];
    }
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) { // ヘッダーのみ
      return [];
    }
    
    // 取得する行数を計算
    const actualCount = Math.min(count, lastRow - 1);
    const startRow = Math.max(2, lastRow - actualCount + 1);
    
    // データを取得
    const range = logSheet.getRange(startRow, 1, actualCount, 5);
    const values = range.getValues();
    
    // オブジェクト配列に変換
    const logs = values.map(row => ({
      timestamp: row[0],
      level: row[1],
      category: row[2],
      message: row[3],
      data: row[4]
    }));
    
    // 新しい順にソート
    logs.reverse();
    
    return logs;
    
  } catch (error) {
    Logger.log(`[logger.gs] ログ取得エラー: ${error.message}`);
    return [];
  }
}

/**
 * 古いログエントリを削除
 * 
 * @param {number} days - 保持日数
 */
function clearOldLogs(days = LOG_CONFIG.RETENTION_DAYS) {
  try {
    const logSheet = getLogSheet();
    if (!logSheet) {
      return;
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      return;
    }
    
    // タイムスタンプ列を取得
    const timestampRange = logSheet.getRange(2, 1, lastRow - 1, 1);
    const timestamps = timestampRange.getValues();
    
    // 削除対象行を特定
    let deleteCount = 0;
    for (let i = 0; i < timestamps.length; i++) {
      const timestamp = timestamps[i][0];
      if (timestamp instanceof Date && timestamp < cutoffDate) {
        deleteCount++;
      } else {
        break; // 日付は昇順なので、新しい日付が見つかったら終了
      }
    }
    
    // 古いログを削除
    if (deleteCount > 0) {
      logSheet.deleteRows(2, deleteCount);
      Logger.log(`[logger.gs] 古いログを削除: ${deleteCount}件 (${days}日以前)`);
    }
    
  } catch (error) {
    Logger.log(`[logger.gs] 古いログ削除エラー: ${error.message}`);
  }
}

/**
 * ログエントリをフォーマット
 * 
 * @param {Object} entry - ログエントリ
 * @return {string} フォーマットされたログ
 */
function formatLogEntry(entry) {
  try {
    if (!entry) {
      return '';
    }
    
    const timestamp = entry.timestamp instanceof Date 
      ? getTimestamp(entry.timestamp) 
      : entry.timestamp;
    
    const level = entry.level || 'INFO';
    const category = entry.category || 'SYSTEM';
    const message = entry.message || '';
    
    let formatted = `[${timestamp}] [${level}] [${category}] ${message}`;
    
    // 詳細データがある場合は追加
    if (entry.data && entry.data.trim() !== '') {
      formatted += `\n  詳細: ${entry.data}`;
    }
    
    return formatted;
    
  } catch (error) {
    Logger.log(`[logger.gs] ログフォーマットエラー: ${error.message}`);
    return `[ERROR] ログフォーマットエラー: ${error.message}`;
  }
}

/**
 * ログレベル別の件数を取得
 * 
 * @param {number} days - 集計対象日数
 * @return {Object} レベル別件数
 */
function getLogStatistics(days = 7) {
  try {
    const logSheet = getLogSheet();
    if (!logSheet) {
      return {};
    }
    
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      return {};
    }
    
    // データを取得
    const range = logSheet.getRange(2, 1, lastRow - 1, 2);
    const values = range.getValues();
    
    // 統計を集計
    const stats = {
      DEBUG: 0,
      INFO: 0,
      WARN: 0,
      ERROR: 0,
      total: 0
    };
    
    values.forEach(row => {
      const timestamp = row[0];
      const level = row[1];
      
      // 指定期間内のログのみカウント
      if (timestamp instanceof Date && timestamp >= cutoffDate) {
        if (stats.hasOwnProperty(level)) {
          stats[level]++;
        }
        stats.total++;
      }
    });
    
    return stats;
    
  } catch (error) {
    Logger.log(`[logger.gs] ログ統計エラー: ${error.message}`);
    return {};
  }
}

/**
 * 特定カテゴリのログを検索
 * 
 * @param {string} category - カテゴリ名
 * @param {number} limit - 取得件数制限
 * @return {Array<Object>} 該当ログエントリ
 */
function searchLogsByCategory(category, limit = 100) {
  try {
    const logSheet = getLogSheet();
    if (!logSheet) {
      return [];
    }
    
    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      return [];
    }
    
    // データを取得
    const range = logSheet.getRange(2, 1, lastRow - 1, 5);
    const values = range.getValues();
    
    // カテゴリでフィルタリング
    const matchingLogs = values
      .filter(row => row[2] === category)
      .slice(-limit) // 最新のN件
      .map(row => ({
        timestamp: row[0],
        level: row[1],
        category: row[2],
        message: row[3],
        data: row[4]
      }));
    
    // 新しい順にソート
    matchingLogs.reverse();
    
    return matchingLogs;
    
  } catch (error) {
    Logger.log(`[logger.gs] ログ検索エラー: ${error.message}`);
    return [];
  }
}

/**
 * ログシートの初期化（再実装 - main.gsとの重複回避）
 * 
 * @param {Sheet} logSheet - ログシート
 */
function initializeLogSheet(logSheet) {
  try {
    const headers = ['タイムスタンプ', 'レベル', 'カテゴリ', 'メッセージ', '詳細データ'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // ヘッダー行のフォーマット
    const headerRange = logSheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#f0f0f0');
    
    // 列幅の調整
    logSheet.setColumnWidth(1, 150); // タイムスタンプ
    logSheet.setColumnWidth(2, 80);  // レベル
    logSheet.setColumnWidth(3, 120); // カテゴリ
    logSheet.setColumnWidth(4, 200); // メッセージ
    logSheet.setColumnWidth(5, 300); // 詳細データ
    
    Logger.log('[logger.gs] ログシート初期化完了');
    
  } catch (error) {
    Logger.log(`[logger.gs] ログシート初期化エラー: ${error.message}`);
  }
}