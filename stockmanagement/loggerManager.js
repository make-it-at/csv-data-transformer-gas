// loggerManager.gs
// Version 1.3
// Last updated: 2025-03-02
// Changes: Fixed SystemLog sheet creation and logging

var LoggerManager = {
  // シートにログを書き込む
  writeLog: function(level, message, data) {
    try {
      var sheet = this.getOrCreateLogSheet();
      var now = new Date();
      var user = Session.getActiveUser().getEmail();
      
      // データをJSON文字列に変換
      var dataStr = data ? JSON.stringify(data) : '';
      
      // ログを追加
      sheet.appendRow([
        now,
        level,
        message,
        dataStr,
        user
      ]);
      
      // コンソールにも出力
      console.log(`[${now.toISOString()}] [${level}] ${message} ${dataStr}`);
    } catch (e) {
      console.error('Error writing to log:', e);
    }
  },

  // ログシートを取得または作成
  getOrCreateLogSheet: function() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SystemLog');
    
    if (!sheet) {
      sheet = ss.insertSheet('SystemLog');
      var headers = [['Timestamp', 'Level', 'Message', 'Data', 'User']];
      sheet.getRange(1, 1, 1, 5).setValues(headers);
      sheet.setFrozenRows(1);
      
      // 列幅の設定
      sheet.setColumnWidth(1, 180);  // Timestamp
      sheet.setColumnWidth(2, 80);   // Level
      sheet.setColumnWidth(3, 300);  // Message
      sheet.setColumnWidth(4, 400);  // Data
      sheet.setColumnWidth(5, 150);  // User
    }
    
    return sheet;
  },

  // ログレベルごとのメソッド
  debug: function(message, data) {
    this.writeLog('DEBUG', message, data);
  },

  info: function(message, data) {
    this.writeLog('INFO', message, data);
  },

  warn: function(message, data) {
    this.writeLog('WARN', message, data);
  },

  error: function(message, data) {
    this.writeLog('ERROR', message, data);
  },

  // ログをクリア
  clearLogs: function() {
    try {
      var sheet = this.getOrCreateLogSheet();
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 5).clear();
      }
    } catch (e) {
      console.error('Error clearing logs:', e);
    }
  }
};

// グローバル関数
function testLogger() {
  LoggerManager.info('Logger test', { test: true });
  LoggerManager.debug('Debug message', { debug: true });
  LoggerManager.warn('Warning message', { warning: true });
  LoggerManager.error('Error message', { error: true });
}

function clearSystemLogs() {
  LoggerManager.clearLogs();
}