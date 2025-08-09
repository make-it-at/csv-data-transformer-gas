// logViewer.gs
// Version 1.1
// Last updated: 2025-03-02
// ログ確認用スクリプト

/**
 * システムログを表示する
 */
function showSystemLogs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SystemLog');
    if (!sheet) {
      console.log('SystemLogシートが見つかりません');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('ログデータがありません');
      return;
    }
    
    // 最新の10件を取得
    const startRow = Math.max(2, lastRow - 9);
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 5);
    const logs = range.getValues();
    
    let message = '最新のログ（最新10件）:\n\n';
    logs.forEach((log, index) => {
      const [timestamp, level, msg, data, user] = log;
      message += `${index + 1}. [${level}] ${msg}\n`;
      message += `   時刻: ${timestamp}\n`;
      message += `   ユーザー: ${user}\n`;
      if (data) {
        message += `   データ: ${data}\n`;
      }
      message += '\n';
    });
    
    console.log(message);
    
    // UIが利用可能な場合のみダイアログを表示
    try {
      SpreadsheetApp.getUi().alert('システムログ', message, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (uiError) {
      console.log('UIダイアログは利用できません。コンソールにログを出力しました。');
    }
    
  } catch (error) {
    console.error('Error showing system logs:', error);
    try {
      SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    } catch (uiError) {
      console.error('UIダイアログエラー:', uiError);
    }
  }
}

/**
 * 進捗ログを表示する
 */
function showProgressLogs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProgressLog');
    if (!sheet) {
      console.log('ProgressLogシートが見つかりません');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('進捗ログデータがありません');
      return;
    }
    
    // 最新の10件を取得
    const startRow = Math.max(2, lastRow - 9);
    const range = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6);
    const logs = range.getValues();
    
    let message = '最新の進捗ログ（最新10件）:\n\n';
    logs.forEach((log, index) => {
      const [timestamp, progress, status, current, total, user] = log;
      message += `${index + 1}. ${status}\n`;
      message += `   進捗: ${progress}% (${current}/${total})\n`;
      message += `   時刻: ${timestamp}\n`;
      message += `   ユーザー: ${user}\n\n`;
    });
    
    console.log(message);
    
    // UIが利用可能な場合のみダイアログを表示
    try {
      SpreadsheetApp.getUi().alert('進捗ログ', message, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (uiError) {
      console.log('UIダイアログは利用できません。コンソールにログを出力しました。');
    }
    
  } catch (error) {
    console.error('Error showing progress logs:', error);
    try {
      SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    } catch (uiError) {
      console.error('UIダイアログエラー:', uiError);
    }
  }
}

/**
 * ログシートを開く
 */
function openLogSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // SystemLogシートを開く
    let systemLogSheet = spreadsheet.getSheetByName('SystemLog');
    if (!systemLogSheet) {
      systemLogSheet = spreadsheet.insertSheet('SystemLog');
      const headers = [['Timestamp', 'Level', 'Message', 'Data', 'User']];
      systemLogSheet.getRange(1, 1, 1, 5).setValues(headers);
      systemLogSheet.setFrozenRows(1);
      
      // 列幅の設定
      systemLogSheet.setColumnWidth(1, 180);  // Timestamp
      systemLogSheet.setColumnWidth(2, 80);   // Level
      systemLogSheet.setColumnWidth(3, 300);  // Message
      systemLogSheet.setColumnWidth(4, 400);  // Data
      systemLogSheet.setColumnWidth(5, 150);  // User
    }
    
    // ProgressLogシートを開く
    let progressLogSheet = spreadsheet.getSheetByName('ProgressLog');
    if (!progressLogSheet) {
      progressLogSheet = spreadsheet.insertSheet('ProgressLog');
      const headers = [['Timestamp', 'Progress (%)', 'Status', 'Current', 'Total', 'User']];
      progressLogSheet.getRange(1, 1, 1, 6).setValues(headers);
      progressLogSheet.setFrozenRows(1);
      
      // 列幅の設定
      progressLogSheet.setColumnWidth(1, 180);  // Timestamp
      progressLogSheet.setColumnWidth(2, 100);  // Progress
      progressLogSheet.setColumnWidth(3, 300);  // Status
      progressLogSheet.setColumnWidth(4, 80);   // Current
      progressLogSheet.setColumnWidth(5, 80);   // Total
      progressLogSheet.setColumnWidth(6, 150);  // User
    }
    
    // SystemLogシートをアクティブにする
    spreadsheet.setActiveSheet(systemLogSheet);
    
    console.log('ログシートを開きました。SystemLogシートがアクティブになりました。');
    
    // UIが利用可能な場合のみダイアログを表示
    try {
      SpreadsheetApp.getUi().alert('ログシートを開きました', 'SystemLogシートがアクティブになりました');
    } catch (uiError) {
      console.log('UIダイアログは利用できません。コンソールにメッセージを出力しました。');
    }
    
  } catch (error) {
    console.error('Error opening log sheets:', error);
    try {
      SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    } catch (uiError) {
      console.error('UIダイアログエラー:', uiError);
    }
  }
}

/**
 * ログをクリアする
 */
function clearAllLogs() {
  try {
    let shouldClear = true;
    
    // UIが利用可能な場合は確認ダイアログを表示
    try {
      const response = SpreadsheetApp.getUi().alert(
        '確認',
        'すべてのログをクリアしますか？\nこの操作は元に戻せません。',
        SpreadsheetApp.getUi().ButtonSet.YES_NO
      );
      shouldClear = (response === SpreadsheetApp.getUi().Button.YES);
    } catch (uiError) {
      console.log('UIダイアログは利用できません。ログをクリアします。');
    }
    
    if (shouldClear) {
      // システムログをクリア
      if (typeof LoggerManager !== 'undefined') {
        LoggerManager.clearLogs();
      }
      
      // 進捗ログをクリア
      const progressSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ProgressLog');
      if (progressSheet) {
        const lastRow = progressSheet.getLastRow();
        if (lastRow > 1) {
          progressSheet.getRange(2, 1, lastRow - 1, 6).clear();
        }
      }
      
      console.log('すべてのログをクリアしました');
      
      // UIが利用可能な場合のみダイアログを表示
      try {
        SpreadsheetApp.getUi().alert('すべてのログをクリアしました');
      } catch (uiError) {
        console.log('UIダイアログは利用できません。コンソールにメッセージを出力しました。');
      }
    }
    
  } catch (error) {
    console.error('Error clearing logs:', error);
    try {
      SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    } catch (uiError) {
      console.error('UIダイアログエラー:', uiError);
    }
  }
}

/**
 * ログの概要を表示する
 */
function showLogSummary() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    let summary = 'ログ概要:\n\n';
    
    // SystemLogの概要
    const systemLogSheet = spreadsheet.getSheetByName('SystemLog');
    if (systemLogSheet) {
      const systemLogRows = systemLogSheet.getLastRow();
      summary += `SystemLog: ${systemLogRows > 1 ? systemLogRows - 1 : 0}件のログ\n`;
    } else {
      summary += 'SystemLog: シートが存在しません\n';
    }
    
    // ProgressLogの概要
    const progressLogSheet = spreadsheet.getSheetByName('ProgressLog');
    if (progressLogSheet) {
      const progressLogRows = progressLogSheet.getLastRow();
      summary += `ProgressLog: ${progressLogRows > 1 ? progressLogRows - 1 : 0}件のログ\n`;
    } else {
      summary += 'ProgressLog: シートが存在しません\n';
    }
    
    console.log(summary);
    
    // UIが利用可能な場合のみダイアログを表示
    try {
      SpreadsheetApp.getUi().alert('ログ概要', summary, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (uiError) {
      console.log('UIダイアログは利用できません。コンソールにログを出力しました。');
    }
    
  } catch (error) {
    console.error('Error showing log summary:', error);
    try {
      SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error.message);
    } catch (uiError) {
      console.error('UIダイアログエラー:', uiError);
    }
  }
} 