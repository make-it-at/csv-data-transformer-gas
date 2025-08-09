// timeoutManager.gs
// Version 1.4
// Last updated: 2025-03-02
// Changes: Improved trigger management and cleanup

var TimeoutManager = {
  DISPLAY_NAMES: {
    'updateMainInfo': 'メイン情報更新',
    'updatePF': '保有株リスト更新',
    'clearPFData': 'リストクリア',
    'updateStockInfoByFrequency': '情報更新',
    'processBatch': 'バッチ処理'
  },

  TIMEOUT_MINUTES: 5,

  // 実行開始時間
  startTime: null,
  
  // 最大実行時間（ミリ秒）- デフォルトは5分30秒（6分の制限に対して余裕を持たせる）
  MAX_EXECUTION_TIME: 5.5 * 60 * 1000,

  createTimeoutTrigger: function(functionName) {
    this.deleteAllTriggers(); // 既存のトリガーをすべて削除
    try {
      const trigger = ScriptApp.newTrigger('handleTimeoutWrapper')
        .timeBased()
        .after(this.TIMEOUT_MINUTES * 60 * 1000)
        .create();
      
      PropertiesService.getScriptProperties().setProperty('runningFunction', functionName);
      PropertiesService.getScriptProperties().setProperty('currentTriggerId', trigger.getUniqueId());
      
      LoggerManager.debug('タイムアウトトリガーを作成しました', {
        function: functionName,
        timeoutMinutes: this.TIMEOUT_MINUTES,
        triggerId: trigger.getUniqueId()
      });
    } catch (e) {
      LoggerManager.error('タイムアウトトリガーの作成に失敗', e);
      LoggerManager.warn('警告: 処理は継続されますが、長時間の実行には注意してください。');
    }
  },

  deleteAllTriggers: function() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      try {
        ScriptApp.deleteTrigger(trigger);
        LoggerManager.debug('トリガーを削除しました', {
          triggerId: trigger.getUniqueId(),
          handlerFunction: trigger.getHandlerFunction()
        });
      } catch (e) {
        LoggerManager.error('トリガーの削除に失敗', {
          triggerId: trigger.getUniqueId(),
          error: e.message
        });
      }
    });
    
    // プロパティもクリア
    PropertiesService.getScriptProperties().deleteProperty('currentTriggerId');
    PropertiesService.getScriptProperties().deleteProperty('runningFunction');
  },

  deleteTimeoutTrigger: function() {
    const currentTriggerId = PropertiesService.getScriptProperties().getProperty('currentTriggerId');
    const triggers = ScriptApp.getProjectTriggers();
    
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'handleTimeoutWrapper' || 
          (currentTriggerId && trigger.getUniqueId() === currentTriggerId)) {
        try {
          ScriptApp.deleteTrigger(trigger);
          LoggerManager.debug('タイムアウトトリガーを削除しました', {
            triggerId: trigger.getUniqueId()
          });
        } catch (e) {
          LoggerManager.error('タイムアウトトリガーの削除に失敗', e);
        }
      }
    });
    
    PropertiesService.getScriptProperties().deleteProperty('currentTriggerId');
  },

  handleTimeout: function(functionName) {
    const displayName = this.DISPLAY_NAMES[functionName] || functionName;
    const message = displayName + 'の実行がタイムアウトしました。';
    
    LoggerManager.warn('タイムアウトが発生しました', {
      function: functionName,
      displayName: displayName,
      timeoutMinutes: this.TIMEOUT_MINUTES
    });
    
    PropertiesService.getScriptProperties().setProperty(
      'latestProgress',
      JSON.stringify({
        progress: 100,
        status: message,
        error: true,
        timeoutAt: new Date().toISOString()
      })
    );
    
    // シート上にタイムアウト情報を記録
    this.logTimeoutToSheet(functionName, message);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(message, 'タイムアウト', 30);
  },

  logTimeoutToSheet: function(functionName, message) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName('TimeoutLog');
      
      if (!sheet) {
        sheet = ss.insertSheet('TimeoutLog');
        sheet.getRange('A1:D1').setValues([['Timestamp', 'Function', 'Message', 'User']]);
        sheet.setFrozenRows(1);
      }

      sheet.appendRow([
        new Date(),
        functionName,
        message,
        Session.getActiveUser().getEmail()
      ]);
    } catch (e) {
      LoggerManager.error('タイムアウトログの記録に失敗', e);
    }
  },

  /**
   * 実行時間がタイムアウト制限に近づいているかチェック
   * @return {boolean} タイムアウトに近づいている場合はtrue
   */
  isExecutionTimedOut: function() {
    if (!this.startTime) {
      return false;
    }
    
    const currentTime = new Date();
    const elapsedTime = currentTime - this.startTime;
    
    // 残り時間をログに出力（30秒ごと）
    if (elapsedTime % 30000 < 1000) {
      const remainingTime = Math.max(0, (this.MAX_EXECUTION_TIME - elapsedTime) / 1000);
      LoggerManager.info(`タイムアウト管理: 残り実行時間 約${Math.round(remainingTime)}秒`);
    }
    
    return elapsedTime > this.MAX_EXECUTION_TIME;
  },

  /**
   * 実行タイマーを開始
   */
  startExecutionTimer: function() {
    this.startTime = new Date();
    LoggerManager.info('タイムアウト管理: 実行タイマーを開始しました');
  },

  clearExecutionTimer: function() {
    PropertiesService.getScriptProperties().deleteProperty('executionStartTime');
    LoggerManager.debug('実行タイマーをクリア');
  },

  checkAndHandleTimeout: function(functionName) {
    if (this.isExecutionTimedOut()) {
      this.handleTimeout(functionName);
      return true;
    }
    return false;
  },

  calculateTimeRemaining: function() {
    const startTime = PropertiesService.getScriptProperties().getProperty('executionStartTime');
    if (!startTime) return this.TIMEOUT_MINUTES * 60;

    const currentTime = new Date().getTime();
    const elapsedSeconds = Math.floor((currentTime - parseInt(startTime)) / 1000);
    const remainingSeconds = Math.max(0, (this.TIMEOUT_MINUTES * 60) - elapsedSeconds);

    LoggerManager.debug('残り実行時間を計算', {
      elapsedSeconds: elapsedSeconds,
      remainingSeconds: remainingSeconds
    });

    return remainingSeconds;
  },

  /**
   * クリーンアップ処理
   */
  cleanup: function() {
    this.deleteTimeoutTrigger();
    this.clearExecutionTimer();
    PropertiesService.getScriptProperties().deleteProperty('runningFunction');
    this.startTime = null;
    LoggerManager.info('タイムアウト管理: クリーンアップ完了');
  },
  
  /**
   * キャンセルフラグをリセットする
   * @returns {boolean} 成功した場合はtrue
   */
  resetCancellationFlag: function() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.deleteProperty('cancelFlag');
      LoggerManager.debug('キャンセルフラグをリセットしました');
      return true;
    } catch (e) {
      LoggerManager.warn('キャンセルフラグのリセットに失敗しました:', e);
      return false;
    }
  },
  
  /**
   * 処理をキャンセルすべきかどうかを確認する
   * @returns {boolean} キャンセルすべき場合はtrue
   */
  shouldCancel: function() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const cancelFlag = scriptProperties.getProperty('cancelFlag');
      const isCancelled = cancelFlag === 'true';
      
      if (isCancelled) {
        LoggerManager.debug('キャンセルフラグが検出されました');
      }
      
      return isCancelled;
    } catch (e) {
      LoggerManager.warn('キャンセルフラグの確認に失敗しました:', e);
      return false;
    }
  }
};

// グローバル関数
function createTimeoutTrigger(functionName) {
  return TimeoutManager.createTimeoutTrigger(functionName);
}

function deleteTimeoutTrigger() {
  return TimeoutManager.deleteTimeoutTrigger();
}

function handleTimeoutWrapper() {
  const functionName = PropertiesService.getScriptProperties().getProperty('runningFunction');
  if (functionName) {
    TimeoutManager.handleTimeout(functionName);
  }
  TimeoutManager.cleanup();
}