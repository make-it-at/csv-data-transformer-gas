// uiManager.gs
// Version 1.8
// Last updated: 2024-05-01

// グローバル変数の定義
var SHEET_NAMES = {
  SETTING: 'Setting',
  MAIN: 'Main',
  ORDERS: '注文一覧',
  REFERENCE: '_Reference'
};

// UIManagerの定義
var UIManager = {
  showSidebar: function() {
    var html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('情報更新')
      .setWidth(300);
    SpreadsheetApp.getUi().showSidebar(html);
  },

  /**
   * 統合メニューを作成する
   * すべてのメニュー項目を一つのメニューに統合
   */
  createMenu: function() {
    try {
      // UIが利用可能かチェック
      var ui = SpreadsheetApp.getUi();
      var menu = ui.createMenu('ポートフォリオ管理');
      
      // 基本機能
      menu.addItem('情報更新パネル', 'showSidebar')
          .addSeparator()
          .addItem('メイン情報更新', 'updateMainInfo')
          .addItem('保有株リスト更新', 'updatePF')
          .addItem('トレンド情報更新', 'updateTrendInfo')
          .addItem('トレンド情報更新（最適化）', 'updateTrendInfoOptimized')
          .addItem('トレンド情報更新（継続）', 'resumeUpdateTrendInfo')
          .addItem('_Reference情報更新', 'updateReferenceInfo')
          .addItem('選択行の_Reference情報更新', 'updateSelectedReferenceInfo')
          .addItem('特定銘柄の_Reference情報を更新', 'showCodeInputDialog')
          .addItem('ダッシュボード表示', 'showDashboard')
          .addItem('チャートURL更新', 'getImageUrlChart')
          .addSeparator();
      
      // キャッシュ管理サブメニュー
      var cacheMenu = ui.createMenu('キャッシュ管理');
      cacheMenu.addItem('すべてのキャッシュをクリア', 'clearAllCache')
               .addItem('株価キャッシュをクリア', 'clearStockCache')
               .addItem('日次更新項目キャッシュをクリア', 'clearDailyCache')
               .addItem('列インデックスキャッシュをクリア', 'clearColumnsCache');
      
      menu.addSubMenu(cacheMenu)
          .addSeparator()
          .addItem('処理をキャンセル', 'cancelCurrentProcess')
          .addSeparator();
      
      // テストツールサブメニュー
      var testMenu = ui.createMenu('テストツール');
      testMenu.addItem('単一銘柄株価取得テスト', 'testStockPrice')
              .addItem('複数銘柄株価バッチテスト', 'batchTestStockPrices')
              .addItem('代替ソースから株価取得', 'getStockPriceFromAlternative')
              .addItem('updatePF機能テスト', 'testUpdatePF');
      
      menu.addSubMenu(testMenu);
      
      menu.addToUi();
      
      if (typeof LoggerManager !== 'undefined') {
        LoggerManager.debug('統合メニューを作成しました');
      } else {
        console.log('統合メニューを作成しました');
      }
    } catch (e) {
      // UIが利用できない場合や他のエラーが発生した場合
      if (typeof LoggerManager !== 'undefined') {
        LoggerManager.error('メニュー作成中にエラーが発生しました', e);
      } else {
        console.error('メニュー作成中にエラーが発生しました: ' + e.message);
      }
    }
  },

  /**
   * onOpen時に呼び出される関数
   * 統合メニューのみを作成する
   */
  onOpen: function() {
    this.createMenu();
  },

  updateProgress: function(current, total, status, details = {}) {
    try {
      var progress = Math.min(100, Math.max(0, Math.round((current / total) * 100)));
      
      // 詳細情報をマージ
      var progressDetails = {
        itemsProcessed: current,
        totalItems: total,
        remainingItems: Math.max(0, total - current),
        estimatedTimeRemaining: this.calculateEstimatedTimeRemaining(current, total),
        ...details
      };
      
      var data = {
        progress: progress,
        status: status || '処理中...',
        current: current,
        total: total,
        lastUpdated: new Date().toISOString(),
        details: progressDetails
      };
      
      var previousData = this.getLatestProgress();
      if (previousData && previousData.current !== current) {
        this.logProgress(data);
      } else if (!previousData) {
        this.logProgress(data);
      }
      
      PropertiesService.getScriptProperties().setProperty(
        'latestProgress',
        JSON.stringify(data)
      );
      
      // ログにも記録（重要な進捗ポイントのみ）
      if (current === 0 || current === total || current % Math.max(1, Math.floor(total / 10)) === 0) {
        LoggerManager.info(`進捗状況: ${progress}% (${current}/${total}) - ${status}`);
      }
    } catch (e) {
      LoggerManager.error('Progress update failed', e);
    }
  },

  /**
   * 残り時間を概算する
   * @param {number} current 現在の進捗
   * @param {number} total 合計
   * @returns {string} 残り時間の概算（分:秒）
   */
  calculateEstimatedTimeRemaining: function(current, total) {
    try {
      if (current <= 0 || current >= total) {
        return '00:00';
      }
      
      var previousData = this.getLatestProgress();
      if (!previousData || !previousData.lastUpdated) {
        return '計算中...';
      }
      
      var now = new Date();
      var lastUpdate = new Date(previousData.lastUpdated);
      var elapsedMs = now.getTime() - lastUpdate.getTime();
      
      if (elapsedMs <= 0 || previousData.current >= current) {
        return '計算中...';
      }
      
      var itemsProcessedSinceLastUpdate = current - previousData.current;
      var msPerItem = elapsedMs / itemsProcessedSinceLastUpdate;
      var remainingItems = total - current;
      var estimatedRemainingMs = msPerItem * remainingItems;
      
      var minutes = Math.floor(estimatedRemainingMs / 60000);
      var seconds = Math.floor((estimatedRemainingMs % 60000) / 1000);
      
      return (minutes < 10 ? '0' : '') + minutes + ':' + (seconds < 10 ? '0' : '') + seconds;
    } catch (e) {
      LoggerManager.error('Error calculating estimated time', e);
      return '計算中...';
    }
  },

  /**
   * 進捗状況を表示する
   * @param {number} current 現在の進捗
   * @param {number} total 合計
   * @param {string} message メッセージ
   * @param {string} processType 処理タイプ（'main'または'pf'）
   * @param {Object} details 詳細情報（オプション）
   */
  showProgress: function(current, total, message, processType = 'main', details = {}) {
    // 詳細情報を追加
    const enhancedDetails = {
      processType: processType,
      ...details
    };
    
    // 進捗を更新
    this.updateProgress(current, total, message, enhancedDetails);
  },
  
  /**
   * メイン情報更新の進捗を表示する
   * @param {number} current 現在の進捗
   * @param {number} total 合計
   * @param {string} message メッセージ
   * @param {Object} details 詳細情報（オプション）
   */
  showMainProgress: function(current, total, message, details = {}) {
    this.showProgress(current, total, message, 'main', details);
  },
  
  /**
   * ポートフォリオ更新の進捗を表示する
   * @param {number} current 現在の進捗
   * @param {number} total 合計
   * @param {string} message メッセージ
   * @param {Object} details 詳細情報（オプション）
   */
  showPFProgress: function(current, total, message, details = {}) {
    this.showProgress(current, total, message, 'pf', details);
  },
  
  /**
   * エラーメッセージを表示する
   * @param {string} message エラーメッセージ
   */
  showError: function(message) {
    this.showProgress(0, 0, 'エラーが発生しました: ' + message);
    LoggerManager.error('UI Error:', message);
  },
  
  /**
   * 完了メッセージを表示する
   * @param {string} message 完了メッセージ
   * @param {string} processType 処理タイプ
   */
  showCompletion: function(message, processType = 'main') {
    const fullMessage = `${message}が完了しました`;
    if (processType === 'main') {
      this.showMainProgress(100, 100, fullMessage);
    } else {
      this.showPFProgress(100, 100, fullMessage);
    }
  },

  logProgress: function(data) {
    try {
      const sheet = this.getProgressLogSheet();
      sheet.appendRow([
        new Date(),
        data.progress,
        data.status,
        data.current,
        data.total,
        Session.getActiveUser().getEmail()
      ]);
    } catch (e) {
      LoggerManager.error('Progress logging failed', e);
    }
  },

  getProgressLogSheet: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('ProgressLog');
    
    if (!sheet) {
      sheet = ss.insertSheet('ProgressLog');
      sheet.getRange('A1:F1').setValues([[
        'Timestamp',
        'Progress (%)',
        'Status',
        'Current',
        'Total',
        'User'
      ]]);
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  },

  /**
   * 最新の進捗状況を取得する
   * @returns {Object} 進捗状況データ
   */
  getLatestProgress: function() {
    try {
      var progressString = PropertiesService.getScriptProperties().getProperty('latestProgress');
      if (progressString) {
        return JSON.parse(progressString);
      } else {
        return {
          progress: 0,
          status: '準備中...',
          current: 0,
          total: 0,
          lastUpdated: new Date().toISOString(),
          details: {
            itemsProcessed: 0,
            totalItems: 0,
            remainingItems: 0,
            estimatedTimeRemaining: '00:00'
          }
        };
      }
    } catch (e) {
      LoggerManager.error('Error getting latest progress', e);
      return null;
    }
  },

  clearProgress: function() {
    try {
      PropertiesService.getScriptProperties().deleteProperty('latestProgress');
      LoggerManager.debug('Progress cleared');
    } catch (e) {
      LoggerManager.error('Error clearing progress', e);
    }
  },

  showToast: function(message, title, timeout) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        message,
        title || 'Info',
        timeout || 5
      );
      LoggerManager.debug('Toast showed', {
        message: message,
        title: title,
        timeout: timeout
      });
    } catch (e) {
      LoggerManager.error('Error showing toast', e);
    }
  },

  confirmAction: function(message) {
    try {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        '確認',
        message,
        ui.ButtonSet.YES_NO
      );
      return response === ui.Button.YES;
    } catch (e) {
      LoggerManager.error('Error showing confirm dialog', e);
      return false;
    }
  },

  validateActiveSheet: function() {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (!sheet) {
      throw new Error('アクティブなシートが見つかりません');
    }
    return sheet;
  },

  getSpreadsheetUrl: function() {
    try {
      return SpreadsheetApp.getActiveSpreadsheet().getUrl();
    } catch (e) {
      LoggerManager.error('Error getting spreadsheet URL', e);
      return null;
    }
  },

  refreshSheet: function() {
    try {
      SpreadsheetApp.flush();
    } catch (e) {
      LoggerManager.error('Error refreshing sheet', e);
    }
  },

  /**
   * 銘柄コード入力ダイアログを表示する
   */
  showCodeInputDialog: function() {
    const ui = SpreadsheetApp.getUi();
    const result = ui.prompt(
      '特定銘柄の_Reference情報を更新',
      '更新したい銘柄コードをカンマ区切りで入力してください（例: 1234,5678,9012）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    // OKボタンがクリックされた場合
    if (result.getSelectedButton() === ui.Button.OK) {
      const codes = result.getResponseText();
      if (codes && codes.trim() !== '') {
        try {
          updateReferenceInfoForCodes(codes);
        } catch (e) {
          ui.alert('エラー', e.message, ui.ButtonSet.OK);
        }
      } else {
        ui.alert('エラー', '銘柄コードが入力されていません', ui.ButtonSet.OK);
      }
    }
  }
};

// Global functions
function onOpen() {
  try {
    // MainController.jsのonOpen関数と競合しないように修正
    UIManager.createMenu();
    console.log('メニューを作成しました');
  } catch (e) {
    console.error('onOpen関数でエラーが発生しました: ' + e.message);
  }
}

function showSidebar() {
  return UIManager.showSidebar();
}

function updateSidebarProgress(data) {
  LoggerManager.debug('Sidebar progress update', data);
}

function getLatestProgressFromProperties() {
  return UIManager.getLatestProgress();
}

function clearProgress() {
  return UIManager.clearProgress();
}