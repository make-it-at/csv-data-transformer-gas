// versionManager.js
// Version 1.0
// システムバージョン管理とメタデータ管理

var VersionManager = {
  
  // システムバージョン情報
  VERSION_INFO: {
    version: '1.2.1',
    buildDate: '2025-08-09T14:00:00Z',
    releaseNotes: [
      'v1.2.1: 分割処理版エラーを修正し従来版を使用するように変更',
      'v1.2.0: _Reference情報更新の処理再開機能を修正',
      'v1.1.0: 分割処理によるバッチ処理システムを実装', 
      'v1.0.0: 初期リリース'
    ]
  },

  /**
   * 現在のバージョン情報を取得
   * @returns {Object} バージョン情報オブジェクト
   */
  getVersionInfo: function() {
    return {
      version: this.VERSION_INFO.version,
      buildDate: this.VERSION_INFO.buildDate,
      formattedBuildDate: this.formatBuildDate(this.VERSION_INFO.buildDate),
      releaseNotes: this.VERSION_INFO.releaseNotes
    };
  },

  /**
   * バージョン番号のみ取得
   * @returns {string} バージョン番号
   */
  getVersion: function() {
    return this.VERSION_INFO.version;
  },

  /**
   * ビルド日時をフォーマット
   * @param {string} isoString ISO形式の日時文字列
   * @returns {string} フォーマット済み日時
   */
  formatBuildDate: function(isoString) {
    try {
      const date = new Date(isoString);
      return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    } catch (error) {
      LoggerManager.warn('日時フォーマットエラー:', error);
      return isoString;
    }
  },

  /**
   * バージョン情報をログに出力
   */
  logVersionInfo: function() {
    const info = this.getVersionInfo();
    LoggerManager.info('システムバージョン情報', {
      version: info.version,
      buildDate: info.formattedBuildDate,
      releaseNotesCount: info.releaseNotes.length
    });
  },

  /**
   * バージョンを更新（開発用）
   * @param {string} newVersion 新しいバージョン番号
   * @param {string} releaseNote リリースノート
   */
  updateVersion: function(newVersion, releaseNote = '') {
    this.VERSION_INFO.version = newVersion;
    this.VERSION_INFO.buildDate = new Date().toISOString();
    
    if (releaseNote) {
      this.VERSION_INFO.releaseNotes.unshift(`v${newVersion}: ${releaseNote}`);
      // 最大10件のリリースノートを保持
      if (this.VERSION_INFO.releaseNotes.length > 10) {
        this.VERSION_INFO.releaseNotes = this.VERSION_INFO.releaseNotes.slice(0, 10);
      }
    }
    
    LoggerManager.info(`バージョンを更新しました: v${newVersion}`, {
      newVersion: newVersion,
      releaseNote: releaseNote,
      buildDate: this.VERSION_INFO.buildDate
    });
  },

  /**
   * システム情報の詳細を取得
   * @returns {Object} システム情報
   */
  getSystemInfo: function() {
    const versionInfo = this.getVersionInfo();
    const scriptProperties = PropertiesService.getScriptProperties();
    
    return {
      version: versionInfo.version,
      buildDate: versionInfo.formattedBuildDate,
      scriptId: ScriptApp.newTrigger('dummy').getHandlerFunction() ? 'Available' : 'Not Available',
      timeZone: Session.getScriptTimeZone(),
      userEmail: Session.getActiveUser().getEmail(),
      lastUpdate: scriptProperties.getProperty('lastSystemUpdate') || 'Unknown',
      releaseNotes: versionInfo.releaseNotes
    };
  },

  /**
   * システム情報を保存
   */
  saveSystemInfo: function() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const systemInfo = {
        version: this.VERSION_INFO.version,
        buildDate: this.VERSION_INFO.buildDate,
        lastUpdate: new Date().toISOString()
      };
      
      scriptProperties.setProperty('systemInfo', JSON.stringify(systemInfo));
      scriptProperties.setProperty('lastSystemUpdate', systemInfo.lastUpdate);
      
      LoggerManager.debug('システム情報を保存しました', systemInfo);
    } catch (error) {
      LoggerManager.error('システム情報保存エラー:', error);
    }
  },

  /**
   * システム情報を復元
   * @returns {Object|null} 保存されたシステム情報
   */
  loadSystemInfo: function() {
    try {
      const scriptProperties = PropertiesService.getScriptProperties();
      const systemInfoJson = scriptProperties.getProperty('systemInfo');
      
      if (systemInfoJson) {
        return JSON.parse(systemInfoJson);
      }
      return null;
    } catch (error) {
      LoggerManager.error('システム情報読み込みエラー:', error);
      return null;
    }
  }
};

/**
 * バージョン情報取得（外部呼び出し用）
 * @returns {Object} バージョン情報
 */
function getVersionInfo() {
  return VersionManager.getVersionInfo();
}

/**
 * バージョン番号取得（外部呼び出し用）  
 * @returns {string} バージョン番号
 */
function getSystemVersion() {
  return VersionManager.getVersion();
}

/**
 * システム情報取得（外部呼び出し用）
 * @returns {Object} システム情報
 */
function getSystemInfo() {
  return VersionManager.getSystemInfo();
}