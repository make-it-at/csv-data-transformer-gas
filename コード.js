/**
 * CSV処理・データ変換ツール - メイン処理
 * 
 * このファイルはCSV処理ツールのメインエントリーポイントです。
 * HTMLサイドバーの表示とメイン処理の制御を行います。
 */

// 定数定義
const MAX_EXECUTION_TIME = 5 * 60 * 1000; // 5分
const BATCH_SIZE = 1000; // バッチ処理サイズ
const SHEET_NAMES = {
  LPCSV: 'LPcsv',
  PPFORMAT: 'PPformat',
  MFFORMAT: 'MFformat',
  MFSETTINGS: 'MF設定',
  LOG: 'ログ'
};

/**
 * スプレッドシート開時に実行される関数
 * カスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('CSV処理ツール')
    .addItem('サイドバーを表示', 'showSidebar')
    .addItem('必要シートを作成', 'initializeSheets')
    .addToUi();
}

/**
 * HTMLサイドバーを表示する
 * スプレッドシートから直接実行するメイン関数
 */
function showSidebar() {
  try {
    Logger.log('[main.gs] サイドバー表示開始');
    
    // スプレッドシートの準備確認
    initializeSheets();
    
    // HTMLテンプレートを読み込み
    const htmlOutput = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('CSV処理ツール v1.0.0')
      .setWidth(320);
    
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
    
    Logger.log('[main.gs] サイドバー表示完了');
    
  } catch (error) {
    Logger.log(`[main.gs] エラー: ${error.message}`);
    showErrorDialog('サイドバー表示エラー', error.message);
  }
}

/**
 * 必要なシートの存在確認・作成
 */
function initializeSheets() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheets = spreadsheet.getSheets().map(sheet => sheet.getName());
    
    // 必要なシートの確認・作成
    Object.values(SHEET_NAMES).forEach(sheetName => {
      if (!existingSheets.includes(sheetName)) {
        const newSheet = spreadsheet.insertSheet(sheetName);
        Logger.log(`[main.gs] シート作成: ${sheetName}`);
        
        // ログシートの場合はヘッダーを設定
        if (sheetName === SHEET_NAMES.LOG) {
          initializeLogSheet(newSheet);
        }
        // MF設定シートの場合は初期設定を設定
        if (sheetName === SHEET_NAMES.MFSETTINGS) {
          initializeMFSettingsSheet(newSheet);
        }
      }
    });
    
    Logger.log('[main.gs] シート初期化完了');
    
  } catch (error) {
    Logger.log(`[main.gs] シート初期化エラー: ${error.message}`);
    throw error;
  }
}

/**
 * MF設定シートの初期化
 * @param {Sheet} settingsSheet - MF設定シート
 */
function initializeMFSettingsSheet(settingsSheet) {
  // 設定項目の定義
  const settings = [
    ['設定項目', '種別', '設定値', '説明'],
    ['獲得_借方勘定科目', '獲得', '事業主貸', '獲得時の借方勘定科目'],
    ['獲得_貸方勘定科目', '獲得', '売掛金', '獲得時の貸方勘定科目'],
    ['利用_借方勘定科目', '利用', '外注費', '利用時の借方勘定科目'],
    ['利用_貸方勘定科目', '利用', '事業主借', '利用時の貸方勘定科目'],
    ['獲得_税区分', '獲得', '対象外', '獲得時の税区分'],
    ['利用_税区分', '利用', '対象外', '利用時の税区分'],
    ['仕訳メモを含める', '共通', 'true', '仕訳メモを含めるかどうか（true/false）']
  ];
  
  // データを書き込み
  settingsSheet.getRange(1, 1, settings.length, 4).setValues(settings);
  
  // ヘッダー行のフォーマット
  const headerRange = settingsSheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8f4fd');
  
  // 列幅の調整
  settingsSheet.setColumnWidth(1, 200); // 設定項目
  settingsSheet.setColumnWidth(2, 100); // 種別
  settingsSheet.setColumnWidth(3, 150); // 設定値
  settingsSheet.setColumnWidth(4, 300); // 説明
  
  Logger.log('[main.gs] MF設定シート初期化完了');
}

/**
 * ログシートの初期化
 * @param {Sheet} logSheet - ログシート
 */
function initializeLogSheet(logSheet) {
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
}

/**
 * CSVファイルをアップロードして処理する
 * HTMLサイドバーから呼び出される関数
 * 
 * @param {string} csvContent - CSVファイルの内容
 * @param {string} fileName - ファイル名
 * @param {Object} options - 処理オプション
 * @return {Object} 処理結果
 */
function processCSVImport(csvContent, fileName, options = {}) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log(`[main.gs] CSV処理開始: ${fileName}`);
    writeLog('INFO', 'CSVインポート', `ファイル: ${fileName}`, { options });
    
    // デフォルトオプションの設定
    const defaultOptions = {
      hasHeader: true,
      encoding: 'UTF-8',
      delimiter: ',',
      maxRows: 10000
    };
    
    const finalOptions = { ...defaultOptions, ...options };
    
    // CSVデータの解析
    const csvData = parseCSVContent(csvContent, finalOptions);
    
    // データの検証
    validateCSVData(csvData, finalOptions);
    
    // LPcsvシートにデータを保存
    const lpCsvSheet = getSheetSafely(SHEET_NAMES.LPCSV);
    importToLPcsv(csvData, lpCsvSheet);
    
    // 処理結果を返す
    const result = {
      success: true,
      message: 'CSVファイルのインポートが完了しました',
      data: {
        fileName: fileName,
        totalRows: csvData.length,
        totalColumns: csvData.length > 0 ? csvData[0].length : 0,
        processingTime: new Date().getTime() - startTime,
        sheetName: SHEET_NAMES.LPCSV
      }
    };
    
    Logger.log(`[main.gs] CSV処理完了: ${result.data.totalRows}行`);
    writeLog('INFO', 'CSVインポート完了', `${result.data.totalRows}行処理`, result.data);
    
    return result;
    
  } catch (error) {
    Logger.log(`[main.gs] CSV処理エラー: ${error.message}`);
    writeLog('ERROR', 'CSVインポートエラー', error.message, { fileName, error: error.stack });
    
    return {
      success: false,
      message: `処理エラー: ${error.message}`,
      error: error.message
    };
  }
}

/**
 * エラーダイアログを表示する
 * 
 * @param {string} title - ダイアログタイトル
 * @param {string} message - エラーメッセージ
 */
function showErrorDialog(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * 現在のシート状況を取得する
 * HTMLサイドバーから呼び出される関数
 * 
 * @return {Object} シート状況
 */
function getSheetStatus() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const sheetStatus = {};
    
    Object.values(SHEET_NAMES).forEach(sheetName => {
      const sheet = sheets.find(s => s.getName() === sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        const lastColumn = sheet.getLastColumn();
        sheetStatus[sheetName] = {
          exists: true,
          rows: lastRow,
          columns: lastColumn,
          hasData: lastRow > 1 // ヘッダー行以外にデータがあるか
        };
      } else {
        sheetStatus[sheetName] = {
          exists: false,
          rows: 0,
          columns: 0,
          hasData: false
        };
      }
    });
    
    return {
      success: true,
      data: sheetStatus
    };
    
  } catch (error) {
    Logger.log(`[main.gs] シート状況取得エラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * データ転記処理の実行
 * LPcsvのデータをPPformat・MFformatに転記
 * 
 * @param {Object} settings - 転記設定（オプション）
 * @return {Object} 処理結果
 */
function executeDataTransfer(settings = null) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log('[main.gs] データ転記実行開始');
    
    // 必要なシートが存在するかチェック
    initializeSheets();
    
    // LPcsvにデータがあるかチェック
    const lpCsvSheet = getSheetSafely(SHEET_NAMES.LPCSV);
    const lastRow = lpCsvSheet.getLastRow();
    
    if (lastRow <= 1) {
      throw new Error('LPcsvシートにデータがありません。まずCSVをインポートしてください。');
    }
    
    // 設定シートから設定を読み込み（settingsがnullの場合）
    let finalSettings = settings;
    if (!finalSettings) {
      Logger.log('[main.gs] 設定シートから設定を読み込み');
      const mfSettings = loadMFSettingsFromSheet();
      finalSettings = {
        ppformat: TRANSFER_CONFIG.DEFAULT_SETTINGS.ppformat,
        mfformat: mfSettings
      };
      Logger.log(`[main.gs] 読み込んだ設定: ${JSON.stringify(finalSettings)}`);
    }
    
    // データ転記実行
    const result = transferDataFromLPcsv(finalSettings);
    
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[main.gs] データ転記完了: ${processingTime}ms`);
    
    // 成功メッセージをログに記録
    writeLog('INFO', 'データ転記', `転記処理完了: ${result.totalProcessed}行処理`, {
      ppformat: result.ppformat,
      mfformat: result.mfformat,
      processingTime: processingTime
    });
    
    return {
      success: true,
      message: `データ転記が完了しました。PPformat: ${result.ppformat?.processedRows || 0}行, MFformat: ${result.mfformat?.processedRows || 0}行`,
      result: result,
      processingTime: processingTime
    };
    
  } catch (error) {
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[main.gs] データ転記エラー: ${error.message}`);
    
    writeLog('ERROR', 'データ転記エラー', error.message, {
      processingTime: processingTime,
      stack: error.stack
    });
    
    showErrorDialog('データ転記エラー', error.message);
    
    return {
      success: false,
      message: `データ転記エラー: ${error.message}`,
      error: error.message,
      processingTime: processingTime
    };
  }
}

/**
 * PPformat CSVエクスポート処理の実行
 * 
 * @param {Object} options - エクスポートオプション
 * @return {Object} 処理結果
 */
function executePPformatExport(options = {}) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log('[main.gs] PPformat CSVエクスポート実行開始');
    
    // PPformatにデータがあるかチェック
    const ppFormatSheet = getSheetSafely(SHEET_NAMES.PPFORMAT);
    const lastRow = ppFormatSheet.getLastRow();
    
    if (lastRow <= 1) {
      throw new Error('PPformatシートにデータがありません。まずデータ転記を実行してください。');
    }
    
    // CSVエクスポート実行
    const result = exportPPformatToCSV(options);
    
    // データURIを使用した直接ダウンロード用URL生成
    const dataUri = createDataUri(result.csvContent, result.fileName);
    
    // Google Drive経由のダウンロードURL生成（フォールバック用）
    const downloadInfo = createDownloadUrl(result.blob, result.fileName);
    
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[main.gs] PPformat CSVエクスポート完了: ${processingTime}ms`);
    
    return {
      success: true,
      message: `PPformat CSVエクスポートが完了しました。${result.rowCount}行を出力しました。`,
      downloadUrl: dataUri, // データURIを優先
      fallbackUrl: downloadInfo.downloadUrl, // フォールバック用
      fileName: result.fileName,
      fileId: downloadInfo.fileId,
      rowCount: result.rowCount,
      processingTime: processingTime
    };
    
  } catch (error) {
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[main.gs] PPformat CSVエクスポートエラー: ${error.message}`);
    
    showErrorDialog('PPformat CSVエクスポートエラー', error.message);
    
    return {
      success: false,
      message: `PPformat CSVエクスポートエラー: ${error.message}`,
      error: error.message,
      processingTime: processingTime
    };
  }
}

/**
 * MFformat CSVエクスポート処理の実行
 * 
 * @param {Object} options - エクスポートオプション
 * @return {Object} 処理結果
 */
function executeMFformatExport(options = {}) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log('[main.gs] MFformat CSVエクスポート実行開始');
    
    // MFformatにデータがあるかチェック
    const mfFormatSheet = getSheetSafely(SHEET_NAMES.MFFORMAT);
    const lastRow = mfFormatSheet.getLastRow();
    
    if (lastRow <= 1) {
      throw new Error('MFformatシートにデータがありません。まずデータ転記を実行してください。');
    }
    
    // CSVエクスポート実行
    const result = exportMFformatToCSV(options);
    
    // データURIを使用した直接ダウンロード用URL生成
    const dataUri = createDataUri(result.csvContent, result.fileName);
    
    // Google Drive経由のダウンロードURL生成（フォールバック用）
    const downloadInfo = createDownloadUrl(result.blob, result.fileName);
    
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[main.gs] MFformat CSVエクスポート完了: ${processingTime}ms`);
    
    return {
      success: true,
      message: `MFformat CSVエクスポートが完了しました。${result.rowCount}行を出力しました。`,
      downloadUrl: dataUri, // データURIを優先
      fallbackUrl: downloadInfo.downloadUrl, // フォールバック用
      fileName: result.fileName,
      fileId: downloadInfo.fileId,
      rowCount: result.rowCount,
      processingTime: processingTime
    };
    
  } catch (error) {
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[main.gs] MFformat CSVエクスポートエラー: ${error.message}`);
    
    showErrorDialog('MFformat CSVエクスポートエラー', error.message);
    
    return {
      success: false,
      message: `MFformat CSVエクスポートエラー: ${error.message}`,
      error: error.message,
      processingTime: processingTime
    };
  }
}

/**
 * データURIを使用した直接ダウンロード用URL生成
 * 
 * @param {string} csvContent - CSVコンテンツ
 * @param {string} fileName - ファイル名
 * @return {string} データURI
 */
function createDataUri(csvContent, fileName) {
  try {
    // UTF-8 BOMを追加
    const csvWithBom = '\uFEFF' + csvContent;
    
    // Base64エンコード
    const base64Content = Utilities.base64Encode(csvWithBom, Utilities.Charset.UTF_8);
    
    // データURIの生成
    const dataUri = `data:text/csv;charset=utf-8;base64,${base64Content}`;
    
    Logger.log(`[main.gs] データURI生成完了: ${fileName}`);
    
    return dataUri;
    
  } catch (error) {
    Logger.log(`[main.gs] データURI生成エラー: ${error.message}`);
    
    // フォールバック: シンプルなデータURI
    const encodedContent = encodeURIComponent('\uFEFF' + csvContent);
    return `data:text/csv;charset=utf-8,${encodedContent}`;
  }
}

/**
 * MF設定シートから設定を読み込む
 * @return {Object} 設定オブジェクト
 */
function loadMFSettingsFromSheet() {
  try {
    const settingsSheet = getSheetSafely(SHEET_NAMES.MFSETTINGS);
    if (!settingsSheet) {
      Logger.log('[main.gs] MF設定シートが見つかりません');
      return TRANSFER_CONFIG.DEFAULT_SETTINGS.mfformat;
    }
    
    const data = settingsSheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('[main.gs] MF設定シートにデータがありません');
      return TRANSFER_CONFIG.DEFAULT_SETTINGS.mfformat;
    }
    
    // 設定値をマップに変換
    const settingsMap = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 3) {
        const key = row[0]; // 設定項目
        const value = row[2]; // 設定値
        settingsMap[key] = value;
      }
    }
    
    // 設定オブジェクトを構築
    const settings = {
      enabled: true,
      filters: {
        includeTypes: ['獲得', '利用'],
        excludeTypes: [],
        minAmount: null,
        maxAmount: null
      },
      transformations: {
        dateFormat: 'YYYY/MM/DD',
        accounts: {
          獲得: {
            debit: settingsMap['獲得_借方勘定科目'] || '事業主貸',
            credit: settingsMap['獲得_貸方勘定科目'] || '売掛金'
          },
          利用: {
            debit: settingsMap['利用_借方勘定科目'] || '外注費',
            credit: settingsMap['利用_貸方勘定科目'] || '事業主借'
          }
        },
        taxCategories: {
          獲得: settingsMap['獲得_税区分'] || '対象外',
          利用: settingsMap['利用_税区分'] || '対象外'
        },
        includeMemo: settingsMap['仕訳メモを含める'] === 'true'
      }
    };
    
    Logger.log(`[main.gs] MF設定読み込み完了: ${JSON.stringify(settings)}`);
    return settings;
    
  } catch (error) {
    Logger.log(`[main.gs] MF設定読み込みエラー: ${error.message}`);
    return TRANSFER_CONFIG.DEFAULT_SETTINGS.mfformat;
  }
}

/**
 * MFformat設定の取得（設定シートから読み込み）
 * @return {Object} MFformat設定
 */
function getMFformatSettings() {
  try {
    Logger.log('[main.gs] MFformat設定取得開始');
    
    // 設定シートから設定を読み込み
    const settings = loadMFSettingsFromSheet();
    
    return {
      success: true,
      settings: settings
    };
    
  } catch (error) {
    Logger.log(`[main.gs] MFformat設定取得エラー: ${error.message}`);
    
    // エラー時はデフォルト設定を返す
    return {
      success: true,
      settings: TRANSFER_CONFIG.DEFAULT_SETTINGS.mfformat
    };
  }
}

/**
 * MF設定シートに設定を書き込む
 * @param {Object} settings - 設定オブジェクト
 */
function saveMFSettingsToSheet(settings) {
  try {
    const settingsSheet = getSheetSafely(SHEET_NAMES.MFSETTINGS);
    if (!settingsSheet) {
      throw new Error('MF設定シートが見つかりません');
    }
    
    // 設定値を更新
    const data = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.length >= 3) {
        const key = row[0]; // 設定項目
        let newValue = row[2]; // 現在の設定値
        
        // 設定項目に応じて値を更新
        switch (key) {
          case '獲得_借方勘定科目':
            newValue = settings.transformations?.accounts?.獲得?.debit || '事業主貸';
            break;
          case '獲得_貸方勘定科目':
            newValue = settings.transformations?.accounts?.獲得?.credit || '売掛金';
            break;
          case '利用_借方勘定科目':
            newValue = settings.transformations?.accounts?.利用?.debit || '外注費';
            break;
          case '利用_貸方勘定科目':
            newValue = settings.transformations?.accounts?.利用?.credit || '事業主借';
            break;
          case '獲得_税区分':
            newValue = settings.transformations?.taxCategories?.獲得 || '対象外';
            break;
          case '利用_税区分':
            newValue = settings.transformations?.taxCategories?.利用 || '対象外';
            break;
          case '仕訳メモを含める':
            newValue = settings.transformations?.includeMemo ? 'true' : 'false';
            break;
        }
        
        // 設定値を更新
        settingsSheet.getRange(i + 1, 3).setValue(newValue);
      }
    }
    
    Logger.log('[main.gs] MF設定シート更新完了');
    
  } catch (error) {
    Logger.log(`[main.gs] MF設定シート更新エラー: ${error.message}`);
    throw error;
  }
}

/**
 * MF設定シートを開く
 */
function openMFSettingsSheet() {
  try {
    Logger.log('[main.gs] MF設定シートを開く');
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = spreadsheet.getSheetByName(SHEET_NAMES.MFSETTINGS);
    
    if (settingsSheet) {
      // 設定シートをアクティブにする
      spreadsheet.setActiveSheet(settingsSheet);
      Logger.log('[main.gs] MF設定シートを開きました');
    } else {
      // 設定シートが存在しない場合は作成
      initializeSheets();
      Logger.log('[main.gs] MF設定シートを作成しました');
    }
    
  } catch (error) {
    Logger.log(`[main.gs] MF設定シートを開くエラー: ${error.message}`);
    throw error;
  }
}

/**
 * MFformat設定の更新（設定シートに書き込み）
 * 
 * @param {Object} newSettings - 新しい設定
 * @return {Object} 更新結果
 */
function updateMFformatSettings(newSettings) {
  try {
    Logger.log('[main.gs] MFformat設定更新開始');
    
    // 設定の検証
    if (!newSettings || typeof newSettings !== 'object') {
      throw new Error('無効な設定データです');
    }
    
    // 設定をマージ（深いコピー）
    const currentSettings = JSON.parse(JSON.stringify(TRANSFER_CONFIG.DEFAULT_SETTINGS.mfformat));
    
    if (newSettings.transformations) {
      if (newSettings.transformations.accounts) {
        currentSettings.transformations.accounts = Object.assign(
          {},
          currentSettings.transformations.accounts,
          newSettings.transformations.accounts
        );
      }
      
      if (newSettings.transformations.taxCategories) {
        currentSettings.transformations.taxCategories = Object.assign(
          {},
          currentSettings.transformations.taxCategories,
          newSettings.transformations.taxCategories
        );
      }
      
      if (typeof newSettings.transformations.includeMemo === 'boolean') {
        currentSettings.transformations.includeMemo = newSettings.transformations.includeMemo;
      }
    }
    
    // 設定シートに書き込み
    saveMFSettingsToSheet(currentSettings);
    
    // ログ記録
    writeLog('INFO', 'MFformat設定更新完了', JSON.stringify(currentSettings), {
      newSettings: currentSettings
    });
    
    Logger.log('[main.gs] MFformat設定更新完了');
    
    return {
      success: true,
      settings: currentSettings
    };
    
  } catch (error) {
    Logger.log(`[main.gs] MFformat設定更新エラー: ${error.message}`);
    
    writeLog('ERROR', 'MFformat設定更新エラー', error.message, {
      newSettings: newSettings
    });
    
    return {
      success: false,
      error: error.message
    };
  }
}
