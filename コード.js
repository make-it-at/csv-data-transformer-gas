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
      }
    });
    
    Logger.log('[main.gs] シート初期化完了');
    
  } catch (error) {
    Logger.log(`[main.gs] シート初期化エラー: ${error.message}`);
    throw error;
  }
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
    
    // データ転記実行
    const result = transferDataFromLPcsv(settings);
    
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
