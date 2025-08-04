/**
 * CSVエクスポート機能
 * PPformat・MFformatシートからCSVファイルを生成・ダウンロード
 */

// 定数定義
const EXPORT_CONFIG = {
  // エクスポート設定
  DEFAULT_OPTIONS: {
    encoding: 'UTF-8',
    delimiter: ',',
    includeHeader: true,
    dateFormat: 'YYYY/MM/DD',
    numberFormat: 'default' // default, comma, plain
  },
  
  // ファイル名設定
  FILENAME_TEMPLATES: {
    ppformat: 'PPformat_export_{timestamp}.csv',
    mfformat: 'MFformat_export_{timestamp}.csv'
  },
  
  // 最大エクスポート行数
  MAX_EXPORT_ROWS: 50000
};

/**
 * PPformatシートからCSVエクスポート
 * 
 * @param {Object} options - エクスポートオプション
 * @return {Object} エクスポート結果
 */
function exportPPformatToCSV(options = {}) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log(`[csvExporter.gs] PPformatエクスポート開始`);
    writeLog('INFO', 'CSVエクスポート', 'PPformatエクスポートを開始します');
    
    // オプションの準備
    const exportOptions = { ...EXPORT_CONFIG.DEFAULT_OPTIONS, ...options };
    
    // PPformatシートの取得
    const sheet = getSheetSafely(SHEET_NAMES.PPFORMAT);
    const data = getSheetDataForExport(sheet);
    
    if (!data || data.length === 0) {
      throw new Error('PPformatシートにデータがありません');
    }
    
    // CSVデータの生成
    const csvContent = generateCSVContent(data, exportOptions);
    
    // ファイル名の生成
    const fileName = generateFileName('ppformat', exportOptions);
    
    // Blobオブジェクトの作成
    const blob = createCSVBlob(csvContent, exportOptions.encoding);
    
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[csvExporter.gs] PPformatエクスポート完了: ${data.length}行, ${processingTime}ms`);
    
    // ログ記録
    writeLog('INFO', 'CSVエクスポート', `PPformatエクスポート完了: ${data.length}行`, {
      fileName: fileName,
      processingTime: processingTime,
      options: exportOptions
    });
    
    return {
      success: true,
      fileName: fileName,
      blob: blob,
      rowCount: data.length,
      processingTime: processingTime
    };
    
  } catch (error) {
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[csvExporter.gs] PPformatエクスポートエラー: ${error.message}`);
    
    writeLog('ERROR', 'CSVエクスポートエラー', `PPformat: ${error.message}`, {
      processingTime: processingTime,
      stack: error.stack
    });
    
    throw error;
  }
}

/**
 * MFformatシートからCSVエクスポート
 * 
 * @param {Object} options - エクスポートオプション
 * @return {Object} エクスポート結果
 */
function exportMFformatToCSV(options = {}) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log(`[csvExporter.gs] MFformatエクスポート開始`);
    writeLog('INFO', 'CSVエクスポート', 'MFformatエクスポートを開始します');
    
    // オプションの準備
    const exportOptions = { ...EXPORT_CONFIG.DEFAULT_OPTIONS, ...options };
    
    // MFformatシートの取得
    const sheet = getSheetSafely(SHEET_NAMES.MFFORMAT);
    const data = getSheetDataForExport(sheet);
    
    if (!data || data.length === 0) {
      throw new Error('MFformatシートにデータがありません');
    }
    
    // CSVデータの生成
    const csvContent = generateCSVContent(data, exportOptions);
    
    // ファイル名の生成
    const fileName = generateFileName('mfformat', exportOptions);
    
    // Blobオブジェクトの作成
    const blob = createCSVBlob(csvContent, exportOptions.encoding);
    
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[csvExporter.gs] MFformatエクスポート完了: ${data.length}行, ${processingTime}ms`);
    
    // ログ記録
    writeLog('INFO', 'CSVエクスポート', `MFformatエクスポート完了: ${data.length}行`, {
      fileName: fileName,
      processingTime: processingTime,
      options: exportOptions
    });
    
    return {
      success: true,
      fileName: fileName,
      blob: blob,
      rowCount: data.length,
      processingTime: processingTime
    };
    
  } catch (error) {
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[csvExporter.gs] MFformatエクスポートエラー: ${error.message}`);
    
    writeLog('ERROR', 'CSVエクスポートエラー', `MFformat: ${error.message}`, {
      processingTime: processingTime,
      stack: error.stack
    });
    
    throw error;
  }
}

/**
 * エクスポート用シートデータ取得
 * 
 * @param {Sheet} sheet - 対象シート
 * @return {Array<Array>} シートデータ
 */
function getSheetDataForExport(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  
  // 行数制限チェック
  if (lastRow > EXPORT_CONFIG.MAX_EXPORT_ROWS) {
    throw new Error(`エクスポート行数が制限を超えています: ${lastRow}行 (最大: ${EXPORT_CONFIG.MAX_EXPORT_ROWS}行)`);
  }
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * CSVコンテンツの生成
 * 
 * @param {Array<Array>} data - シートデータ
 * @param {Object} options - エクスポートオプション
 * @return {string} CSVコンテンツ
 */
function generateCSVContent(data, options) {
  const csvRows = [];
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // ヘッダー行のスキップ処理
    if (i === 0 && !options.includeHeader) {
      continue;
    }
    
    // 行データの変換
    const csvRow = row.map(cell => formatCellForCSV(cell, options));
    
    // CSV行の構築
    const csvLine = csvRow.map(cell => escapeCsvField(cell, options.delimiter)).join(options.delimiter);
    csvRows.push(csvLine);
  }
  
  return csvRows.join('\n');
}

/**
 * セルデータのCSV用フォーマット
 * 
 * @param {*} cell - セル値
 * @param {Object} options - フォーマットオプション
 * @return {string} フォーマット済み値
 */
function formatCellForCSV(cell, options) {
  if (cell === null || cell === undefined) {
    return '';
  }
  
  // 日付の処理
  if (cell instanceof Date) {
    return formatDateForExport(cell, options.dateFormat);
  }
  
  // 数値の処理
  if (typeof cell === 'number') {
    return formatNumberForExport(cell, options.numberFormat);
  }
  
  // 文字列の処理
  return String(cell).trim();
}

/**
 * 日付のエクスポート用フォーマット
 * 
 * @param {Date} date - 日付オブジェクト
 * @param {string} format - フォーマット
 * @return {string} フォーマット済み日付
 */
function formatDateForExport(date, format) {
  if (!date || isNaN(date.getTime())) {
    return '';
  }
  
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  
  switch (format) {
    case 'YYYY/MM/DD':
      return `${year}/${month}/${day}`;
    case 'YYYY/MM/DD HH:mm:ss':
      return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
    case 'YYYY-MM-DD':
      return `${year}-${month}-${day}`;
    case 'YYYY-MM-DD HH:mm:ss':
      return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
    default:
      return `${year}/${month}/${day}`;
  }
}

/**
 * 数値のエクスポート用フォーマット
 * 
 * @param {number} number - 数値
 * @param {string} format - フォーマット
 * @return {string} フォーマット済み数値
 */
function formatNumberForExport(number, format) {
  if (isNaN(number)) {
    return '';
  }
  
  switch (format) {
    case 'comma':
      return number.toLocaleString();
    case 'plain':
      return String(number);
    default:
      // 整数の場合はそのまま、小数の場合は適切に表示
      return number % 1 === 0 ? String(number) : number.toString();
  }
}

/**
 * CSVフィールドのエスケープ処理
 * 
 * @param {string} field - フィールド値
 * @param {string} delimiter - 区切り文字
 * @return {string} エスケープ済みフィールド
 */
function escapeCsvField(field, delimiter) {
  const fieldStr = String(field);
  
  // エスケープが必要な条件
  const needsEscape = fieldStr.includes(delimiter) || 
                     fieldStr.includes('"') || 
                     fieldStr.includes('\n') || 
                     fieldStr.includes('\r');
  
  if (needsEscape) {
    // ダブルクォートをエスケープしてフィールド全体をクォートで囲む
    return `"${fieldStr.replace(/"/g, '""')}"`;
  }
  
  return fieldStr;
}

/**
 * ファイル名の生成
 * 
 * @param {string} type - エクスポートタイプ（ppformat/mfformat）
 * @param {Object} options - オプション
 * @return {string} ファイル名
 */
function generateFileName(type, options) {
  const now = new Date();
  const timestamp = formatDateForExport(now, 'YYYY-MM-DD_HH:mm:ss').replace(/[:\s]/g, '-');
  
  const template = EXPORT_CONFIG.FILENAME_TEMPLATES[type] || `${type}_export_{timestamp}.csv`;
  
  return template.replace('{timestamp}', timestamp);
}

/**
 * CSV Blobオブジェクトの作成
 * 
 * @param {string} csvContent - CSVコンテンツ
 * @param {string} encoding - 文字エンコーディング
 * @return {Blob} Blobオブジェクト
 */
function createCSVBlob(csvContent, encoding) {
  let mimeType = 'text/csv';
  
  // エンコーディングに応じてMIMEタイプを調整
  if (encoding === 'UTF-8') {
    mimeType = 'text/csv;charset=utf-8';
    // UTF-8 BOMを追加（Excelでの文字化け防止）
    csvContent = '\uFEFF' + csvContent;
  } else if (encoding === 'Shift_JIS') {
    mimeType = 'text/csv;charset=shift_jis';
  }
  
  return Utilities.newBlob(csvContent, mimeType);
}

/**
 * CSVファイルのダウンロード用URL生成
 * 
 * @param {Blob} blob - CSVファイルのBlob
 * @param {string} fileName - ファイル名
 * @return {string} ダウンロードURL
 */
function createDownloadUrl(blob, fileName) {
  try {
    // Google Driveに一時ファイルとして保存
    const file = DriveApp.createFile(blob.setName(fileName));
    
    // 共有設定を変更（リンクを知っている人は閲覧可能）
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // ダウンロードURL生成
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    
    Logger.log(`[csvExporter.gs] ファイル作成完了: ${fileName}, ID: ${file.getId()}`);
    
    // 24時間後に自動削除するトリガーを設定（オプション）
    // scheduleFileDeletion(file.getId());
    
    return {
      downloadUrl: downloadUrl,
      fileId: file.getId(),
      fileName: fileName
    };
    
  } catch (error) {
    Logger.log(`[csvExporter.gs] ファイル作成エラー: ${error.message}`);
    throw new Error(`ダウンロードファイルの作成に失敗しました: ${error.message}`);
  }
}

/**
 * エクスポート設定の取得
 * 
 * @return {Object} 現在のエクスポート設定
 */
function getExportSettings() {
  // 今後の拡張用：設定の永続化
  // PropertiesServiceを使用して設定を保存・取得
  
  return EXPORT_CONFIG.DEFAULT_OPTIONS;
}

/**
 * エクスポート設定の保存
 * 
 * @param {Object} settings - 保存する設定
 */
function saveExportSettings(settings) {
  // 今後の拡張用：設定の永続化
  // PropertiesServiceを使用して設定を保存
  
  Logger.log(`[csvExporter.gs] エクスポート設定保存: ${JSON.stringify(settings)}`);
}

/**
 * 一時ファイルのクリーンアップ
 * 
 * @param {string} fileId - 削除するファイルID
 */
function cleanupTempFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    Logger.log(`[csvExporter.gs] 一時ファイル削除完了: ${fileId}`);
  } catch (error) {
    Logger.log(`[csvExporter.gs] 一時ファイル削除エラー: ${fileId}, ${error.message}`);
  }
}