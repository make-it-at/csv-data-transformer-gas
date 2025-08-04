/**
 * データ転記・変換機能
 * LPcsvシートのデータをPPformat・MFformatシートに転記・加工する
 */

// 定数定義
const TRANSFER_CONFIG = {
  // デフォルト設定
  DEFAULT_SETTINGS: {
    ppformat: {
      enabled: true,
      columnMapping: {
        'ID': 'A',
        '日付': 'B', 
        '種別': 'C',
        'ポイント数': 'D',
        '相手': 'E'
      },
      filters: [],
      transformations: []
    },
    mfformat: {
      enabled: true,
      columnMapping: {
        'ID': 'A',
        '内容': 'B',
        'ポイント数': 'C', 
        'メッセージ': 'D'
      },
      filters: [],
      transformations: []
    }
  },
  
  // 最大処理行数
  MAX_ROWS: 10000,
  
  // バッチサイズ
  BATCH_SIZE: 500
};

/**
 * メイン転記処理
 * LPcsvのデータをPPformat・MFformatに転記
 * 
 * @param {Object} settings - 転記設定
 * @return {Object} 処理結果
 */
function transferDataFromLPcsv(settings = null) {
  const startTime = new Date().getTime();
  
  try {
    Logger.log(`[dataTransformer.gs] データ転記開始`);
    writeLog('INFO', 'データ転記', 'データ転記処理を開始します');
    
    // 設定の準備
    const config = settings || TRANSFER_CONFIG.DEFAULT_SETTINGS;
    
    // ソースシート（LPcsv）の取得
    const sourceSheet = getSheetSafely(SHEET_NAMES.LPCSV);
    const sourceData = getSheetData(sourceSheet);
    
    if (!sourceData || sourceData.length <= 1) {
      throw new Error('LPcsvシートにデータがありません');
    }
    
    Logger.log(`[dataTransformer.gs] ソースデータ: ${sourceData.length}行`);
    
    const results = {
      ppformat: null,
      mfformat: null,
      totalProcessed: 0,
      errors: []
    };
    
    // PPformatへの転記
    if (config.ppformat.enabled) {
      try {
        results.ppformat = transferToPPformat(sourceData, config.ppformat);
        Logger.log(`[dataTransformer.gs] PPformat転記完了: ${results.ppformat.processedRows}行`);
      } catch (error) {
        results.errors.push(`PPformat転記エラー: ${error.message}`);
        Logger.log(`[dataTransformer.gs] PPformat転記エラー: ${error.message}`);
      }
    }
    
    // MFformatへの転記
    if (config.mfformat.enabled) {
      try {
        results.mfformat = transferToMFformat(sourceData, config.mfformat);
        Logger.log(`[dataTransformer.gs] MFformat転記完了: ${results.mfformat.processedRows}行`);
      } catch (error) {
        results.errors.push(`MFformat転記エラー: ${error.message}`);
        Logger.log(`[dataTransformer.gs] MFformat転記エラー: ${error.message}`);
      }
    }
    
    // 処理結果の集計
    results.totalProcessed = (results.ppformat?.processedRows || 0) + (results.mfformat?.processedRows || 0);
    
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[dataTransformer.gs] データ転記完了: ${results.totalProcessed}行, ${processingTime}ms`);
    
    // ログ記録
    writeLog('INFO', 'データ転記', `転記完了: PPformat=${results.ppformat?.processedRows || 0}行, MFformat=${results.mfformat?.processedRows || 0}行`, {
      processingTime: processingTime,
      errors: results.errors
    });
    
    return results;
    
  } catch (error) {
    const processingTime = new Date().getTime() - startTime;
    Logger.log(`[dataTransformer.gs] データ転記エラー: ${error.message}`);
    
    writeLog('ERROR', 'データ転記エラー', error.message, {
      processingTime: processingTime,
      stack: error.stack
    });
    
    throw error;
  }
}

/**
 * PPformatシートへの転記
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - PPformat設定
 * @return {Object} 転記結果
 */
function transferToPPformat(sourceData, config) {
  const targetSheet = getSheetSafely(SHEET_NAMES.PPFORMAT);
  
  // ヘッダー行の設定
  const headers = extractHeaders(sourceData);
  const targetHeaders = buildTargetHeaders(headers, config.columnMapping);
  
  // データの変換
  const transformedData = transformData(sourceData, config, headers);
  
  // シートに書き込み
  writeDataToSheet(targetSheet, [targetHeaders, ...transformedData]);
  
  return {
    processedRows: transformedData.length,
    targetColumns: targetHeaders.length
  };
}

/**
 * MFformatシートへの転記
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - MFformat設定
 * @return {Object} 転記結果
 */
function transferToMFformat(sourceData, config) {
  const targetSheet = getSheetSafely(SHEET_NAMES.MFFORMAT);
  
  // ヘッダー行の設定
  const headers = extractHeaders(sourceData);
  const targetHeaders = buildTargetHeaders(headers, config.columnMapping);
  
  // データの変換
  const transformedData = transformData(sourceData, config, headers);
  
  // シートに書き込み
  writeDataToSheet(targetSheet, [targetHeaders, ...transformedData]);
  
  return {
    processedRows: transformedData.length,
    targetColumns: targetHeaders.length
  };
}

/**
 * シートからデータを取得
 * 
 * @param {Sheet} sheet - 対象シート
 * @return {Array<Array>} シートデータ
 */
function getSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

/**
 * ターゲットヘッダーの構築
 * 
 * @param {Array<string>} sourceHeaders - ソースヘッダー
 * @param {Object} columnMapping - 列マッピング
 * @return {Array<string>} ターゲットヘッダー
 */
function buildTargetHeaders(sourceHeaders, columnMapping) {
  const targetHeaders = [];
  
  // マッピング設定に基づいてヘッダーを構築
  for (const [sourceCol, targetPos] of Object.entries(columnMapping)) {
    if (sourceHeaders.includes(sourceCol)) {
      targetHeaders.push(sourceCol);
    }
  }
  
  return targetHeaders;
}

/**
 * データ変換処理
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - 変換設定
 * @param {Array<string>} headers - ヘッダー行
 * @return {Array<Array>} 変換済みデータ
 */
function transformData(sourceData, config, headers) {
  const transformedData = [];
  
  // ヘッダー行をスキップ（1行目）
  for (let i = 1; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];
    
    // フィルタリング処理
    if (!passesFilters(sourceRow, headers, config.filters)) {
      continue;
    }
    
    // 列マッピングに基づいてデータを変換
    const targetRow = buildTargetRow(sourceRow, headers, config.columnMapping);
    
    // データ変換処理を適用
    const transformedRow = applyTransformations(targetRow, config.transformations);
    
    transformedData.push(transformedRow);
  }
  
  return transformedData;
}

/**
 * ターゲット行の構築
 * 
 * @param {Array} sourceRow - ソース行データ
 * @param {Array<string>} headers - ヘッダー行
 * @param {Object} columnMapping - 列マッピング
 * @return {Array} ターゲット行データ
 */
function buildTargetRow(sourceRow, headers, columnMapping) {
  const targetRow = [];
  
  for (const [sourceCol, targetPos] of Object.entries(columnMapping)) {
    const sourceIndex = headers.indexOf(sourceCol);
    const value = sourceIndex >= 0 ? sourceRow[sourceIndex] : '';
    targetRow.push(value);
  }
  
  return targetRow;
}

/**
 * フィルタリング処理
 * 
 * @param {Array} row - データ行
 * @param {Array<string>} headers - ヘッダー行
 * @param {Array} filters - フィルター設定
 * @return {boolean} フィルター通過可否
 */
function passesFilters(row, headers, filters) {
  if (!filters || filters.length === 0) {
    return true;
  }
  
  // 今後の拡張用：フィルター条件の実装
  // 例: 種別が「獲得」のみ、ポイント数が正の値のみ等
  
  return true;
}

/**
 * データ変換処理の適用
 * 
 * @param {Array} row - データ行
 * @param {Array} transformations - 変換設定
 * @return {Array} 変換済み行データ
 */
function applyTransformations(row, transformations) {
  if (!transformations || transformations.length === 0) {
    return row;
  }
  
  // 今後の拡張用：データ変換ルールの実装
  // 例: 日付フォーマット変更、数値計算、文字列置換等
  
  return row;
}

/**
 * シートにデータを書き込み
 * 
 * @param {Sheet} sheet - 対象シート
 * @param {Array<Array>} data - 書き込みデータ
 */
function writeDataToSheet(sheet, data) {
  if (!data || data.length === 0) {
    Logger.log(`[dataTransformer.gs] 書き込みデータがありません`);
    return;
  }
  
  // シートをクリア
  sheet.clear();
  
  // データを書き込み
  const range = sheet.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
  
  Logger.log(`[dataTransformer.gs] シート書き込み完了: ${data.length}行 x ${data[0].length}列`);
}

/**
 * 転記設定の取得
 * 
 * @return {Object} 現在の転記設定
 */
function getTransferSettings() {
  // 今後の拡張用：設定の永続化
  // PropertiesServiceを使用して設定を保存・取得
  
  return TRANSFER_CONFIG.DEFAULT_SETTINGS;
}

/**
 * 転記設定の保存
 * 
 * @param {Object} settings - 保存する設定
 */
function saveTransferSettings(settings) {
  // 今後の拡張用：設定の永続化
  // PropertiesServiceを使用して設定を保存
  
  Logger.log(`[dataTransformer.gs] 設定保存: ${JSON.stringify(settings)}`);
}