/**
 * データ転記・変換機能
 * LPcsvシートのデータをPPformat・MFformatシートに転記・加工する
 */

// 定数定義
const TRANSFER_CONFIG = {
  // PPformat（PayPay履歴形式）の列定義
  PPFORMAT_COLUMNS: [
    '取引日',
    '出金金額（円）',
    '入金金額（円）',
    '海外出金金額',
    '通貨',
    '変換レート（円）',
    '利用国',
    '取引内容',
    '取引先',
    '取引方法',
    '支払い区分',
    '利用者',
    '取引番号'
  ],
  
  // MFformat（マネーフォワード仕訳形式）の列定義
  MFFORMAT_COLUMNS: [
    '取引No',
    '取引日',
    '借方勘定科目',
    '借方補助科目',
    '借方部門',
    '借方取引先',
    '借方税区分',
    '借方インボイス',
    '借方金額(円)',
    '借方税額',
    '貸方勘定科目',
    '貸方補助科目',
    '貸方部門',
    '貸方取引先',
    '貸方税区分',
    '貸方インボイス',
    '貸方金額(円)',
    '貸方税額',
    '摘要',
    '仕訳メモ',
    'タグ',
    'MF仕訳タイプ',
    '決算整理仕訳',
    '作成日時',
    '作成者',
    '最終更新日時',
    '最終更新者'
  ],
  
  // デフォルト設定
  DEFAULT_SETTINGS: {
    ppformat: {
      enabled: true,
      filters: {
        includeTypes: ['獲得', '利用'], // 含める種別
        excludeTypes: [], // 除外する種別
        minAmount: null, // 最小金額
        maxAmount: null  // 最大金額
      },
      transformations: {
        dateFormat: 'YYYY/MM/DD 00:00:00', // 日付フォーマット（時刻は00:00:00固定）
        amountHandling: 'split' // 金額の処理方法（split: 獲得/利用で分ける, single: 単一列）
      }
    },
    mfformat: {
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
            debit: 'ポイント',
            credit: 'ポイント収益'
          },
          利用: {
            debit: 'ポイント利用',
            credit: 'ポイント'
          }
        },
        taxCategories: {
          獲得: '対象外',
          利用: '対象外'
        },
        includeMemo: false // 仕訳メモを含めるかどうか
      }
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
 * PPformatシートへの転記（PayPay履歴形式）
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - PPformat設定
 * @return {Object} 転記結果
 */
function transferToPPformat(sourceData, config) {
  const targetSheet = getSheetSafely(SHEET_NAMES.PPFORMAT);
  
  // ヘッダー行の設定
  const headerInfo = extractHeaders(sourceData);
  const sourceHeaders = headerInfo.headers;
  const targetHeaders = TRANSFER_CONFIG.PPFORMAT_COLUMNS;
  
  // データの変換
  const transformedData = transformToPPformatData(sourceData, config, sourceHeaders);
  
  // シートに書き込み
  writeDataToSheet(targetSheet, [targetHeaders, ...transformedData]);
  
  return {
    processedRows: transformedData.length,
    targetColumns: targetHeaders.length
  };
}

/**
 * PPformat用データ変換
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - 設定
 * @param {Array<string>} sourceHeaders - ソースヘッダー
 * @return {Array<Array>} 変換済みデータ
 */
function transformToPPformatData(sourceData, config, sourceHeaders) {
  const transformedData = [];
  
  // ヘッダー行をスキップ（1行目）
  for (let i = 1; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];
    
    // フィルタリング処理
    if (!passesPPformatFilters(sourceRow, sourceHeaders, config.filters)) {
      continue;
    }
    
    // PayPay履歴形式に変換
    const targetRow = buildPPformatRow(sourceRow, sourceHeaders, config.transformations);
    
    transformedData.push(targetRow);
  }
  
  return transformedData;
}

/**
 * PPformat行の構築
 * 
 * @param {Array} sourceRow - ソース行データ
 * @param {Array<string>} sourceHeaders - ソースヘッダー
 * @param {Object} transformations - 変換設定
 * @return {Array} PPformat行データ
 */
function buildPPformatRow(sourceRow, sourceHeaders, transformations) {
  // ソースデータのインデックス取得
  const idIndex = sourceHeaders.indexOf('ID');
  const dateIndex = sourceHeaders.indexOf('日付');
  const typeIndex = sourceHeaders.indexOf('種別');
  const contentIndex = sourceHeaders.indexOf('内容');
  const amountIndex = sourceHeaders.indexOf('ポイント数');
  const messageIndex = sourceHeaders.indexOf('メッセージ');
  const partnerIndex = sourceHeaders.indexOf('相手');
  
  // ソースデータ取得
  const id = sourceRow[idIndex] || '';
  const date = sourceRow[dateIndex] || '';
  const type = sourceRow[typeIndex] || '';
  const content = sourceRow[contentIndex] || '';
  const amount = sourceRow[amountIndex] || '';
  const message = sourceRow[messageIndex] || '';
  const partner = sourceRow[partnerIndex] || '';
  
  // 日付フォーマット変換
  const formattedDate = formatDateForPPformat(date, transformations.dateFormat);
  
  // 金額処理（獲得/利用で出金・入金を分ける）
  let withdrawAmount = '-'; // 出金金額
  let depositAmount = '-';  // 入金金額
  
  if (transformations.amountHandling === 'split') {
    const numericAmount = parseFloat(amount.toString().replace(/[^\d.-]/g, ''));
    if (!isNaN(numericAmount)) {
      if (type === '利用' && numericAmount < 0) {
        withdrawAmount = Math.abs(numericAmount).toLocaleString();
      } else if (type === '獲得' && numericAmount > 0) {
        depositAmount = numericAmount.toLocaleString();
      }
    }
  }
  
  // PPformat行データ構築（13列）
  return [
    formattedDate,           // 取引日
    withdrawAmount,          // 出金金額（円）
    depositAmount,           // 入金金額（円）
    '-',                     // 海外出金金額
    '-',                     // 通貨
    '-',                     // 変換レート（円）
    '-',                     // 利用国
    '-',                     // 取引内容
    partner || '',           // 取引先（相手）
    '-',                     // 取引方法
    '-',                     // 支払い区分
    '-',                     // 利用者
    id                       // 取引番号
  ];
}

/**
 * PPformat用フィルタリング
 * 
 * @param {Array} row - データ行
 * @param {Array<string>} headers - ヘッダー行
 * @param {Object} filters - フィルター設定
 * @return {boolean} フィルター通過可否
 */
function passesPPformatFilters(row, headers, filters) {
  const typeIndex = headers.indexOf('種別');
  const amountIndex = headers.indexOf('ポイント数');
  
  const type = row[typeIndex] || '';
  const amount = parseFloat((row[amountIndex] || '').toString().replace(/[^\d.-]/g, ''));
  
  // 種別フィルター
  if (filters.includeTypes && filters.includeTypes.length > 0) {
    if (!filters.includeTypes.includes(type)) {
      return false;
    }
  }
  
  if (filters.excludeTypes && filters.excludeTypes.length > 0) {
    if (filters.excludeTypes.includes(type)) {
      return false;
    }
  }
  
  // 金額フィルター
  if (filters.minAmount !== null && !isNaN(amount) && amount < filters.minAmount) {
    return false;
  }
  
  if (filters.maxAmount !== null && !isNaN(amount) && amount > filters.maxAmount) {
    return false;
  }
  
  return true;
}

/**
 * PPformat用日付フォーマット
 * 
 * @param {string} dateStr - 日付文字列
 * @param {string} format - フォーマット
 * @return {string} フォーマット済み日付
 */
function formatDateForPPformat(dateStr, format) {
  if (!dateStr) return '';
  
  try {
    Logger.log(`[dataTransformer.gs] PPformat日付変換入力: ${dateStr}`);
    
    // タイムゾーン情報を含む日付文字列を適切に処理
    let date;
    if (dateStr.includes('PDT') || dateStr.includes('PST')) {
      // 太平洋時間の場合、UTCとして解釈してから日本時間に変換
      const utcStr = dateStr.replace(/PDT|PST/, 'UTC');
      date = new Date(utcStr);
      // 日本時間（JST）に変換（UTC+9）
      date.setHours(date.getHours() + 9);
    } else {
      date = new Date(dateStr);
    }
    
    if (isNaN(date.getTime())) {
      Logger.log(`[dataTransformer.gs] 無効な日付: ${dateStr}`);
      return dateStr;
    }
    
    // YYYY/MM/DD 00:00:00 形式（時刻は常に00:00:00）
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    const result = `${year}/${month}/${day} 00:00:00`;
    Logger.log(`[dataTransformer.gs] PPformat日付変換結果: ${result}`);
    return result;
  } catch (error) {
    Logger.log(`[dataTransformer.gs] 日付変換エラー: ${dateStr}, ${error.message}`);
    return dateStr;
  }
}

/**
 * MFformatシートへの転記（マネーフォワード仕訳形式）
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - MFformat設定
 * @return {Object} 転記結果
 */
function transferToMFformat(sourceData, config) {
  const targetSheet = getSheetSafely(SHEET_NAMES.MFFORMAT);
  
  // 設定のデバッグログ
  Logger.log(`[dataTransformer.gs] MFformat転記設定: ${JSON.stringify(config)}`);
  
  // ヘッダー行の設定
  const headerInfo = extractHeaders(sourceData);
  const sourceHeaders = headerInfo.headers;
  const targetHeaders = TRANSFER_CONFIG.MFFORMAT_COLUMNS;
  
  // データの変換
  const transformedData = transformToMFformatData(sourceData, config, sourceHeaders);
  
  // シートに書き込み
  writeDataToSheet(targetSheet, [targetHeaders, ...transformedData]);
  
  return {
    processedRows: transformedData.length,
    targetColumns: targetHeaders.length
  };
}

/**
 * MFformat用データ変換
 * 
 * @param {Array<Array>} sourceData - ソースデータ
 * @param {Object} config - 設定
 * @param {Array<string>} sourceHeaders - ソースヘッダー
 * @return {Array<Array>} 変換済みデータ
 */
function transformToMFformatData(sourceData, config, sourceHeaders) {
  const transformedData = [];
  let filteredCount = 0;
  
  Logger.log(`[dataTransformer.gs] MFformat変換開始: ${sourceData.length - 1}行処理予定`);
  
  // 日付インデックスを取得
  const dateIndex = sourceHeaders.indexOf('日付');
  
  // ヘッダー行をスキップしてデータ行のみを取得
  const dataRows = sourceData.slice(1);
  
  // 日付順でソート（古い順）
  const sortedRows = dataRows.sort((a, b) => {
    const dateA = new Date(a[dateIndex] || '');
    const dateB = new Date(b[dateIndex] || '');
    
    // 無効な日付は最後に配置
    if (isNaN(dateA.getTime()) && isNaN(dateB.getTime())) return 0;
    if (isNaN(dateA.getTime())) return 1;
    if (isNaN(dateB.getTime())) return -1;
    
    return dateA.getTime() - dateB.getTime();
  });
  
  Logger.log(`[dataTransformer.gs] データを日付順でソート完了: ${sortedRows.length}行`);
  
  // ソート済みデータを処理
  let transactionNo = 1;
  for (let i = 0; i < sortedRows.length; i++) {
    const sourceRow = sortedRows[i];
    
    // フィルタリング処理
    if (!passesMFformatFilters(sourceRow, sourceHeaders, config.filters)) {
      filteredCount++;
      continue;
    }
    
    // マネーフォワード仕訳形式に変換
    const targetRow = buildMFformatRow(sourceRow, sourceHeaders, config.transformations, transactionNo);
    
    transformedData.push(targetRow);
    transactionNo++;
  }
  
  Logger.log(`[dataTransformer.gs] MFformat変換完了: ${transformedData.length}行変換, ${filteredCount}行除外（日付順）`);
  
  return transformedData;
}

/**
 * MFformat行の構築
 * 
 * @param {Array} sourceRow - ソース行データ
 * @param {Array<string>} sourceHeaders - ソースヘッダー
 * @param {Object} transformations - 変換設定
 * @param {number} transactionNo - 取引番号
 * @return {Array} MFformat行データ
 */
function buildMFformatRow(sourceRow, sourceHeaders, transformations, transactionNo) {
  // 設定のデバッグログ
  Logger.log(`[dataTransformer.gs] buildMFformatRow設定: ${JSON.stringify(transformations)}`);
  
  // ソースデータのインデックス取得
  const idIndex = sourceHeaders.indexOf('ID');
  const dateIndex = sourceHeaders.indexOf('日付');
  const typeIndex = sourceHeaders.indexOf('種別');
  const contentIndex = sourceHeaders.indexOf('内容');
  const amountIndex = sourceHeaders.indexOf('ポイント数');
  const messageIndex = sourceHeaders.indexOf('メッセージ');
  const partnerIndex = sourceHeaders.indexOf('相手');
  
  // ソースデータ取得
  const id = sourceRow[idIndex] || '';
  const date = sourceRow[dateIndex] || '';
  const type = sourceRow[typeIndex] || '';
  const content = sourceRow[contentIndex] || '';
  const amount = sourceRow[amountIndex] || '';
  const message = sourceRow[messageIndex] || '';
  const partner = sourceRow[partnerIndex] || '';
  
  // 日付フォーマット変換
  const formattedDate = formatDateForMFformat(date, transformations.dateFormat);
  
  // 金額処理
  const numericAmount = Math.abs(parseFloat(amount.toString().replace(/[^\d.-]/g, '')) || 0);
  
  // 勘定科目の決定（ユーザー設定に基づく）
  let debitAccount = '';
  let creditAccount = '';
  let taxCategory = '対象外';
  
  Logger.log(`[dataTransformer.gs] 種別: ${type}, 設定: ${JSON.stringify(transformations.accounts)}`);
  
  if (transformations.accounts && transformations.accounts[type]) {
    debitAccount = transformations.accounts[type].debit || '';
    creditAccount = transformations.accounts[type].credit || '';
    Logger.log(`[dataTransformer.gs] 勘定科目決定: 借方=${debitAccount}, 貸方=${creditAccount}`);
  } else {
    Logger.log(`[dataTransformer.gs] 勘定科目設定が見つかりません: 種別=${type}`);
  }
  
  // 税区分の決定（ユーザー設定に基づく）
  Logger.log(`[dataTransformer.gs] 税区分設定: ${JSON.stringify(transformations.taxCategories)}`);
  if (transformations.taxCategories && transformations.taxCategories[type]) {
    taxCategory = transformations.taxCategories[type];
    Logger.log(`[dataTransformer.gs] 税区分決定: ${taxCategory}`);
  } else {
    Logger.log(`[dataTransformer.gs] 税区分設定が見つかりません: 種別=${type}`);
  }
  
  // 現在時刻
  const now = new Date();
  const createdAt = formatDateForMFformat(now.toISOString(), 'YYYY/MM/DD HH:mm:ss');
  
  // MFformat行データ構築（27列）
  return [
    transactionNo,                    // 取引No
    formattedDate,                    // 取引日
    debitAccount,                     // 借方勘定科目
    '',                               // 借方補助科目
    '',                               // 借方部門
    type === '利用' ? (partner || '') : '',  // 借方取引先（利用時のみ相手）
    taxCategory,                      // 借方税区分
    '',                               // 借方インボイス
    numericAmount,                    // 借方金額(円)
    0,                                // 借方税額
    creditAccount,                    // 貸方勘定科目
    '',                               // 貸方補助科目
    '',                               // 貸方部門
    type === '獲得' ? (partner || '') : '',  // 貸方取引先（獲得時のみ相手）
    taxCategory,                      // 貸方税区分
    '',                               // 貸方インボイス
    numericAmount,                    // 貸方金額(円)
    0,                                // 貸方税額
    content || message || '',         // 摘要
    transformations.includeMemo ? `${type}:${id}` : '',  // 仕訳メモ（設定により制御）
    '',                               // タグ
    '',                               // MF仕訳タイプ
    '',                               // 決算整理仕訳
    createdAt,                        // 作成日時
    'GAS転記',                        // 作成者
    createdAt,                        // 最終更新日時
    'GAS転記'                         // 最終更新者
  ];
}

/**
 * MFformat用フィルタリング
 * 
 * @param {Array} row - データ行
 * @param {Array<string>} headers - ヘッダー行
 * @param {Object} filters - フィルター設定
 * @return {boolean} フィルター通過可否
 */
function passesMFformatFilters(row, headers, filters) {
  const typeIndex = headers.indexOf('種別');
  const amountIndex = headers.indexOf('ポイント数');
  
  const type = row[typeIndex] || '';
  const amount = parseFloat((row[amountIndex] || '').toString().replace(/[^\d.-]/g, ''));
  
  // 種別フィルター
  if (filters.includeTypes && filters.includeTypes.length > 0) {
    if (!filters.includeTypes.includes(type)) {
      return false;
    }
  }
  
  if (filters.excludeTypes && filters.excludeTypes.length > 0) {
    if (filters.excludeTypes.includes(type)) {
      return false;
    }
  }
  
  // 金額フィルター
  if (filters.minAmount !== null && !isNaN(amount) && Math.abs(amount) < filters.minAmount) {
    return false;
  }
  
  if (filters.maxAmount !== null && !isNaN(amount) && Math.abs(amount) > filters.maxAmount) {
    return false;
  }
  
  return true;
}

/**
 * MFformat用日付フォーマット
 * 
 * @param {string} dateStr - 日付文字列
 * @param {string} format - フォーマット
 * @return {string} フォーマット済み日付
 */
function formatDateForMFformat(dateStr, format) {
  if (!dateStr) return '';
  
  try {
    Logger.log(`[dataTransformer.gs] MFformat日付変換入力: ${dateStr}`);
    
    // タイムゾーン情報を含む日付文字列を適切に処理
    let date;
    if (dateStr.includes('PDT') || dateStr.includes('PST')) {
      // 太平洋時間の場合、UTCとして解釈してから日本時間に変換
      const utcStr = dateStr.replace(/PDT|PST/, 'UTC');
      date = new Date(utcStr);
      // 日本時間（JST）に変換（UTC+9）
      date.setHours(date.getHours() + 9);
    } else {
      date = new Date(dateStr);
    }
    
    if (isNaN(date.getTime())) {
      Logger.log(`[dataTransformer.gs] 無効な日付: ${dateStr}`);
      return dateStr;
    }
    
    // YYYY/MM/DD 形式
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    if (format && format.includes('HH:mm:ss')) {
      const hours = String(date.getHours()).padStart(2, '0');
      const minutes = String(date.getMinutes()).padStart(2, '0');
      const seconds = String(date.getSeconds()).padStart(2, '0');
      const result = `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
      Logger.log(`[dataTransformer.gs] MFformat日付変換結果（時刻付き）: ${result}`);
      return result;
    }
    
    const result = `${year}/${month}/${day}`;
    Logger.log(`[dataTransformer.gs] MFformat日付変換結果: ${result}`);
    return result;
  } catch (error) {
    Logger.log(`[dataTransformer.gs] 日付変換エラー: ${dateStr}, ${error.message}`);
    return dateStr;
  }
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