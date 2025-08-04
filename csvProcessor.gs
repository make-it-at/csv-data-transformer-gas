/**
 * CSV処理機能
 * 
 * CSVファイルの解析、検証、インポート処理を担当
 */

/**
 * CSV文字列を解析して2次元配列に変換
 * 
 * @param {string} csvContent - CSV文字列
 * @param {Object} options - 解析オプション
 * @return {Array<Array<string>>} 解析されたCSVデータ
 */
function parseCSVContent(csvContent, options = {}) {
  try {
    Logger.log('[csvProcessor.gs] CSV解析開始');
    
    if (!csvContent || csvContent.trim() === '') {
      throw new Error('CSVファイルが空です');
    }
    
    // デフォルトオプション
    const defaultOptions = {
      delimiter: ',',
      hasHeader: true,
      maxRows: 10000
    };
    
    const opts = { ...defaultOptions, ...options };
    
    // 行に分割（改行文字の統一）
    const lines = csvContent.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    
    // 空行を除去
    const nonEmptyLines = lines.filter(line => line.trim() !== '');
    
    if (nonEmptyLines.length === 0) {
      throw new Error('有効なデータ行がありません');
    }
    
    // 行数制限チェック
    if (nonEmptyLines.length > opts.maxRows) {
      throw new Error(`行数が制限を超えています: ${nonEmptyLines.length}行 (最大: ${opts.maxRows}行)`);
    }
    
    // CSV解析
    const csvData = [];
    
    for (let i = 0; i < nonEmptyLines.length; i++) {
      const line = nonEmptyLines[i];
      const row = parseCSVLine(line, opts.delimiter);
      csvData.push(row);
    }
    
    Logger.log(`[csvProcessor.gs] CSV解析完了: ${csvData.length}行, ${csvData[0]?.length || 0}列`);
    
    return csvData;
    
  } catch (error) {
    Logger.log(`[csvProcessor.gs] CSV解析エラー: ${error.message}`);
    throw error;
  }
}

/**
 * CSV行を解析して配列に変換
 * 
 * @param {string} line - CSV行
 * @param {string} delimiter - 区切り文字
 * @return {Array<string>} 解析された行データ
 */
function parseCSVLine(line, delimiter = ',') {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const nextChar = line[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // エスケープされたクォート
        current += '"';
        i++; // 次の文字をスキップ
      } else {
        // クォートの開始/終了
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      // 区切り文字（クォート外）
      result.push(current.trim());
      current = '';
    } else {
      // 通常の文字
      current += char;
    }
  }
  
  // 最後のフィールドを追加
  result.push(current.trim());
  
  return result;
}

/**
 * CSVデータの検証
 * 
 * @param {Array<Array<string>>} csvData - CSVデータ
 * @param {Object} options - 検証オプション
 */
function validateCSVData(csvData, options = {}) {
  try {
    Logger.log('[csvProcessor.gs] データ検証開始');
    
    if (!Array.isArray(csvData) || csvData.length === 0) {
      throw new Error('CSVデータが無効です');
    }
    
    // 最大列数チェック
    const maxColumns = 100;
    const firstRowColumns = csvData[0].length;
    
    if (firstRowColumns > maxColumns) {
      throw new Error(`列数が制限を超えています: ${firstRowColumns}列 (最大: ${maxColumns}列)`);
    }
    
    // 各行の列数一貫性チェック
    const inconsistentRows = [];
    for (let i = 1; i < csvData.length; i++) {
      if (csvData[i].length !== firstRowColumns) {
        inconsistentRows.push(i + 1); // 1ベースの行番号
      }
    }
    
    if (inconsistentRows.length > 0) {
      const maxReportRows = 5;
      const reportRows = inconsistentRows.slice(0, maxReportRows);
      const moreRows = inconsistentRows.length > maxReportRows ? ` (他${inconsistentRows.length - maxReportRows}行)` : '';
      
      Logger.log(`[csvProcessor.gs] 警告: 列数が不一致の行があります: ${reportRows.join(', ')}行目${moreRows}`);
      // 警告として記録するが、処理は継続
    }
    
    // ヘッダー行の検証（存在する場合）
    if (options.hasHeader && csvData.length > 0) {
      const headers = csvData[0];
      const emptyHeaders = headers.filter((header, index) => {
        return !header || header.trim() === '';
      });
      
      if (emptyHeaders.length > 0) {
        Logger.log(`[csvProcessor.gs] 警告: 空のヘッダーが${emptyHeaders.length}個あります`);
      }
    }
    
    Logger.log('[csvProcessor.gs] データ検証完了');
    
  } catch (error) {
    Logger.log(`[csvProcessor.gs] データ検証エラー: ${error.message}`);
    throw error;
  }
}

/**
 * ヘッダー行の抽出・分析
 * 
 * @param {Array<Array<string>>} csvData - CSVデータ
 * @return {Object} ヘッダー情報
 */
function extractHeaders(csvData) {
  try {
    if (!csvData || csvData.length === 0) {
      return { hasHeaders: false, headers: [], columnCount: 0 };
    }
    
    const firstRow = csvData[0];
    const hasHeaders = firstRow.some(cell => isNaN(cell) && cell.trim() !== '');
    
    return {
      hasHeaders: hasHeaders,
      headers: hasHeaders ? firstRow : firstRow.map((_, index) => `列${index + 1}`),
      columnCount: firstRow.length
    };
    
  } catch (error) {
    Logger.log(`[csvProcessor.gs] ヘッダー抽出エラー: ${error.message}`);
    throw error;
  }
}

/**
 * LPcsvシートにデータをインポート
 * 
 * @param {Array<Array<string>>} csvData - CSVデータ
 * @param {Sheet} targetSheet - 対象シート
 */
function importToLPcsv(csvData, targetSheet) {
  try {
    Logger.log('[csvProcessor.gs] LPcsvインポート開始');
    
    if (!csvData || csvData.length === 0) {
      throw new Error('インポートするデータがありません');
    }
    
    if (!targetSheet) {
      throw new Error('対象シートが見つかりません');
    }
    
    // シートをクリア（既存データを削除）
    targetSheet.clear();
    
    // データをバッチで書き込み
    const batchSize = BATCH_SIZE || 1000;
    let processedRows = 0;
    
    for (let i = 0; i < csvData.length; i += batchSize) {
      const batch = csvData.slice(i, Math.min(i + batchSize, csvData.length));
      const startRow = i + 1; // 1ベースの行番号
      const numRows = batch.length;
      const numCols = batch[0].length;
      
      // 範囲を取得して値を設定
      const range = targetSheet.getRange(startRow, 1, numRows, numCols);
      range.setValues(batch);
      
      processedRows += numRows;
      Logger.log(`[csvProcessor.gs] 処理済み: ${processedRows}/${csvData.length}行`);
    }
    
    // ヘッダー行のフォーマット（1行目がヘッダーの場合）
    if (csvData.length > 0) {
      const headerRange = targetSheet.getRange(1, 1, 1, csvData[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8f4fd');
      
      // 列幅の自動調整
      for (let col = 1; col <= csvData[0].length; col++) {
        targetSheet.autoResizeColumn(col);
      }
    }
    
    Logger.log(`[csvProcessor.gs] LPcsvインポート完了: ${csvData.length}行, ${csvData[0].length}列`);
    
  } catch (error) {
    Logger.log(`[csvProcessor.gs] LPcsvインポートエラー: ${error.message}`);
    throw error;
  }
}

/**
 * 文字コードの自動判定（簡易版）
 * 
 * @param {string} content - ファイル内容
 * @return {string} 推定される文字コード
 */
function detectEncoding(content) {
  try {
    // 簡易的な文字コード判定
    // 実際のGASでは詳細な判定は困難なため、基本的な判定のみ
    
    if (!content) {
      return 'UTF-8';
    }
    
    // BOMの確認
    if (content.charCodeAt(0) === 0xFEFF) {
      return 'UTF-8 (BOM)';
    }
    
    // 日本語文字の存在確認
    const hasJapanese = /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(content);
    
    if (hasJapanese) {
      // 文字化けパターンの確認
      const hasMojibake = /[・｡｢｣]/.test(content);
      if (hasMojibake) {
        return 'Shift_JIS (推定)';
      }
    }
    
    return 'UTF-8';
    
  } catch (error) {
    Logger.log(`[csvProcessor.gs] 文字コード判定エラー: ${error.message}`);
    return 'UTF-8';
  }
}