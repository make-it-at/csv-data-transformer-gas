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
    
    // 改良されたCSV解析を使用（複数行クォート対応）
    const csvData = parseCSVWithMultilineSupport(csvContent, opts.delimiter);
    
    if (csvData.length === 0) {
      throw new Error('有効なデータ行がありません');
    }
    
    // 行数制限チェック
    if (csvData.length > opts.maxRows) {
      throw new Error(`行数が制限を超えています: ${csvData.length}行 (最大: ${opts.maxRows}行)`);
    }
    
    Logger.log(`[csvProcessor.gs] CSV解析完了: ${csvData.length}行, ${csvData[0]?.length || 0}列`);
    
    return csvData;
    
  } catch (error) {
    Logger.log(`[csvProcessor.gs] CSV解析エラー: ${error.message}`);
    throw error;
  }
}

/**
 * CSV行を解析して配列に変換（改行を含むクォート対応）
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
      // 通常の文字（改行も含む）
      current += char;
    }
  }
  
  // 最後のフィールドを追加
  result.push(current.trim());
  
  return result;
}

/**
 * 改良されたCSV解析（複数行にまたがるクォートフィールド対応）
 * 
 * @param {string} csvContent - CSV文字列全体
 * @param {string} delimiter - 区切り文字
 * @return {Array<Array<string>>} 解析されたCSVデータ
 */
function parseCSVWithMultilineSupport(csvContent, delimiter = ',') {
  const result = [];
  let current = '';
  let inQuotes = false;
  let currentRow = [];
  let currentField = '';
  
  for (let i = 0; i < csvContent.length; i++) {
    const char = csvContent[i];
    const nextChar = csvContent[i + 1];
    
    if (char === '"') {
      if (inQuotes && nextChar === '"') {
        // エスケープされたクォート
        currentField += '"';
        i++; // 次の文字をスキップ
      } else {
        // クォートの開始/終了
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      // 区切り文字（クォート外）
      currentRow.push(currentField.trim());
      currentField = '';
    } else if ((char === '\n' || char === '\r') && !inQuotes) {
      // 行の終了（クォート外）
      if (currentField || currentRow.length > 0) {
        currentRow.push(currentField.trim());
        if (currentRow.some(field => field !== '')) {
          result.push(currentRow);
        }
        currentRow = [];
        currentField = '';
      }
      // \r\n の場合は \n をスキップ
      if (char === '\r' && nextChar === '\n') {
        i++;
      }
    } else {
      // 通常の文字（クォート内の改行も含む）
      currentField += char;
    }
  }
  
  // 最後のフィールドと行を追加
  if (currentField || currentRow.length > 0) {
    currentRow.push(currentField.trim());
    if (currentRow.some(field => field !== '')) {
      result.push(currentRow);
    }
  }
  
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
    
    // 各行の列数を正規化（不足している列を空文字で補完）
    const maxColumnsInData = Math.max(...csvData.map(row => row.length));
    const maxAllowedColumns = 100;
    const firstRowColumns = csvData[0].length;
    
    // 最大列数チェック
    if (maxColumnsInData > maxAllowedColumns) {
      throw new Error(`列数が制限を超えています: ${maxColumnsInData}列 (最大: ${maxAllowedColumns}列)`);
    }
    const inconsistentRows = [];
    
    for (let i = 0; i < csvData.length; i++) {
      const row = csvData[i];
      if (row.length !== maxColumnsInData) {
        inconsistentRows.push(i + 1); // 1ベースの行番号
        
        // 不足している列を空文字で補完
        while (row.length < maxColumnsInData) {
          row.push('');
        }
      }
    }
    
    if (inconsistentRows.length > 0) {
      const maxReportRows = 5;
      const reportRows = inconsistentRows.slice(0, maxReportRows);
      const moreRows = inconsistentRows.length > maxReportRows ? ` (他${inconsistentRows.length - maxReportRows}行)` : '';
      
      Logger.log(`[csvProcessor.gs] 警告: 列数が不一致の行がありましたが、自動補完しました: ${reportRows.join(', ')}行目${moreRows}`);
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