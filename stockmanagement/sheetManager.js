// sheetManager.gs
// Version 1.3
// Last updated: 2025-03-02

var SheetManager = {
  getUntilBlank: function(sheet, startRow, startCol, numCols) {
    try {
      let result = [];
      let row = startRow;
      
      while (true) {
        const range = sheet.getRange(row, startCol, 1, numCols);
        const values = range.getValues()[0];
        if (values.every(cell => cell === '')) break;
        result.push(values);
        row++;
      }
      return result;
    } catch (e) {
      LoggerManager.error('Error in getUntilBlank:', e);
      throw e;
    }
  },

  getKeywords: function(settingSheet, frequency) {
    if (!settingSheet) {
      LoggerManager.error('設定シートが見つかりません');
      return [];
    }

    try {
      const lastRow = settingSheet.getLastRow();
      const settingRange = settingSheet.getRange('A2:E' + lastRow);
      const settings = settingRange.getValues();
      
      const keywords = settings
        .filter(row => row[4] === frequency)
        .map(row => row[0])
        .filter(keyword => typeof keyword === 'string' && keyword.trim() !== '');
      
      LoggerManager.debug('設定シートから取得したキーワード', {
        frequency: frequency,
        keywords: keywords,
        totalKeywords: keywords.length
      });

      return keywords;
    } catch (e) {
      LoggerManager.error('Error in getKeywords:', e);
      return [];
    }
  },

  getKeywordsFromRow: function(sheet, row) {
    if (!sheet) throw new Error('シートが見つかりません');

    try {
      const requiredKeywords = [
        'No.', 'コード', '保有数', '購入価格', '購入単価', 
        '口座種別', '株価', '商品種別'
      ];
    
      const keywordRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
      const keywords = keywordRange.getValues()[0]
        .map(value => value.toString().trim())
        .filter(value => value !== '');

      const missingKeywords = requiredKeywords.filter(keyword => 
        !keywords.includes(keyword)
      );
      
      if (missingKeywords.length > 0) {
        LoggerManager.error('必須キーワードが見つかりません', {
          missingKeywords: missingKeywords,
          availableKeywords: keywords
        });
        throw new Error('必須キーワードが見つかりません: ' + missingKeywords.join(', '));
      }

      LoggerManager.debug('取得したキーワード', { keywords: keywords });
      return keywords;
    } catch (e) {
      LoggerManager.error('Error in getKeywordsFromRow:', e);
      throw e;
    }
  },

  getColumnIndices: function(sheet, keywords) {
    if (!sheet) throw new Error('シートが見つかりません。シート名を確認してください。');

    try {
      const requiredColumns = [
        'No.', 'コード', '保有数', '購入価格', '購入単価', 
        '口座種別', '株価', '商品種別'
      ];
      
      const headerRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      LoggerManager.debug('ヘッダー行', { headers: headerRow });
      
      const columnIndices = {};
      
      for (const column of requiredColumns) {
        const index = headerRow.indexOf(column);
        if (index === -1) {
          LoggerManager.error('必須列が見つかりません', {
            missingColumn: column,
            availableColumns: headerRow
          });
          throw new Error('必須列「' + column + '」が見つかりません');
        }
        columnIndices[column] = index + 1;
      }
      
      keywords.forEach(keyword => {
        if (!requiredColumns.includes(keyword)) {
          const colIndex = headerRow.indexOf(keyword);
          if (colIndex !== -1) {
            columnIndices[keyword] = colIndex + 1;
          } else {
            LoggerManager.warn('キーワードがヘッダーに見つかりません', {
              keyword: keyword,
              headers: headerRow
            });
          }
        }
      });
      
      LoggerManager.debug('列インデックス', { indices: columnIndices });
      return columnIndices;
    } catch (e) {
      LoggerManager.error('Error in getColumnIndices:', e);
      throw e;
    }
  },

  getOrderSheetColumns: function(sheet, requiredColumns) {
    if (!sheet) throw new Error('シートが見つかりません。シート名を確認してください。');
    
    try {
      const lastColumn = sheet.getLastColumn();
      const headerRows = sheet.getRange(1, 1, 3, lastColumn).getValues();
      LoggerManager.debug('最初の3行', { headerRows: headerRows });
      
      let headerRow;
      let headerRowIndex = -1;
      
      for (let i = 0; i < headerRows.length; i++) {
        if (headerRows[i].includes("コード")) {
          headerRow = headerRows[i];
          headerRowIndex = i;
          break;
        }
      }
      
      if (headerRowIndex === -1) {
        LoggerManager.error('ヘッダー行が見つかりません', {
          firstThreeRows: headerRows
        });
        throw new Error('ヘッダー行が見つかりません。「コード」列の存在を確認してください。');
      }
      
      LoggerManager.debug('見つかったヘッダー行', {
        rowIndex: headerRowIndex + 1,
        headers: headerRow
      });
      
      const columnIndices = {};
      for (const column of requiredColumns) {
        const index = headerRow.indexOf(column);
        if (index === -1) {
          LoggerManager.error('必要な列が見つかりません', {
            missingColumn: column,
            availableColumns: headerRow
          });
          throw new Error('必要な列が見つかりません: ' + column + '\n利用可能な列: ' + headerRow.join(', '));
        }
        columnIndices[column] = index;
      }
      
      return { 
        indices: columnIndices, 
        headerRowIndex: headerRowIndex 
      };
    } catch (e) {
      LoggerManager.error('Error in getOrderSheetColumns:', e);
      throw e;
    }
  },

  validateSheetExists: function(sheet, sheetName) {
    if (!sheet) {
      LoggerManager.error('シートが見つかりません', { sheetName: sheetName });
      throw new Error(sheetName + 'シートが見つかりません');
    }
    return sheet;
  },

  getLastRowWithData: function(sheet, column) {
    try {
      const lastRow = sheet.getLastRow();
      const dataRange = sheet.getRange(1, column, lastRow, 1).getValues();
      
      for (let i = lastRow - 1; i >= 0; i--) {
        if (dataRange[i][0] !== '') {
          return i + 1;
        }
      }
      return 0;
    } catch (e) {
      LoggerManager.error('Error in getLastRowWithData:', e);
      throw e;
    }
  },

  clearRange: function(sheet, startRow, startCol, numRows, numCols) {
    try {
      if (numRows > 0 && numCols > 0) {
        sheet.getRange(startRow, startCol, numRows, numCols).clearContent();
        LoggerManager.debug('Range cleared successfully', {
          sheet: sheet.getName(),
          startRow: startRow,
          startCol: startCol,
          numRows: numRows,
          numCols: numCols
        });
      }
    } catch (e) {
      LoggerManager.error('Error clearing range:', e);
      throw e;
    }
  },

  backupSheet: function(sheet, backupName) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const backupSheet = ss.insertSheet(backupName);
      
      const sourceData = sheet.getDataRange();
      const values = sourceData.getValues();
      const numRows = values.length;
      const numCols = values[0].length;
      
      backupSheet.getRange(1, 1, numRows, numCols).setValues(values);
      
      LoggerManager.debug('Sheet backup created', {
        originalSheet: sheet.getName(),
        backupName: backupName,
        rowsCopied: numRows,
        columnsCopied: numCols
      });
      
      return backupSheet;
    } catch (e) {
      LoggerManager.error('Error creating sheet backup:', e);
      throw e;
    }
  }
};

// グローバル関数
function getKeywordsFromRow(sheet, row) {
  return SheetManager.getKeywordsFromRow(sheet, row);
}