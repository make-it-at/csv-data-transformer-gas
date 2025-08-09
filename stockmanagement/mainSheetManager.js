// mainSheetManager.gs
// Version 1.7
// Last updated: 2024-05-01
// Changes: 列インデックス計算の改善とエラーハンドリングの強化

var MainSheetManager = {
  updateSheet: function(sheetMain, stockData, columnIndices) {
    try {
      const startRow = 3;
      const lastRowMain = sheetMain.getLastRow();
      
      // 既存データのクリア
      if (lastRowMain >= startRow) {
        const rangeToClear = sheetMain.getRange(startRow, 1, lastRowMain - startRow + 1, sheetMain.getLastColumn());
        rangeToClear.clearContent();
        LoggerManager.debug('既存データをクリア', {
          startRow: startRow,
          rowCount: lastRowMain - startRow + 1
        });
      }

      // メインデータの生成
      const mainData = this.createMainData(stockData, columnIndices, sheetMain.getLastColumn());
      
      if (mainData.length > 0) {
        this.writeMainData(sheetMain, mainData, startRow, columnIndices);
        LoggerManager.info('シートの更新完了', { rowsWritten: mainData.length });
      }

      UIManager.updateProgress(mainData.length, mainData.length, 'Main シートの更新が完了しました。更新行数: ' + mainData.length);
    } catch (e) {
      LoggerManager.error('Error in updateSheet:', e);
      throw e;
    }
  },

  createMainData: function(stockData, columnIndices, lastColumn) {
    try {
      const mainData = [];
      
      // 銘柄コードでソートするためのキー配列を作成
      const sortedKeys = Object.keys(stockData).sort((a, b) => {
        const codeA = stockData[a].code.toString();
        const codeB = stockData[b].code.toString();
        return codeA.localeCompare(codeB);
      });
      
      sortedKeys.forEach((key, index) => {
        const stock = stockData[key];
        const rowData = new Array(lastColumn).fill('');
        
        // 必須フィールドの設定
        this.setRequiredFields(rowData, stock, index + 1, columnIndices);
        mainData.push(rowData);
      });

      return mainData;
    } catch (e) {
      LoggerManager.error('Error in createMainData:', e);
      throw e;
    }
  },

  setRequiredFields: function(rowData, stock, rowNum, columnIndices) {
    const fieldsToSet = {
      'No.': rowNum,
      'コード': stock.code.toString(),
      '保有数': Math.floor(stock.totalNum), // 小数点以下切り捨て
      '購入価格': stock.totalCost,
      '購入単価': stock.totalCost / stock.totalNum,
      '口座種別': stock.accountType || "未分類",
      '商品種別': stock.productType || "未分類",
      '口座名義人': stock.accountHolder || "未設定"
    };

    Object.entries(fieldsToSet).forEach(([field, value]) => {
      if (columnIndices[field]) {
        rowData[columnIndices[field] - 1] = value;
      }
    });
  },

  writeMainData: function(sheetMain, mainData, startRow, columnIndices) {
    try {
      // データの一括書き込み
      const range = sheetMain.getRange(startRow, 1, mainData.length, mainData[0].length);
      range.setValues(mainData);
      
      // 各行のフォーマット設定
      for (let i = 0; i < mainData.length; i++) {
        const currentRow = startRow + i;
        this.formatRow(sheetMain, currentRow, columnIndices);
      }
    } catch (e) {
      LoggerManager.error('Error in writeMainData:', e);
      throw e;
    }
  },

  formatRow: function(sheet, row, columnIndices) {
    try {
      // 数値フォーマットの設定
      this.setNumberFormats(sheet, row, columnIndices);
      
      // 株価計算式の設定
      if (columnIndices['株価'] && columnIndices['コード']) {
        const code = sheet.getRange(row, columnIndices['コード']).getValue();
        const formula = this.createPriceFormula(row, code);
        sheet.getRange(row, columnIndices['株価']).setFormula(formula);
      }
    } catch (e) {
      LoggerManager.error('Error in formatRow:', e);
    }
  },

  setNumberFormats: function(sheet, row, columnIndices) {
    const formats = {
      '保有数': '#,##0',  // 小数点なし
      '購入価格': '¥#,##0.00',  // 通貨表記、小数点2位
      '購入単価': '¥#,##0.00',  // 通貨表記、小数点2位
      '株価': '¥#,##0.00'  // 通貨表記、小数点2位
    };

    Object.entries(formats).forEach(([column, format]) => {
      if (columnIndices[column]) {
        sheet.getRange(row, columnIndices[column]).setNumberFormat(format);
      }
    });
  },

  createPriceFormula: function(row, code) {
    if (typeof code === 'string') {
      if (code.startsWith('JP')) {
        return '=IF(ISBLANK(B' + row + '), "", STOCKPRICEJP("TOSHIN", B' + row + '))';
      }
      if (code === 'FMJXX0000000') {
        return '=E' + row;
      }
      if (['BND', 'HDV', 'SPYD', 'TLT', 'VIG'].includes(code)) {
        return '=IF(ISBLANK(B' + row + '), "", GOOGLEFINANCE(B' + row + '))';
      }
    }
    return '=IF(ISBLANK(B' + row + '),"",STOCKPRICEJP("JP", B' + row + '))';
  },
  
  /**
   * メインシートの列インデックスを取得する
   * キャッシュから取得するか、存在しない場合は計算して返す
   * @returns {Object} 列インデックス
   */
  getColumnIndices: function() {
    try {
      // キャッシュから取得を試みる
      let columnIndices = CacheManager.getColumnIndices(SHEET_NAMES.MAIN);
      
      // キャッシュに存在しない場合は計算
      if (!columnIndices) {
        LoggerManager.debug('列インデックスをキャッシュから取得できませんでした。再計算します。');
        columnIndices = this.calculateColumnIndices();
        
        // キャッシュに保存
        if (columnIndices) {
          CacheManager.setColumnIndices(SHEET_NAMES.MAIN, columnIndices);
        }
      }
      
      return columnIndices;
    } catch (e) {
      LoggerManager.error('列インデックスの取得に失敗しました', e);
      return null;
    }
  },
  
  /**
   * メインシートの列インデックスを計算する
   * @returns {Object} 列インデックス
   */
  calculateColumnIndices: function() {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
      if (!sheet) {
        throw new Error('メインシートが見つかりません');
      }
      
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const indices = {};
      
      // 必須列のマッピング
      const requiredColumns = {
        'code': ['銘柄コード', 'コード', '証券コード'],
        'name': ['銘柄名', '名称', '証券名'],
        'productType': ['種別', '商品種別', 'タイプ'],
        'price': ['株価', '現在値', '価格'],
        'priceDate': ['株価取得日', '取得日', '日付']
      };
      
      // 列名からインデックスを取得
      headerRow.forEach((header, index) => {
        if (header) {
          // 0ベースのインデックスを格納
          indices[header] = index;
          
          // 必須列の別名マッピング
          for (const [key, aliases] of Object.entries(requiredColumns)) {
            if (aliases.includes(header)) {
              indices[key] = index;
            }
          }
        }
      });
      
      // 必須列が見つからない場合はデフォルト値を設定
      if (indices.code === undefined) indices.code = 0; // デフォルトは最初の列
      if (indices.name === undefined) indices.name = 1; // デフォルトは2列目
      if (indices.productType === undefined) indices.productType = 2; // デフォルトは3列目
      if (indices.price === undefined) indices.price = 3; // デフォルトは4列目
      if (indices.priceDate === undefined) indices.priceDate = 4; // デフォルトは5列目
      
      LoggerManager.debug('列インデックスを計算しました', { indices: indices });
      return indices;
    } catch (e) {
      LoggerManager.error('列インデックスの計算に失敗しました', e);
      // エラー時のデフォルト値を返す
      return {
        code: 0,
        name: 1,
        productType: 2,
        price: 3,
        priceDate: 4
      };
    }
  }
};