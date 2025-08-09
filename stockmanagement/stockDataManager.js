// stockDataManager.gs
// Version 1.5
// Last updated: 2025-03-02
// Changes: Added support for multi-row headers and improved header detection

var StockDataManager = {
  processOrders: function(ordersData) {
    if (!ordersData || !Array.isArray(ordersData) || ordersData.length === 0) {
      throw new Error('注文データが見つかりません。注文一覧シートを確認してください。');
    }

    // 最初の3行を確認してヘッダー行を特定
    let headerRowIndex = -1;
    let headerValues = null;
    
    for (let i = 0; i < Math.min(3, ordersData.length); i++) {
      const row = ordersData[i];
      if (row.includes('コード') || row.includes('code')) {
        headerRowIndex = i;
        headerValues = row;
        break;
      }
    }

    if (headerRowIndex === -1 || !headerValues) {
      throw new Error('ヘッダー行が見つかりません。「コード」列の存在を確認してください。');
    }

    LoggerManager.debug('ヘッダー行を検出', {
      rowIndex: headerRowIndex + 1,
      headers: headerValues
    });

    // ヘッダー行のカラム名を取得
    const columnMap = {
      code: ['コード', 'code'],
      type: ['取引種類', 'type', '取引区分'],
      num: ['数量', 'num', '取引数量'],
      unitPrice: ['単価', 'price', '取引単価', '約定単価'],
      accountType: ['口座種別', 'accountType', '口座区分'],
      productType: ['商品種別', 'productType', '商品区分'],
      accountHolder: ['口座名義人', 'accountHolder', '名義人']
    };

    var columnIndices = {};
    
    // 各カラムについて、複数の可能な名称をチェック
    Object.entries(columnMap).forEach(([key, possibleNames]) => {
      const index = possibleNames.reduce((found, name) => {
        if (found !== -1) return found;
        return headerValues.findIndex(header => 
          header && header.toString().trim().toLowerCase() === name.toLowerCase()
        );
      }, -1);
      columnIndices[key] = index;
    });

    // カラムの存在チェック
    const missingColumns = Object.entries(columnIndices)
      .filter(([_, value]) => value === -1)
      .map(([key, _]) => columnMap[key][0]);

    if (missingColumns.length > 0) {
      const availableColumns = headerValues
        .filter(Boolean)
        .map(header => `"${header}"`)
        .join(', ');
        
      throw new Error(
        `必要なカラムが見つかりません: ${missingColumns.join(', ')}\n` +
        `利用可能なカラム: ${availableColumns}`
      );
    }

    LoggerManager.debug('カラムインデックス', columnIndices);

    var stockData = {};
    LoggerManager.info('処理開始', { 
      totalRecords: ordersData.length - (headerRowIndex + 1)
    });

    // ヘッダー行の次の行から処理開始
    for (var i = headerRowIndex + 1; i < ordersData.length; i++) {
      var row = ordersData[i];
      if (!row || row.every(cell => !cell)) continue;

      var code = row[columnIndices.code];
      var type = row[columnIndices.type];
      var numStr = String(row[columnIndices.num]).replace(/,/g, '');
      var num = parseFloat(numStr);
      var unitPrice = row[columnIndices.unitPrice];
      var accountType = (row[columnIndices.accountType] || "").toString().trim();
      var productType = (row[columnIndices.productType] || "").toString().trim();
      var accountHolder = (row[columnIndices.accountHolder] || "").toString().trim();

      if (!code || !type || isNaN(num) || !unitPrice || !productType) {
        LoggerManager.debug('無効なデータ行をスキップ', {
          rowIndex: i + 1,
          code,
          type,
          num,
          unitPrice,
          productType
        });
        continue;
      }

      if (typeof unitPrice === 'string') {
        unitPrice = unitPrice.replace(/[¥$,]/g, '');
      }
      unitPrice = parseFloat(unitPrice);

      if (isNaN(unitPrice)) {
        LoggerManager.debug('無効な単価をスキップ', {
          rowIndex: i + 1,
          unitPrice
        });
        continue;
      }

      var key = `${code}|${accountType}|${accountHolder}`;
      
      if (!stockData[key]) {
        stockData[key] = {
          code: code,
          totalNum: 0,
          totalCost: 0,
          accountType: accountType || "未分類",
          productType: productType,
          accountHolder: accountHolder || "未設定",
          isUSD: productType === 'US' || productType === 'MMF',
          isToshin: productType === 'FUND',
          isMMF: productType === 'MMF'
        };
      }

      var amount = (productType === 'FUND' || productType === 'MMF') ? 
        (num * unitPrice / 10000) : (num * unitPrice);

      if (type === "購入") {
        stockData[key].totalNum += num;
        stockData[key].totalCost += amount;
      } else if (type === "売却") {
        if (stockData[key].totalNum > 0) {
          var avgCost = stockData[key].totalCost / stockData[key].totalNum;
          stockData[key].totalNum -= num;
          stockData[key].totalCost -= avgCost * num;
        }
      }
    }

    // 保有数が0以下のデータを除外
    Object.keys(stockData).forEach(key => {
      if (stockData[key].totalNum <= 0) {
        delete stockData[key];
      }
    });

    LoggerManager.info('処理完了', { 
      validRecords: Object.keys(stockData).length 
    });
    
    return stockData;
  }
};