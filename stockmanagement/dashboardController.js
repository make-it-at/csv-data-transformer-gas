// dashboardController.gs
// Version 1.3
// Last updated: 2025-03-02
// Changes: Fixed yield calculation logic

function showDashboard() {
  var template = HtmlService.createTemplateFromFile('dashboardView')
    .evaluate()
    .setWidth(1600)
    .setHeight(1000)
    .setTitle('ポートフォリオダッシュボード');
  SpreadsheetApp.getUi().showModalDialog(template, 'ポートフォリオダッシュボード');
}

function getDashboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main');
  if (!sheet) {
    throw new Error('Mainシートが見つかりません');
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('データが見つかりません');
  }

  // 2行目（インデックス1）をヘッダーとして使用
  var headers = data[1];
  
  // 必要な列のインデックスを取得
  var columnIndexes = {
    code: headers.indexOf('コード'),
    name: headers.indexOf('銘柄名'),
    industry: headers.indexOf('業種17分類'),
    value: headers.indexOf('評価額'),
    purchase_price: headers.indexOf('購入価格'),
    dividend: headers.indexOf('受取配当金'),
    shares: headers.indexOf('保有数'),
    sensitivity: headers.indexOf('景気感応度'),
    accountType: headers.indexOf('口座種別'),
    productType: headers.indexOf('商品種別')
  };

  // データ列の存在確認とデバッグ情報
  console.log('ヘッダー行:', headers);
  console.log('取得したカラムインデックス:', columnIndexes);

  // 必須列の存在確認
  if (Object.values(columnIndexes).some(index => index === -1)) {
    console.error('見つからない列があります:', columnIndexes);
    throw new Error('必要なカラムが見つかりません。シートの列名を確認してください。');
  }

  // インデックスの検証
  validateColumnIndexes(columnIndexes, headers);

  var industryData = {};
  var stockData = [];
  var totalValue = 0;
  var totalDividend = 0;
  var totalPurchasePrice = 0;
  var sensitivityData = {};
  var productTypeData = {};

  // データの集計（ヘッダー行をスキップ、3行目から開始）
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    if (!row || row.every(cell => !cell)) continue; // 空行をスキップ
    
    // 各行のデータを抽出
    var stockInfo = {
      code: row[columnIndexes.code],
      name: row[columnIndexes.name] || '不明',
      industry: row[columnIndexes.industry] || '未分類',
      value: parseFloat(row[columnIndexes.value]) || 0,
      purchase_price: parseFloat(row[columnIndexes.purchase_price]) || 0,
      dividend: parseFloat(row[columnIndexes.dividend]) || 0,
      shares: parseFloat(row[columnIndexes.shares]) || 0,
      sensitivity: row[columnIndexes.sensitivity] || '未分類',
      accountType: row[columnIndexes.accountType] || '未分類',
      productType: row[columnIndexes.productType] || '未分類'
    };

    // 評価額が0以下のデータはスキップ
    if (stockInfo.value <= 0) continue;

    // 業種データの集計
    if (!industryData[stockInfo.industry]) {
      industryData[stockInfo.industry] = { value: 0, stocks: {} };
    }
    industryData[stockInfo.industry].value += stockInfo.value;
    industryData[stockInfo.industry].stocks[stockInfo.name] = stockInfo.value;

    // 景気感応度データの集計
    if (!sensitivityData[stockInfo.sensitivity]) {
      sensitivityData[stockInfo.sensitivity] = 0;
    }
    sensitivityData[stockInfo.sensitivity] += stockInfo.value;

    // 商品種別データの集計
    if (!productTypeData[stockInfo.productType]) {
      productTypeData[stockInfo.productType] = 0;
    }
    productTypeData[stockInfo.productType] += stockInfo.value;

    // 合計の計算
    totalValue += stockInfo.value;
    totalDividend += stockInfo.dividend;
    totalPurchasePrice += stockInfo.purchase_price;

    // 株式データ配列に追加
    stockData.push(stockInfo);
  }

  // 簿価利回りの計算 (年間配当金÷購入価格×100)
  var bookYield = totalPurchasePrice > 0 ? (totalDividend / totalPurchasePrice) * 100 : 0;

  // 現物利回りの計算 (年間配当金÷時価評価額×100)
  var currentYield = totalValue > 0 ? (totalDividend / totalValue) * 100 : 0;

  return {
    industryData: industryData,
    stockData: stockData,
    totalValue: totalValue,
    totalDividend: totalDividend,
    totalPurchasePrice: totalPurchasePrice,
    bookYield: bookYield,
    currentYield: currentYield,
    sensitivityData: sensitivityData,
    productTypeData: productTypeData
  };
}

function validateColumnIndexes(indexes, headers) {
  const requiredColumns = [
    'コード', '銘柄名', '業種17分類', '評価額', '購入価格', '受取配当金', '保有数',
    '景気感応度', '口座種別', '商品種別'
  ];
  
  const missingColumns = requiredColumns.filter(col => 
    headers.indexOf(col) === -1
  );
  
  if (missingColumns.length > 0) {
    console.error('利用可能なヘッダー:', headers);
    console.error('見つからない列:', missingColumns);
    throw new Error('必要な列が見つかりません: ' + missingColumns.join(', '));
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}