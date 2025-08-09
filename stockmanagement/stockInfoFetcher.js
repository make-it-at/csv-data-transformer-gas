// stockInfoFetcher.gs
// Version 1.0
// Last updated: 2025-03-02

const StockInfoFetcher = {
  async fetchStockData(stockCodes, keywordMap) {
    const throttledFetch = AsyncUtils.throttle(
      this.fetchSingleStock.bind(this),
      2
    );
    
    return await AsyncUtils.processInBatches(
      stockCodes,
      10,
      async batch => {
        return await Promise.all(
          batch.map(code => throttledFetch(code, keywordMap))
        );
      },
      { maxConcurrent: 3 }
    );
  },

  async fetchSingleStock(code, keywordMap) {
    return await AsyncUtils.withRetry(async () => {
      const result = {};
      
      for (const [keyword, config] of Object.entries(keywordMap)) {
        if (keyword === 'コード') continue;
        
        const url = config.baseUrl.replace('{code}', code);
        const html = await UrlFetchApp.fetch(url).getContentText();
        const value = this.extractValue(html, config.regex, keyword);
        
        result[keyword] = value;
      }
      
      return {
        code,
        data: result,
        timestamp: new Date().toISOString()
      };
    }, {
      maxAttempts: 3,
      baseDelay: 2000
    });
  },

  extractValue(html, regex, keyword) {
    const match = regex.exec(html);
    if (!match || !match[1]) return null;

    const value = match[1].trim();
    return this.parseValue(value, keyword);
  },

  parseValue(value, keyword) {
    if (['配当金', '株価', 'PER', 'PBR'].includes(keyword)) {
      const numericValue = value.replace(/[^\d.]/g, '');
      return isNaN(numericValue) ? null : parseFloat(numericValue);
    }
    return value;
  }
};