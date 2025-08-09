// dataFetcher.gs
// Version 1.2
// Last updated: 2025-03-02

var DataFetcher = {
  // 基本的なHTMLフェッチ用オプション
  FETCH_OPTIONS: { 
    muteHttpExceptions: true,
    followRedirects: true,
    timeout: 30000
  },

  async fetchHtml(url) {
    try {
      LoggerManager.debug('Fetching URL:', url);
      const startTime = new Date().getTime();
      
      const response = await UrlFetchApp.fetch(url, this.FETCH_OPTIONS);
      const contentText = response.getContentText();
      
      const endTime = new Date().getTime();
      LoggerManager.debug('Fetch completed', {
        url: url,
        status: response.getResponseCode(),
        duration: `${endTime - startTime}ms`,
        contentLength: contentText.length
      });

      if (response.getResponseCode() !== 200) {
        throw new Error(`HTTP Error: ${response.getResponseCode()}`);
      }

      return contentText;
    } catch (e) {
      LoggerManager.error('Error fetching HTML from ' + url, e);
      throw e;
    }
  },

  extractInfo(html, regex, keyword) {
    try {
      const match = regex.exec(html);
      if (match && match[1]) {
        let result = match[1];
        
        // 特定のキーワードの場合、数値のみを抽出
        if (['配当金', '株価', '株価診断', '個人予想', 'アナリスト評価', 'PER', 'PBR'].includes(keyword)) {
          result = result.replace(/[円倍,%％]/g, '');
        }
        
        const cleanedResult = result.trim();
        LoggerManager.debug('Extracted info', {
          keyword: keyword,
          rawValue: match[1],
          cleanedValue: cleanedResult
        });
        
        return cleanedResult;
      }
      
      // キーワードに応じてデフォルト値を返す
      return keyword === '配当金' || keyword === '分配金' ? '0' : '-';
    } catch (e) {
      LoggerManager.error('Error extracting info for ' + keyword, e);
      return keyword === '配当金' || keyword === '分配金' ? '0' : '-';
    }
  },

  async getStockPriceGoogle(code) {
    try {
      const url = 'https://www.google.com/finance/quote/' + code + ':TYO';
      const html = await this.fetchHtml(url);
      
      const price = Parser.data(html)
        .from('<div class="YMlKec fxKbKc">\\xA5')
        .to('</div>')
        .build();

      LoggerManager.debug('Google Finance price fetched', {
        code: code,
        price: price
      });

      return price;
    } catch (e) {
      LoggerManager.error('Error getting stock price from Google:', e);
      throw e;
    }
  },

  async getPriceToshin(code) {
    try {
      const url = 'https://www.rakuten-sec.co.jp/web/fund/detail/?ID=' + code;
      const html = await this.fetchHtml(url);
      
      const price = Parser.data(html)
        .from('<span class="value-01">')
        .to('</span>')
        .build();

      LoggerManager.debug('Toshin price fetched', {
        code: code,
        price: price
      });

      return price;
    } catch (e) {
      LoggerManager.error('Error getting toshin price:', e);
      throw e;
    }
  },

  validateResponse(response) {
    const validStatus = [200, 201, 202];
    if (!validStatus.includes(response.getResponseCode())) {
      throw new Error(`Invalid response status: ${response.getResponseCode()}`);
    }
    
    const contentText = response.getContentText();
    if (!contentText) {
      throw new Error('Empty response content');
    }
    
    return contentText;
  },

  handleError(error, context) {
    LoggerManager.error('DataFetcher error:', {
      context: context,
      error: error.message,
      stack: error.stack
    });
    throw error;
  }
};