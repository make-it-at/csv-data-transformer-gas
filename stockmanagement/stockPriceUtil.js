// stockPriceUtil.gs
// Version 1.7
// Last updated: 2024-05-01
// Changes: 投資信託の価格取得処理を改善、新しい取得パターンを追加

var StockPriceUtil = {
  FETCH_OPTIONS: {
    muteHttpExceptions: true,
    followRedirects: true,
    validateHttpsCertificates: true,
    timeout: 30000
  },

  STOCKPRICEJP: function(torihiki_code, shoken_code) {
    var param = torihiki_code;
    var error = '';
    
    try {
      if (String(param).length <= 0) {
        error = '取引コードが無し!';
        throw error;
      } 
      
      if (String(param).length > 0 &&
          !((param == 'JP' && String(shoken_code).length == 4) ||
            (param == 'TOSHIN' &&
              String(shoken_code).length == 12 &&
              String(shoken_code).startsWith('JP')))) {
        error = '取引、証券コードが正しくありません';
        throw error;
      }

      if ('JP' == param) {
        return this.updateStockPrices(shoken_code);
      } else if ('TOSHIN' == param) {
        return this.updateToshinPrices(shoken_code);
      }
      return '取引コードが不正！';

    } catch (e) {
      LoggerManager.error('StockPriceUtil error in STOCKPRICEJP', e);
      return error || e.message;
    }
  },

  updateStockPrices: function(stockCode) {
    try {
      // キャッシュから株価を取得
      if (typeof CacheManager !== 'undefined') {
        const cachedPrice = CacheManager.getStockPrice(stockCode);
        if (cachedPrice && cachedPrice.value) {
          LoggerManager.debug('Stock price from cache', {
            code: stockCode,
            price: cachedPrice.value
          });
          return cachedPrice.value;
        }
      }
      
      var url = 'https://www.google.com/finance/quote/' + stockCode + ':TYO';
      var response = UrlFetchApp.fetch(url, this.FETCH_OPTIONS);
      
      if (response.getResponseCode() !== 200) {
        LoggerManager.error('Failed to fetch stock price', {
          code: stockCode,
          status: response.getResponseCode()
        });
        throw new Error('Failed to fetch stock price');
      }

      var content = response.getContentText();
      
      // 複数の正規表現パターンを試す
      var pricePatterns = [
        /"YMlKec fxKbKc">\¥([\d,]+)/, // 古いパターン
        /"YMlKec">\¥([\d,]+)/, // クラス名が変わった可能性
        /data-last-price="([\d.]+)"/, // data属性を探す
        /<div[^>]*class="[^"]*"[^>]*>\¥([\d,]+)<\/div>/, // より一般的なパターン
        /class="[^"]*price[^"]*"[^>]*>\¥?([\d,]+)/ // priceという単語を含むクラス
      ];
      
      var price = null;
      
      // 各パターンを試す
      for (var i = 0; i < pricePatterns.length; i++) {
        var match = content.match(pricePatterns[i]);
        if (match && match[1]) {
          price = Number(match[1].replace(/,/g, ''));
          LoggerManager.debug('Stock price found with pattern ' + i, {
            code: stockCode,
            price: price
          });
          break;
        }
      }
      
      // 価格が見つからなかった場合は別の方法を試す
      if (!price) {
        // Yahoo Financeを代替ソースとして使用
        try {
          var yahooUrl = 'https://finance.yahoo.co.jp/quote/' + stockCode + '.T';
          var yahooResponse = UrlFetchApp.fetch(yahooUrl, this.FETCH_OPTIONS);
          
          if (yahooResponse.getResponseCode() === 200) {
            var yahooContent = yahooResponse.getContentText();
            
            // 複数の正規表現パターンを試す（現在のYahoo Financeに対応）
            var yahooPatterns = [
              /"price":"([\d,\.]+)"/, // JSON形式（最も確実）
              /<span[^>]*class="stocksDetail__price"[^>]*>([^<]+)<\/span>/, // CSS クラス
              /class="[^"]*price[^"]*"[^>]*>([\d,\.]+)/, // 一般的なpriceクラス
              /_3rXWJKZF _11kV6f2G">([\d,]+)</, // 既存パターン（後方互換）
              /現在値<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/, // 現在値ラベル
              /<span[^>]*>\s*([\d,\.]+)\s*<\/span>/ // 汎用spanパターン
            ];
            
            for (var i = 0; i < yahooPatterns.length; i++) {
              var yahooMatch = yahooContent.match(yahooPatterns[i]);
              if (yahooMatch && yahooMatch[1]) {
                // カンマやピリオドを除去して数値に変換
                var cleanPrice = yahooMatch[1].replace(/,/g, '');
                if (!isNaN(cleanPrice) && cleanPrice !== '') {
                  price = Number(cleanPrice);
                  LoggerManager.info('Yahoo Financeから株価取得成功', {
                    code: stockCode,
                    price: price,
                    pattern: i,
                    matchedValue: yahooMatch[1]
                  });
                  break;
                }
              }
            }
            
            if (!price) {
              LoggerManager.warn('Yahoo Financeでパターンマッチせず', {
                code: stockCode,
                contentLength: yahooContent.length,
                patterns: yahooPatterns.length
              });
            }
          } else {
            LoggerManager.warn('Yahoo Finance HTTPエラー', {
              code: stockCode,
              status: yahooResponse.getResponseCode()
            });
          }
        } catch (yahooError) {
          LoggerManager.error('Yahoo Finance取得エラー', {
            code: stockCode,
            error: yahooError.message
          });
        }
      }
      
      // 4. 楽天証券から取得を試みる（新しいソースを追加）
      if (!price) {
        try {
          const rakutenUrl = 'https://www.rakuten-sec.co.jp/web/market/search/quote.html?ric=' + stockCode + '.T';
          const rakutenResponse = UrlFetchApp.fetch(rakutenUrl, this.FETCH_OPTIONS);
          
          if (rakutenResponse.getResponseCode() === 200) {
            const rakutenContent = rakutenResponse.getContentText();
            // 楽天証券の株価表示パターンを複数試す
            const rakutenPatterns = [
              /現在値<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /現在値<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /class="[^"]*price[^"]*"[^>]*>([\d,]+)</,
              />([\d,]+)円</,
              /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /前日終値<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/
            ];
            
            for (let i = 0; i < rakutenPatterns.length; i++) {
              const rakutenMatch = rakutenContent.match(rakutenPatterns[i]);
              if (rakutenMatch && rakutenMatch[1]) {
                price = Number(rakutenMatch[1].replace(/,/g, ''));
                LoggerManager.info('楽天証券から株価取得成功', {
                  code: stockCode,
                  price: price,
                  pattern: i
                });
                break;
              }
            }
          }
        } catch (rakutenError) {
          LoggerManager.warn('楽天証券からの取得に失敗', {
            code: stockCode,
            error: rakutenError.message
          });
        }
      }
      
      // 5. Japan-REIT.comから取得を試みる（REIT専用ソース）
      if (!price) {
        try {
          // REITの場合は特殊なソースを試す
          const jreitUrl = 'https://j-reit.jp/brand/detail/' + stockCode + '/';
          const jreitResponse = UrlFetchApp.fetch(jreitUrl, this.FETCH_OPTIONS);
          
          if (jreitResponse.getResponseCode() === 200) {
            const jreitContent = jreitResponse.getContentText();
            // Japan-REITの株価表示パターンを複数試す
            const jreitPatterns = [
              /現在値<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /現在値<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /基準価額<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /前日終値<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /前日終値<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /<span[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)<\/span>/,
              /<div[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)<\/div>/,
              />([\d,]+)円</
            ];
            
            for (let i = 0; i < jreitPatterns.length; i++) {
              const jreitMatch = jreitContent.match(jreitPatterns[i]);
              if (jreitMatch && jreitMatch[1]) {
                price = Number(jreitMatch[1].replace(/,/g, ''));
                LoggerManager.info('Japan-REITから株価取得成功', {
                  code: stockCode,
                  price: price,
                  pattern: i
                });
                break;
              }
            }
          }
        } catch (jreitError) {
          LoggerManager.warn('Japan-REITからの取得に失敗', {
            code: stockCode,
            error: jreitError.message
          });
        }
      }
      
      // 6. REIT.or.jpから取得を試みる（REIT専用ソース）
      if (!price) {
        try {
          // REITの場合は特殊なソースを試す
          const reitOrJpUrl = 'https://j-reit.or.jp/brand/' + stockCode + '/';
          const reitOrJpResponse = UrlFetchApp.fetch(reitOrJpUrl, this.FETCH_OPTIONS);
          
          if (reitOrJpResponse.getResponseCode() === 200) {
            const reitOrJpContent = reitOrJpResponse.getContentText();
            // REIT.or.jpの株価表示パターンを複数試す
            const reitOrJpPatterns = [
              /現在値<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /現在値<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /基準価額<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /前日終値<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /前日終値<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /<span[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)<\/span>/,
              /<div[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)<\/div>/,
              />([\d,]+)円</
            ];
            
            for (let i = 0; i < reitOrJpPatterns.length; i++) {
              const reitOrJpMatch = reitOrJpContent.match(reitOrJpPatterns[i]);
              if (reitOrJpMatch && reitOrJpMatch[1]) {
                price = Number(reitOrJpMatch[1].replace(/,/g, ''));
                LoggerManager.info('REIT.or.jpから株価取得成功', {
                  code: stockCode,
                  price: price,
                  pattern: i
                });
                break;
              }
            }
          }
        } catch (reitOrJpError) {
          LoggerManager.warn('REIT.or.jpからの取得に失敗', {
            code: stockCode,
            error: reitOrJpError.message
          });
        }
      }
      
      // 7. 8593特有のパターン（三菱地所物流リート投資法人）
      if (!price && stockCode === '8593') {
        try {
          // 三菱地所物流リート投資法人の公式サイト
          const mflpUrl = 'https://mel-reit.co.jp/';
          const mflpResponse = UrlFetchApp.fetch(mflpUrl, this.FETCH_OPTIONS);
          
          if (mflpResponse.getResponseCode() === 200) {
            const mflpContent = mflpResponse.getContentText();
            // 公式サイトの株価表示パターンを複数試す
            const mflpPatterns = [
              /投資口価格<[^>]*>[\s\S]*?<[^>]*>([\d,]+)円<\/[^>]*>/,
              /投資口価格[\s\S]*?<[^>]*>([\d,]+)円<\/[^>]*>/,
              /投資口価格[\s\S]*?<[^>]*>([\d,]+)<\/[^>]*>/,
              /<span[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)<\/span>/,
              /<div[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)<\/div>/,
              />([\d,]+)円</
            ];
            
            for (let i = 0; i < mflpPatterns.length; i++) {
              const mflpMatch = mflpContent.match(mflpPatterns[i]);
              if (mflpMatch && mflpMatch[1]) {
                price = Number(mflpMatch[1].replace(/,/g, ''));
                LoggerManager.info('三菱地所物流リート公式サイトから株価取得成功', {
                  code: stockCode,
                  price: price,
                  pattern: i
                });
                break;
              }
            }
          }
        } catch (mflpError) {
          LoggerManager.warn('三菱地所物流リート公式サイトからの取得に失敗', {
            code: stockCode,
            error: mflpError.message
          });
        }
      }
      
      if (!price) {
        var errorMessage = `株価データ取得失敗: ${stockCode}`;
        LoggerManager.error('Price not found after all attempts', {
          code: stockCode,
          attempts: ['Google Finance', 'Yahoo Finance', '楽天証券', 'Japan-REIT', 'REIT.or.jp']
        });
        
        // より具体的なエラーメッセージを返す
        return `価格未取得: ${stockCode}`;
      }

      // キャッシュに保存
      if (typeof CacheManager !== 'undefined') {
        CacheManager.setStockPrice(stockCode, price);
      }

      return price;

    } catch (e) {
      LoggerManager.error('株価取得で予期しないエラー発生', {
        code: stockCode,
        error: e.message,
        stack: e.stack,
        timestamp: new Date().toISOString()
      });
      
      // より詳細なエラーメッセージを返す
      if (e.message.includes('timeout')) {
        return `タイムアウト: ${stockCode}`;
      } else if (e.message.includes('network')) {
        return `ネットワークエラー: ${stockCode}`;
      } else {
        return `システムエラー: ${stockCode}`;
      }
    }
  },

  updateToshinPrices: function(toshinCode) {
    try {
      // キャッシュから価格を取得
      if (typeof CacheManager !== 'undefined') {
        const cachedPrice = CacheManager.getStockPrice(toshinCode);
        if (cachedPrice && cachedPrice.value) {
          LoggerManager.debug('Toshin price from cache', {
            code: toshinCode,
            price: cachedPrice.value
          });
          return cachedPrice.value;
        }
      }
      
      var price = null;
      
      // 1. 楽天証券から取得を試みる
      try {
        var url = 'https://www.rakuten-sec.co.jp/web/fund/detail/?ID=' + toshinCode;
        var response = UrlFetchApp.fetch(url, this.FETCH_OPTIONS);
        
        if (response.getResponseCode() === 200) {
          var content = response.getContentText();
          
          // 複数の正規表現パターンを試す
          var pricePatterns = [
            /<span class="fund-price">([\d,]+)円<\/span>/,
            /<td[^>]*class="[^"]*price[^"]*"[^>]*>([\d,]+)円<\/td>/,
            /基準価額[^<]*<[^>]*>([\d,]+)円/,
            /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
            /基準価額<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
            /class="[^"]*price[^"]*"[^>]*>([\d,]+)</,
            />([\d,]+)円</
          ];
          
          // 各パターンを試す
          for (var i = 0; i < pricePatterns.length; i++) {
            var match = content.match(pricePatterns[i]);
            if (match && match[1]) {
              price = Number(match[1].replace(/,/g, ''));
              LoggerManager.debug('Toshin price found with pattern ' + i, {
                code: toshinCode,
                price: price,
                source: '楽天証券'
              });
              break;
            }
          }
        }
      } catch (rakutenError) {
        LoggerManager.error('Error fetching from Rakuten Securities', {
          code: toshinCode,
          error: rakutenError.message
        });
      }
      
      // 2. モーニングスターから取得を試みる
      if (!price) {
        try {
          var msUrl = 'https://toushin.morningstar.co.jp/FundData/SnapShot.do?fnc=' + toshinCode;
          var msResponse = UrlFetchApp.fetch(msUrl, this.FETCH_OPTIONS);
          
          if (msResponse.getResponseCode() === 200) {
            var msContent = msResponse.getContentText();
            var msPatterns = [
              /基準価額[^<]*<[^>]*>([\d,]+)円/,
              /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /基準価額<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /class="[^"]*price[^"]*"[^>]*>([\d,]+)</,
              />([\d,]+)円</
            ];
            
            for (var j = 0; j < msPatterns.length; j++) {
              var msMatch = msContent.match(msPatterns[j]);
              if (msMatch && msMatch[1]) {
                price = Number(msMatch[1].replace(/,/g, ''));
                LoggerManager.debug('Toshin price found from Morningstar', {
                  code: toshinCode,
                  price: price,
                  pattern: j
                });
                break;
              }
            }
          }
        } catch (msError) {
          LoggerManager.error('Error fetching from Morningstar', {
            code: toshinCode,
            error: msError.message
          });
        }
      }
      
      // 3. 投信協会から取得を試みる
      if (!price) {
        try {
          var toushinUrl = 'https://toushin-lib.fwg.ne.jp/FdsWeb/FDST000000?isinCd=' + toshinCode;
          var toushinResponse = UrlFetchApp.fetch(toushinUrl, this.FETCH_OPTIONS);
          
          if (toushinResponse.getResponseCode() === 200) {
            var toushinContent = toushinResponse.getContentText();
            var toushinPatterns = [
              /基準価額[^<]*<[^>]*>([\d,]+)円/,
              /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /基準価額<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /class="[^"]*price[^"]*"[^>]*>([\d,]+)</,
              />([\d,]+)円</
            ];
            
            for (var k = 0; k < toushinPatterns.length; k++) {
              var toushinMatch = toushinContent.match(toushinPatterns[k]);
              if (toushinMatch && toushinMatch[1]) {
                price = Number(toushinMatch[1].replace(/,/g, ''));
                LoggerManager.debug('Toshin price found from Investment Trust Association', {
                  code: toshinCode,
                  price: price,
                  pattern: k
                });
                break;
              }
            }
          }
        } catch (toushinError) {
          LoggerManager.error('Error fetching from Investment Trust Association', {
            code: toshinCode,
            error: toushinError.message
          });
        }
      }
      
      // 4. Yahoo!ファイナンスから取得を試みる
      if (!price) {
        try {
          // Yahoo!ファイナンスでは投信コードの形式が異なる場合がある
          var yahooCode = toshinCode;
          if (toshinCode.startsWith('JP')) {
            yahooCode = toshinCode.substring(2);
          }
          
          var yahooUrl = 'https://finance.yahoo.co.jp/quote/' + yahooCode;
          var yahooResponse = UrlFetchApp.fetch(yahooUrl, this.FETCH_OPTIONS);
          
          if (yahooResponse.getResponseCode() === 200) {
            var yahooContent = yahooResponse.getContentText();
            
            // 投資信託用の改良されたYahoo Financeパターン
            var yahooPatterns = [
              /"price":"([\d,\.]+)"/, // JSON形式（最も確実）
              /<span[^>]*class="stocksDetail__price"[^>]*>([^<]+)<\/span>/, // CSS クラス
              /基準価額[^<]*<[^>]*>([\d,]+)円/,
              /基準価額<\/th>[\s\S]*?<td[^>]*>([\d,]+)<\/td>/,
              /基準価額<\/dt>[\s\S]*?<dd[^>]*>([\d,]+)<\/dd>/,
              /class="[^"]*price[^"]*"[^>]*>([\d,\.]+)/, // 一般的なpriceクラス
              /_3rXWJKZF _11kV6f2G">([\d,]+)</, // 既存パターン（後方互換）
              />([\d,]+)円</
            ];
            
            for (var l = 0; l < yahooPatterns.length; l++) {
              var yahooMatch = yahooContent.match(yahooPatterns[l]);
              if (yahooMatch && yahooMatch[1]) {
                var cleanPrice = yahooMatch[1].replace(/,/g, '');
                if (!isNaN(cleanPrice) && cleanPrice !== '') {
                  price = Number(cleanPrice);
                  LoggerManager.info('Yahoo Financeから投信価格取得成功', {
                    code: toshinCode,
                    price: price,
                    pattern: l,
                    matchedValue: yahooMatch[1]
                  });
                  break;
                }
              }
            }
            
            if (!price) {
              LoggerManager.warn('Yahoo Finance投信でパターンマッチせず', {
                code: toshinCode,
                contentLength: yahooContent.length
              });
            }
          } else {
            LoggerManager.warn('Yahoo Finance投信 HTTPエラー', {
              code: toshinCode,
              status: yahooResponse.getResponseCode()
            });
          }
        } catch (yahooError) {
          LoggerManager.error('Error fetching from Yahoo Finance', {
            code: toshinCode,
            error: yahooError.message
          });
        }
      }
      
      if (!price) {
        LoggerManager.error('Toshin price not found for code: ' + toshinCode, {
          code: toshinCode,
          attempts: 4
        });
        // 価格が見つからない場合でもエラーではなく0を返す
        return 0;
      }

      // キャッシュに保存
      if (typeof CacheManager !== 'undefined') {
        CacheManager.setStockPrice(toshinCode, price);
      }

      return price;

    } catch (e) {
      LoggerManager.error('Error in updateToshinPrices', {
        code: toshinCode,
        error: e.message,
        stack: e.stack
      });
      
      // エラーメッセージを返す代わりに、エラーを明示的に表示
      return '取得エラー: ' + e.message;
    }
  },

  validateCode: function(code, type) {
    if (type === 'JP' && !/^\d{4}$/.test(code)) {
      throw new Error('Invalid JP stock code format');
    }
    if (type === 'TOSHIN' && !/^JP\d{10}$/.test(code)) {
      throw new Error('Invalid TOSHIN code format');
    }
  },

  handleApiError: function(response, code) {
    if (response.getResponseCode() !== 200) {
      const errorMessage = `API request failed with status ${response.getResponseCode()}`;
      LoggerManager.error(errorMessage, {
        code: code,
        status: response.getResponseCode()
      });
      throw new Error(errorMessage);
    }
  },

  retry: async function(func, retries = 3, delay = 1000) {
    for (let i = 0; i < retries; i++) {
      try {
        return await func();
      } catch (error) {
        if (i === retries - 1) throw error;
        LoggerManager.warn(`Retry attempt ${i + 1} failed`, { error: error.message });
        Utilities.sleep(delay * (i + 1));
      }
    }
  },

  // テスト用の関数を追加
  testStockPrice: function() {
    const ui = SpreadsheetApp.getUi();
    
    // テストする銘柄コードの入力を求める
    const codeResponse = ui.prompt(
      '株価取得テスト',
      '銘柄コードを入力してください（例: 7203）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (codeResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const stockCode = codeResponse.getResponseText().trim();
    
    // 取引コードの選択
    const typeResponse = ui.prompt(
      '株価取得テスト',
      '取引コードを選択してください（JP または TOSHIN）:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (typeResponse.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const tradeType = typeResponse.getResponseText().trim().toUpperCase();
    
    if (tradeType !== 'JP' && tradeType !== 'TOSHIN') {
      ui.alert('エラー', '取引コードは JP または TOSHIN を入力してください。', ui.ButtonSet.OK);
      return;
    }
    
    try {
      // 処理開始を通知
      ui.alert('処理開始', `${tradeType} コード ${stockCode} の株価を取得します...`, ui.ButtonSet.OK);
      
      // 株価取得の実行
      const startTime = new Date().getTime();
      const price = this.STOCKPRICEJP(tradeType, stockCode);
      const endTime = new Date().getTime();
      const executionTime = (endTime - startTime) / 1000;
      
      // 結果表示
      const resultMessage = `
銘柄コード: ${stockCode}
取引タイプ: ${tradeType}
取得結果: ${price}
実行時間: ${executionTime.toFixed(2)}秒

詳細ログはログシートで確認できます。
      `;
      
      ui.alert('株価取得結果', resultMessage, ui.ButtonSet.OK);
      
      // ログに記録
      LoggerManager.info('Stock price test completed', {
        code: stockCode,
        type: tradeType,
        result: price,
        executionTime: executionTime
      });
      
      // テスト結果をシートに記録
      this.recordTestResult(stockCode, tradeType, price, executionTime);
      
    } catch (e) {
      ui.alert('エラー', `株価取得中にエラーが発生しました: ${e.message}`, ui.ButtonSet.OK);
      LoggerManager.error('Stock price test error', e);
    }
  },
  
  // テスト結果をシートに記録する関数
  recordTestResult: function(code, type, result, executionTime) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName('StockPriceTests');
      
      // シートが存在しない場合は作成
      if (!sheet) {
        sheet = ss.insertSheet('StockPriceTests');
        sheet.getRange('A1:F1').setValues([[
          'タイムスタンプ',
          '銘柄コード',
          '取引タイプ',
          '取得結果',
          '実行時間(秒)',
          'ユーザー'
        ]]);
        sheet.setFrozenRows(1);
      }
      
      // 結果を記録
      sheet.appendRow([
        new Date(),
        code,
        type,
        result,
        executionTime,
        Session.getActiveUser().getEmail()
      ]);
      
    } catch (e) {
      LoggerManager.error('Error recording test result', e);
    }
  },
  
  // 複数の銘柄をバッチテストする関数
  batchTestStockPrices: function() {
    const ui = SpreadsheetApp.getUi();
    try {
      // テスト対象の銘柄コードリスト
      const testCodes = ['7203', '9984', '6758', '6861', '4755'];
      const results = [];
      
      // 各銘柄コードでテスト
      for (let i = 0; i < testCodes.length; i++) {
        const code = testCodes[i];
        const price = this.STOCKPRICEJP(code.substring(0, 2), code);
        results.push({
          code: code,
          price: price,
          success: price !== null && !isNaN(price)
        });
      }
      
      // 結果を表示
      let message = 'バッチテスト結果:\n\n';
      let successCount = 0;
      
      for (let i = 0; i < results.length; i++) {
        const result = results[i];
        message += `銘柄コード ${result.code}: ${result.success ? '成功' : '失敗'}`;
        if (result.success) {
          message += ` (${result.price}円)`;
          successCount++;
        }
        message += '\n';
      }
      
      message += `\n成功率: ${successCount}/${results.length} (${Math.round(successCount / results.length * 100)}%)`;
      
      ui.alert('バッチテスト完了', message, ui.ButtonSet.OK);
      return results;
    } catch (e) {
      ui.alert('エラー', `バッチテスト中にエラーが発生しました: ${e.message}`, ui.ButtonSet.OK);
      LoggerManager.error('Batch test error', e);
    }
  },

  getStockPriceFromAlternativeSources: function(code) {
    // Implementation of getStockPriceFromAlternativeSources function
  },
  
  /**
   * 複数の銘柄コードの株価を一括で取得するグローバル関数
   * @param {string[]} stockCodes - 取得する銘柄コードの配列
   * @return {Object} - 銘柄コードをキー、株価を値とするオブジェクト
   */
  updateMultipleStockPrices: function(stockCodes) {
    // 引数チェックを追加
    if (!stockCodes) {
      LoggerManager.error('グローバルupdateMultipleStockPrices: 引数がnullまたはundefined', {
        stockCodes: stockCodes
      });
      return {}; // 空のオブジェクトを返す
    }
    
    if (!Array.isArray(stockCodes)) {
      LoggerManager.error('グローバルupdateMultipleStockPrices: 引数が配列ではない', {
        stockCodes: stockCodes,
        type: typeof stockCodes
      });
      // 配列に変換を試みる
      try {
        stockCodes = [].concat(stockCodes);
      } catch (e) {
        return {}; // 変換できない場合は空のオブジェクトを返す
      }
    }
    
    const result = {};
    
    const batchSize = 5; // 一度に処理する銘柄数
    const waitTime = 500; // バッチ間の待機時間（ミリ秒）
    
    LoggerManager.info('複数銘柄の株価取得を開始', {
      totalCodes: stockCodes.length,
      batchSize: batchSize
    });
    
    // バッチ処理
    for (let i = 0; i < stockCodes.length; i += batchSize) {
      const currentBatch = stockCodes.slice(i, i + batchSize);
      
      // 進捗ログ
      LoggerManager.debug('株価取得バッチ処理', {
        batch: Math.floor(i / batchSize) + 1,
        totalBatches: Math.ceil(stockCodes.length / batchSize),
        currentCodes: currentBatch
      });
      
      // バッチ内の各銘柄を処理
      for (let j = 0; j < currentBatch.length; j++) {
        const code = currentBatch[j];
        try {
          // 株価取得
          const price = this.updateStockPrices(code);
          
          // 数値に変換できる場合のみ結果に追加
          if (typeof price === 'number' || (typeof price === 'string' && !isNaN(Number(price)) && !price.startsWith('取得エラー'))) {
            result[code] = typeof price === 'number' ? price : Number(price);
            LoggerManager.debug('株価取得成功', {
              code: code,
              price: result[code]
            });
          } else {
            // エラーの場合はnullを設定
            result[code] = null;
            LoggerManager.warn('株価取得失敗', {
              code: code,
              error: price
            });
          }
        } catch (e) {
          // エラーの場合はnullを設定
          result[code] = null;
          LoggerManager.error('株価取得中にエラー発生', {
            code: code,
            error: e.message
          });
        }
      }
      
      // 最後のバッチ以外は待機
      if (i + batchSize < stockCodes.length) {
        Utilities.sleep(waitTime);
      }
    }
    
    LoggerManager.info('複数銘柄の株価取得が完了', {
      totalCodes: stockCodes.length,
      successCount: Object.values(result).filter(price => price !== null).length
    });
    
    return result;
  }
};

// グローバル関数
function STOCKPRICEJP(torihiki_code, shoken_code) {
  return StockPriceUtil.STOCKPRICEJP(torihiki_code, shoken_code);
}

// テスト用のグローバル関数
function testStockPrice() {
  return StockPriceUtil.testStockPrice();
}

function batchTestStockPrices() {
  return StockPriceUtil.batchTestStockPrices();
}

/**
 * 代替ソースから株価を取得するUI関数
 * メニューから呼び出される
 */
function getStockPriceFromAlternative() {
  const ui = SpreadsheetApp.getUi();
  
  // 銘柄コードの入力を求める
  const codeResponse = ui.prompt(
    '代替ソースから株価取得',
    '銘柄コードを入力してください:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (codeResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const code = codeResponse.getResponseText().trim();
  
  if (!code || !/^\d{4}$/.test(code)) {
    ui.alert('エラー', '有効な銘柄コード（4桁の数字）を入力してください。', ui.ButtonSet.OK);
    return;
  }
  
  try {
    // 代替ソースから株価を取得
    const price = StockPriceUtil.getStockPriceFromAlternativeSources(code);
    
    if (price && !isNaN(price)) {
      // 成功した場合、スプレッドシートに反映するか確認
      const confirmResponse = ui.alert(
        '株価取得成功',
        `銘柄コード ${code} の株価: ${price}円\n\nこの価格をスプレッドシートに反映しますか？`,
        ui.ButtonSet.YES_NO
      );
      
      if (confirmResponse === ui.Button.YES) {
        // スプレッドシートに反映
        const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MAIN);
        if (!mainSheet) {
          throw new Error('メインシートが見つかりません');
        }
        
        // 銘柄コードを検索
        const dataRange = mainSheet.getDataRange();
        const values = dataRange.getValues();
        
        // 列インデックスを取得
        const headerRow = values[0];
        const codeColIndex = headerRow.findIndex(header => header === '銘柄コード' || header === 'コード');
        const priceColIndex = headerRow.findIndex(header => header === '株価' || header === '現在値');
        
        if (codeColIndex === -1 || priceColIndex === -1) {
          throw new Error('銘柄コードまたは株価の列が見つかりません');
        }
        
        // 銘柄コードを検索して株価を更新
        let updated = false;
        for (let i = 1; i < values.length; i++) {
          if (values[i][codeColIndex] == code) { // 文字列と数値の比較のため == を使用
            mainSheet.getRange(i + 1, priceColIndex + 1).setValue(price);
            updated = true;
          }
        }
        
        if (updated) {
          ui.alert('更新完了', `銘柄コード ${code} の株価を ${price}円 に更新しました。`, ui.ButtonSet.OK);
        } else {
          ui.alert('更新失敗', `銘柄コード ${code} がスプレッドシートに見つかりませんでした。`, ui.ButtonSet.OK);
        }
      }
    } else {
      ui.alert('取得失敗', `銘柄コード ${code} の株価を代替ソースから取得できませんでした。`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('エラー', `株価取得中にエラーが発生しました: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * 複数の銘柄コードの株価を一括で取得するグローバル関数
 * @param {string[]} stockCodes - 取得する銘柄コードの配列
 * @return {Object} - 銘柄コードをキー、株価を値とするオブジェクト
 */
function updateMultipleStockPrices(stockCodes) {
  return StockPriceUtil.updateMultipleStockPrices(stockCodes);
}

