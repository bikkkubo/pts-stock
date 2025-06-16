// ==== fetchers.gs  (night PTS ranking) ====
/**
 * @param {'increase'|'decrease'} type
 * @return {Array.<{code:string,name:string,pts:number,prev:number,rate:number}>}
 */
function fetchNightPts(type) {
  var url = 'https://kabutan.jp/warning/pts_night_price_' + type;
  var opt = { headers: { 'User-Agent': 'Mozilla/5.0' }, muteHttpExceptions: true };
  var html, resp, retry = 0;
  do {
    resp = UrlFetchApp.fetch(url, opt);
    if (resp.getResponseCode() === 200) { html = resp.getContentText('utf-8'); break; }
    Utilities.sleep(1500);
  } while (++retry < 3);
  if (!html) throw new Error('Failed to fetch PTS ranking: ' + url);

  var rowRe = /<tr[^>]*?>[\s\S]*?<\/tr>/g,
      cellRe = /<td[^>]*?>([\s\S]*?)<\/td>/g,
      tagRe = /<[^>]+>/g,
      rows = [];

  (html.match(rowRe) || []).forEach(function (r) {
    var cells = (r.match(cellRe) || []).map(function (c) { return c.replace(tagRe, '').trim(); });
    if (cells.length < 7 || !/^\d{3,4}$/.test(cells[0])) return; // skip header/etc
    rows.push({
      code: cells[0],
      name: cells[1],
      pts:  Number(cells[4].replace(/,/g, '')),
      prev: Number(cells[5].replace(/,/g, '')),
      rate: Number(cells[6].replace(/[+%]/g, ''))   // already %
    });
  });
  return rows;
}

/** entry point for main() */
function fetchPts() {
  return fetchNightPts('increase').concat(fetchNightPts('decrease'));
}

/**
 * Fetch news articles for a specific stock symbol from multiple sources
 * @param {string} code - Stock symbol code
 * @return {Array} Array of news article objects
 */
function fetchNews(code) {
  try {
    var allNews = [];
    
    // 1. Fetch news from Kabutan stock page
    var kabutanNews = fetchKabutanStockNews(code);
    if (kabutanNews && kabutanNews.length > 0) {
      allNews = allNews.concat(kabutanNews);
      Logger.log('Fetched ' + kabutanNews.length + ' articles for ' + code + ' from Kabutan');
    }
    
    // 2. Enhanced TDnet IR information
    var tdnetNews = fetchEnhancedTdnetNews(code);
    if (tdnetNews && tdnetNews.length > 0) {
      allNews = allNews.concat(tdnetNews);
      Logger.log('Fetched ' + tdnetNews.length + ' IR articles for ' + code + ' from Enhanced TDnet');
    }
    
    // 3. Company IR URLs (TDnet RSS fallback to be added later)
    var irUrls = []; // removed generic IR URL generation to prevent DNS errors
    
    // 4. Enhanced Nikkei search with company name
    var nikkeiNews = fetchEnhancedNikkeiNews(code);
    if (nikkeiNews && nikkeiNews.length > 0) {
      allNews = allNews.concat(nikkeiNews);
      Logger.log('Fetched ' + nikkeiNews.length + ' articles for ' + code + ' from Nikkei');
    }
    
    // 5. Yahoo Finance Japan news
    var yahooNews = fetchYahooFinanceNews(code);
    if (yahooNews && yahooNews.length > 0) {
      allNews = allNews.concat(yahooNews);
      Logger.log('Fetched ' + yahooNews.length + ' articles for ' + code + ' from Yahoo Finance');
    }
    
    // 6. Bloomberg Japan (RSS)
    var bloombergNews = fetchBloombergJapanNews(code);
    if (bloombergNews && bloombergNews.length > 0) {
      allNews = allNews.concat(bloombergNews);
      Logger.log('Fetched ' + bloombergNews.length + ' articles for ' + code + ' from Bloomberg Japan');
    }
    
    // 7. Toyo Keizai Online
    var toyokeizaiNews = fetchToyoKeizaiNews(code);
    if (toyokeizaiNews && toyokeizaiNews.length > 0) {
      allNews = allNews.concat(toyokeizaiNews);
      Logger.log('Fetched ' + toyokeizaiNews.length + ' articles for ' + code + ' from Toyo Keizai');
    }
    
    // Remove duplicates and return
    if (allNews.length > 0) {
      var uniqueNews = removeDuplicateNews(allNews);
      Logger.log('Total unique articles for ' + code + ': ' + uniqueNews.length);
      return uniqueNews;
    }
    
    Logger.log('No news found for stock code: ' + code);
    return [];
    
  } catch (error) {
    Logger.log('Error fetching news for ' + code + ': ' + error.toString());
    return [];
  }
}

/**
 * Get company overview from Kabutan
 * @param {string} code - Stock symbol code
 * @return {string} Company overview text
 */
function getCompanyOverview(code) {
  try {
    var url = 'https://kabutan.jp/stock/?code=' + code;
    
    var response = UrlFetchApp.fetch(url, {
      'method': 'GET',
      'headers': {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('Failed to fetch company overview for ' + code + ': HTTP ' + response.getResponseCode());
      return '';
    }
    
    var html = response.getContentText();
    
    // Look for company overview section
    var overviewPattern = /<th[^>]*>概要<\/th>[\s\S]*?<td[^>]*>(.*?)<\/td>/i;
    var match = overviewPattern.exec(html);
    
    if (match) {
      var overview = match[1]
        .replace(/<[^>]*>/g, '')  // Remove HTML tags
        .replace(/\s+/g, ' ')     // Normalize whitespace
        .trim();
        
      Logger.log('Retrieved overview for ' + code + ': ' + overview.substring(0, 100) + '...');
      return overview;
    }
    
    Logger.log('No overview found for ' + code);
    return '';
    
  } catch (error) {
    Logger.log('Error fetching company overview for ' + code + ': ' + error.toString());
    return '';
  }
}

// Legacy functions for fallback compatibility
function fetchKabutanStockNews(code) { return []; }
function fetchEnhancedTdnetNews(code) { return []; }
function fetchEnhancedNikkeiNews(code) { return []; }
function fetchYahooFinanceNews(code) { return []; }
function fetchBloombergJapanNews(code) { return []; }
function fetchToyoKeizaiNews(code) { return []; }
function removeDuplicateNews(articles) { return articles; }