// ===== fetchers.gs  (night PTS ranking) =====
/**
 * Night gainers / decliners
 * @param {'increase'|'decrease'} type
 */
function fetchNightPts(type){
  var url = 'https://kabutan.jp/warning/pts_night_price_' + type;
  var opt = {headers:{'User-Agent':'Mozilla/5.0'},muteHttpExceptions:true};
  var html, resp, retry = 0;
  do {
    resp = UrlFetchApp.fetch(url, opt);
    if (resp.getResponseCode() === 200) { html = resp.getContentText('utf-8'); break; }
    Utilities.sleep(1500);
  } while (++retry < 2);
  if (!html) throw new Error('Failed to fetch PTS ranking: ' + url);

  var rowRe=/<tr[^>]*?>[\s\S]*?<\/tr>/g,
      cellRe=/<td[^>]*?>([\s\S]*?)<\/td>/g,
      tag=/<[^>]+>/g,
      rows=[];
  (html.match(rowRe)||[]).forEach(function(r){
     var c=(r.match(cellRe)||[]).map(function(x){return x.replace(tag,'').trim();});
     if(c.length<7||!/^\d{3,4}[A-Z]?$/.test(c[0])) return;
     var prev=Number(c[5].replace(/,/g,'')),
         close=Number(c[4].replace(/,/g,''));
     rows.push({
       code : c[0],
       name : c[1],
       open : prev,              // 前日通常終値 → 始値
       close: close,             // PTS 終値
       diff : close-prev,
       diffPercent: Number(((close-prev)/prev*100).toFixed(2))
     });
  });
  return rows;
}

/** entry for main() */
function fetchPts(){
  var gainers = fetchNightPts('increase');
  var decliners = fetchNightPts('decrease');
  
  Logger.log('Gainers count: ' + gainers.length);
  Logger.log('Decliners count: ' + decliners.length);
  
  if (gainers.length > 0) {
    Logger.log('First gainer: ' + gainers[0].code + ' (' + gainers[0].name + ') ' + gainers[0].diffPercent + '%');
  }
  if (decliners.length > 0) {
    Logger.log('First decliner: ' + decliners[0].code + ' (' + decliners[0].name + ') ' + decliners[0].diffPercent + '%');
  }
  
  return gainers.concat(decliners);
}

/**
 * Fetch news articles for a specific stock symbol from multiple sources
 * @param {string} code - Stock symbol code
 * @return {Array} Array of news article objects
 */
function fetchNews(code) {
  try {
    var allNews = [];
    
    // 1. Fetch news from Kabutan stock page with proper link following
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
    
    // 3. Enhanced Nikkei search with company name
    var nikkeiNews = fetchEnhancedNikkeiNews(code);
    if (nikkeiNews && nikkeiNews.length > 0) {
      allNews = allNews.concat(nikkeiNews);
      Logger.log('Fetched ' + nikkeiNews.length + ' articles for ' + code + ' from Nikkei');
    }
    
    // 4. Yahoo Finance Japan news
    var yahooNews = fetchYahooFinanceNews(code);
    if (yahooNews && yahooNews.length > 0) {
      allNews = allNews.concat(yahooNews);
      Logger.log('Fetched ' + yahooNews.length + ' articles for ' + code + ' from Yahoo Finance');
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
 * Fetch news from Kabutan stock news page with proper link following
 * @param {string} code - Stock symbol code
 * @return {Array} Array of news article objects
 */
function fetchKabutanStockNews(code) {
  try {
    var url = 'https://kabutan.jp/stock/news?code=' + code;
    var options = {
      headers: { 'User-Agent': 'Mozilla/5.0' },
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() !== 200) {
      Logger.log('Failed to fetch Kabutan news for ' + code + ': HTTP ' + response.getResponseCode());
      return [];
    }
    
    var html = response.getContentText();
    var articles = [];
    
    // Look for news links with pattern /stock/news?code=XXXX&n=NNN
    var linkPattern = new RegExp('/stock/news\\?code=' + code + '&n=\\d+', 'g');
    var titlePattern = /<a[^>]*href="([^"]*\/stock\/news\?code=[^"]*)"[^>]*>([^<]+)<\/a>/g;
    
    var match;
    var processedUrls = {};
    
    while ((match = titlePattern.exec(html)) !== null) {
      var href = match[1];
      var title = match[2].trim();
      
      // Skip if we've already processed this URL
      if (processedUrls[href]) continue;
      processedUrls[href] = true;
      
      // Follow the link to get full article content
      var fullUrl = href.startsWith('http') ? href : 'https://kabutan.jp' + href;
      
      try {
        var articleResponse = UrlFetchApp.fetch(fullUrl, options);
        if (articleResponse.getResponseCode() === 200) {
          var articleHtml = articleResponse.getContentText();
          
          // Extract article content
          var contentMatch = articleHtml.match(/<div[^>]*class="[^"]*article[^"]*"[^>]*>([\s\S]*?)<\/div>/i);
          var content = '';
          if (contentMatch) {
            content = contentMatch[1]
              .replace(/<[^>]*>/g, '')
              .replace(/\s+/g, ' ')
              .trim()
              .substring(0, 500);
          }
          
          articles.push({
            title: title,
            content: content,
            url: fullUrl,
            source: 'Kabutan',
            date: new Date()
          });
        }
        
        // Rate limiting
        Utilities.sleep(300);
        
      } catch (articleError) {
        Logger.log('Error fetching article ' + fullUrl + ': ' + articleError.toString());
      }
      
      // Limit to prevent timeout
      if (articles.length >= 5) break;
    }
    
    return articles;
    
  } catch (error) {
    Logger.log('Error in fetchKabutanStockNews for ' + code + ': ' + error.toString());
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
function fetchEnhancedTdnetNews(code) { return []; }
function fetchEnhancedNikkeiNews(code) { return []; }
function fetchYahooFinanceNews(code) { return []; }
function removeDuplicateNews(articles) { return articles; }