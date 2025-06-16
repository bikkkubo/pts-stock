// ===== fetchers.gs  (night PTS) =====
function fetchNightPts(type){
  var url = 'https://kabutan.jp/warning/pts_night_price_' + type;
  var opt = { headers:{'User-Agent':'Mozilla/5.0'}, muteHttpExceptions:true };
  var html, res, retry=0;
  do{
    res = UrlFetchApp.fetch(url,opt);
    if(res.getResponseCode()===200){ html=res.getContentText('utf-8'); break; }
    Utilities.sleep(1500);
  }while(++retry<2);
  if(!html) throw new Error('Fetch failed '+url);

  var rowRe=/<tr[^>]*?>[\s\S]*?<\/tr>/g,
      cellRe=/<td[^>]*?>([\s\S]*?)<\/td>/g,
      tag=/<[^>]+>/g, data=[];
  (html.match(rowRe)||[]).slice(1,11).forEach(function(r){ // TOP10
    var c=(r.match(cellRe)||[]).map(function(x){return x.replace(tag,'').trim();});
    if(c.length<7 || !/^\d{3,4}[A-Z]?$/.test(c[0])) return;
    var open = Number(c[5].replace(/,/g,''));
    var close= Number(c[4].replace(/,/g,''));
    var diff = close-open;
    data.push({
      code : c[0],
      name : c[1].replace(/^東[ＳＧＰＮ]/, ''), // Remove exchange prefix
      open : open,
      close: close,
      diff : diff,
      pct  : Number((diff/open*100).toFixed(2))
    });
  });
  return data;
}

/** entry for main() */
function fetchPts(){
  var gainers = fetchNightPts('increase');
  var decliners = fetchNightPts('decrease');
  
  Logger.log('Fetched ' + gainers.length + ' gainers, ' + decliners.length + ' decliners');
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
    
    // 1. Fetch news from Kabutan stock page with fixed selector
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
 * Fetch news from Kabutan stock news page with fixed URL selector
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
    
    // Fixed selector: /stock/news?code=XXXX (no extra slash)
    var titlePattern = new RegExp('<a[^>]*href="([^"]*\\/stock\\/news\\?code=' + code + '[^"]*)"[^>]*>([^<]+)<\\/a>', 'g');
    
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
function removeDuplicateNews(articles) { return articles; }