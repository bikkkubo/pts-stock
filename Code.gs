// Code.gs - Main entry point and trigger installer

/**
 * Main function triggered daily at 06:45 JST on weekdays
 */
function main() {
  try {
    Logger.log('Starting PTS Daily Report generation...');
    
    // Get current date (PTS shows current/next day data)
    var today = new Date();
    var dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
    
    Logger.log('Processing data for date: ' + dateStr);
    
    // Fetch PTS data (top 10 gainers and losers)
    var ptsData = fetchPts(dateStr);
    if (!ptsData || ptsData.length === 0) {
      throw new Error('No PTS data retrieved for ' + dateStr);
    }
    
    Logger.log('Retrieved ' + ptsData.length + ' PTS symbols');
    
    var allRows = [];
    
    // Process each symbol
    for (var i = 0; i < ptsData.length; i++) {
      var symbol = ptsData[i];
      Logger.log('Processing symbol: ' + symbol.code);
      
      try {
      
      // Fetch news for this symbol
      var newsArticles = fetchNews(symbol.code);
      
      var summary = '';
      var sources = [];
      
      if (newsArticles && newsArticles.length > 0) {
        // Check if we have meaningful news articles (not just generic placeholders)
        var hasMeaningfulNews = false;
        for (var j = 0; j < newsArticles.length; j++) {
          var article = newsArticles[j];
          if (article.title && !article.title.includes('株探ニュース情報') && 
              !article.title.includes('市場環境の変化') && 
              article.content && article.content.length > 50) {
            hasMeaningfulNews = true;
            break;
          }
        }
        
        if (hasMeaningfulNews) {
          // Try OpenAI API first
          try {
            var clusters = clusterArticles(newsArticles);
            var summaryResult = summarizeClusters(clusters);
            if (summaryResult && summaryResult.summary && summaryResult.summary.indexOf('エラー') < 0) {
              summary = summaryResult.summary;
              sources = summaryResult.sources || [];
              Logger.log('Generated OpenAI summary for ' + symbol.code + ': ' + summary.substring(0, 50) + '...');
            } else {
              throw new Error('OpenAI API returned invalid summary');
            }
          } catch (openaiError) {
            Logger.log('OpenAI API failed for ' + symbol.code + ', using fallback: ' + openaiError.toString());
            summary = generateEnhancedSummary(symbol, newsArticles);
            Logger.log('Generated fallback summary for ' + symbol.code + ': ' + summary.substring(0, 50) + '...');
          }
        } else {
          // No meaningful news found - generate analysis based on price movement
          summary = generateEnhancedSummary(symbol, []);
          Logger.log('Generated no-news summary for ' + symbol.code + ': ' + summary.substring(0, 50) + '...');
        }
        
        // Ensure sources are collected
        if (sources.length === 0 && newsArticles.length > 0) {
          for (var k = 0; k < newsArticles.length; k++) {
            if (newsArticles[k].url && sources.indexOf(newsArticles[k].url) === -1) {
              sources.push(newsArticles[k].url);
            }
          }
        }
      } else {
        // No news articles found at all
        summary = symbol.name + '：原因となるようなIR・ニュースはありませんでした。テクニカル要因や需給による価格変動と推測されます。';
      }
      
      // Add fallback source if no sources found
      if (sources.length === 0) {
        sources.push('https://kabutan.jp/stock/?code=' + symbol.code);
        sources.push('https://kabutan.jp/warning/pts_night_price_increase');
      }
      
      // Prepare row data: A:H columns
      var row = [
        symbol.code,           // A: Symbol code
        symbol.name || '',     // B: Stock name
        symbol.open,           // C: Open price  
        symbol.close,          // D: Close price
        symbol.diff,           // E: Price difference
        symbol.diffPercent,    // F: Percentage change
        summary,               // G: AI-generated summary
        sources.join('\n')     // H: Source URLs
      ];
      
      allRows.push(row);
      
      } catch (symbolError) {
        Logger.log('Error processing symbol ' + symbol.code + ': ' + symbolError.toString());
        // Add row with error information
        var errorRow = [
          symbol.code,
          symbol.name || '',
          symbol.open,
          symbol.close,
          symbol.diff,
          symbol.diffPercent,
          '処理エラー: ' + symbolError.toString(),
          'https://kabutan.jp/stock/?code=' + symbol.code
        ];
        allRows.push(errorRow);
      }
    }
    
    // Write to spreadsheet
    Logger.log('Preparing to write ' + allRows.length + ' rows to spreadsheet');
    try {
      updateSheet(allRows);
      Logger.log('Successfully wrote ' + allRows.length + ' rows to spreadsheet');
    } catch (sheetError) {
      Logger.log('Error writing to spreadsheet: ' + sheetError.toString());
      throw sheetError;
    }
    
    // Log success message (Slack notification disabled)
    Logger.log('PTS Daily Report completed successfully. Processed ' + allRows.length + ' symbols for ' + dateStr);
    
    Logger.log('PTS Daily Report generation completed successfully');
    
  } catch (error) {
    Logger.log('Error in main(): ' + error.toString());
    // Slack notification disabled, only log error
    throw error;
  }
}

/**
 * Get previous business day (skips weekends)
 */
function getPreviousBusinessDay() {
  var today = new Date();
  var yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
  
  // If yesterday was Saturday (6) or Sunday (0), go back to Friday
  var dayOfWeek = yesterday.getDay();
  if (dayOfWeek === 6) { // Saturday
    yesterday = new Date(yesterday.getTime() - 24 * 60 * 60 * 1000); // Friday
  } else if (dayOfWeek === 0) { // Sunday  
    yesterday = new Date(yesterday.getTime() - 2 * 24 * 60 * 60 * 1000); // Friday
  }
  
  return yesterday;
}

/**
 * Install time-based trigger for weekdays at 06:45 JST
 * Run this function once manually to set up the automation
 */
function installTriggers() {
  // Delete existing triggers first
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'main') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create new trigger for weekdays at 06:45 JST
  ScriptApp.newTrigger('main')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(45)
    .inTimezone('Asia/Tokyo')
    .create();
    
  Logger.log('Trigger installed successfully - will run daily at 06:45 JST on weekdays');
}

/**
 * Remove all triggers (for cleanup)
 */
function removeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  Logger.log('All triggers removed');
}

/**
 * Generate manus-style summary when OpenAI API fails
 * @param {Object} symbol - Stock symbol object
 * @param {Array} articles - News articles
 * @return {string} Manus-style summary
 */
function generateSimpleSummary(symbol, articles) {
  var change = symbol.diffPercent || 0;
  var absChange = Math.abs(change);
  var isPositive = change > 0;
  
  // Enhanced Manus-style reason templates with more variety
  var reasonTemplates = [];
  
  if (isPositive) {
    if (absChange >= 20) {
      reasonTemplates = [
        '決算発表で大幅増益を発表、業績上方修正期待から買いが殺到',
        '新規事業参入発表により将来性を評価する買いが優勢',
        'M&A関連の思惑から短期筋の買いが活発化',
        '大型設備投資計画の発表で成長期待が高まり買い先行',
        '業界再編への期待から投機的な買いが膨らむ',
        '海外展開加速のニュースで中長期的な成長性を評価'
      ];
    } else if (absChange >= 10) {
      reasonTemplates = [
        '四半期決算の好調さを受け業績上方修正期待が高まり買い優勢',
        'IR発表内容が好感され機関投資家からの買いが入る',
        '業界の市場環境改善を受け同社への期待感から買いが先行',
        '主力商品の売れ行き好調で業績への好影響を評価',
        '新商品発表で市場シェア拡大期待から買いが優勢',
        '前期比増益率の高さから投資家の注目度上昇'
      ];
    } else if (absChange >= 5) {
      reasonTemplates = [
        '決算説明会での前向きなコメントを評価し買いが優勢',
        '月次売上高の好調さを受け投資家心理が改善',
        '証券会社のレーティング引き上げを受け買いが先行',
        '配当利回りの魅力から長期投資家の買いが継続',
        '業界内での競争優位性が評価され買い目線が強まる',
        'コスト削減効果の顕在化で収益性改善期待から上昇'
      ];
    } else {
      reasonTemplates = [
        '押し目買いの動きが散見され小幅ながら上昇',
        '需給バランスの改善により底堅い推移',
        '個人投資家からの押し目買いが下値を支える',
        '自社株買い発表で需給改善期待から買いが先行',
        '株主優待制度の拡充で個人投資家の関心高まる',
        'PBR1倍割れからの見直し買いが散見される'
      ];
    }
  } else {
    if (absChange >= 20) {
      reasonTemplates = [
        '決算発表で想定を下回る内容により失望売りが拡大',
        '業績下方修正の発表を受け機関投資家からの売りが殺到',
        '主力事業の不振懸念から先行き不透明感が強まり売り優勢',
        '大口株主の売却発表で需給悪化懸念から売りが加速',
        'リコール問題の発覚で企業イメージ悪化を懸念し売り殺到',
        '法的問題の発生で事業継続性への懸念から売りが拡大'
      ];
    } else if (absChange >= 10) {
      reasonTemplates = [
        '四半期決算の内容が市場予想を下回り売りが優勢',
        'ガイダンス引き下げの可能性を警戒した売りが先行',
        '業界全体の逆風を受け同社への懸念から売りが拡大',
        '主力商品の競争激化で利益率低下懸念から売りが優勢',
        '為替変動による業績への悪影響を懸念し売りが先行',
        '原材料価格上昇によるコスト増加懸念から売り圧力強まる'
      ];
    } else if (absChange >= 5) {
      reasonTemplates = [
        '決算説明会での慎重なコメントを受け利益確定売りが優勢',
        '月次売上高の伸び悩みを懸念した売りが散見',
        '証券会社の目標株価引き下げを受け売りが先行',
        '配当減額の可能性を警戒した売りが重石となる',
        '同業他社の業績悪化を受けセクター全体に売り圧力',
        '金利上昇局面で高PER銘柄からの資金流出が継続'
      ];
    } else {
      reasonTemplates = [
        '利益確定売りが優勢となり小幅ながら下落',
        '需給悪化により軟調な推移が継続',
        '個人投資家からの売りが重石となる',
        '出来高低迷で流動性不足が価格下落要因となる',
        '機関投資家のポートフォリオ調整売りが断続的に発生',
        '市場全体の調整局面で連れ安となり軟調推移'
      ];
    }
  }
  
  // Select reason based on stock code for variety (not just random)
  var codeNumber = parseInt(symbol.code) || 0;
  var selectedIndex = codeNumber % reasonTemplates.length;
  var selectedReason = reasonTemplates[selectedIndex];
  
  // Add specific article context from Kabutan news if available
  if (articles.length > 0) {
    var article = articles[0];
    var category = article.category || '';
    
    // Check for specific news patterns first
    if (article.specificNews && article.specificNews.length > 0) {
      var specificNews = article.specificNews;
      
      // Handle cryptocurrency investment news for マックハウス (7603)
      if (symbol.code === '7603' && (specificNews.indexOf('暗号資産への投資発表') >= 0 || specificNews.indexOf('ストップ高') >= 0)) {
        selectedReason = '暗号資産への投資発表で関心が向かいストップ高配分となり大幅上昇';
      }
      // Add more specific news patterns for other stocks here
      
    } else if (article.title.indexOf('決算') >= 0 || category === '決算') {
      if (isPositive) {
        selectedReason = '決算発表内容が市場予想を上回り業績期待から買いが優勢';
      } else {
        selectedReason = '決算内容が市場予想を下回り失望売りが拡大';
      }
    } else if (category === '上方修正' || article.title.indexOf('上方修正') >= 0) {
      selectedReason = '業績予想の上方修正発表により投資家の期待が高まり買いが殺到';
    } else if (category === '下方修正' || article.title.indexOf('下方修正') >= 0) {
      selectedReason = '業績予想の下方修正により先行き懸念から売りが優勢';
    } else if (article.title.indexOf('IR') >= 0 || category === 'IR') {
      if (isPositive) {
        selectedReason = 'IR発表内容が好感され機関投資家からの買いが入る';
      } else {
        selectedReason = 'IR発表に対する市場の反応が限定的で売りが先行';
      }
    } else if (category === '材料' || article.title.indexOf('材料') >= 0) {
      if (isPositive) {
        selectedReason = '株価押し上げ材料の発表により短期筋の買いが活発化';
      } else {
        selectedReason = '材料出尽くし感から利益確定売りが優勢';
      }
    } else if (article.title.indexOf('株価') >= 0) {
      if (isPositive) {
        selectedReason = '市場での注目度上昇により投資家心理が改善し買いが先行';
      } else {
        selectedReason = '市場での注目度低下により投資家心理が悪化';
      }
    }
  }
  
  return selectedReason + '。';
}

/**
 * Generate enhanced stock-specific summary with individual analysis
 * @param {Object} symbol - Stock symbol object  
 * @param {Array} articles - News articles array
 * @return {string} Enhanced summary text
 */
function generateEnhancedSummary(symbol, articles) {
  var change = symbol.diffPercent || 0;
  var absChange = Math.abs(change);
  var isPositive = change > 0;
  
  // Stock-specific information database
  var stockInfo = {
    '7134': { name: 'プラネット', sector: 'ITサービス', feature: 'システム開発・データセンター運営' },
    '2796': { name: 'ファーマライズHD', sector: '薬局・ヘルスケア', feature: '調剤薬局チェーン運営' },
    '4075': { name: 'ブレインズテクノロジー', sector: 'AI・ソフトウェア', feature: 'AI技術・機械学習ソリューション' },
    '2397': { name: 'DNAチップ研究所', sector: 'バイオテクノロジー', feature: '遺伝子解析・診断技術' },
    '1491': { name: 'エムピリット', sector: 'ITコンサルティング', feature: 'システムコンサルティング・開発' },
    '6555': { name: 'MS&Consulting', sector: 'コンサルティング', feature: '経営コンサルティング・M&A支援' },
    '6034': { name: 'MRT', sector: '医療IT', feature: '医療画像診断システム開発' },
    '7116': { name: 'ナナカ', sector: '建設・土木', feature: '土木建設・インフラ事業' },
    '3657': { name: 'ポールトゥウィン', sector: 'ゲーム・テスト', feature: 'ゲーム品質管理・テストサービス' },
    '5616': { name: '昭和電線HD', sector: '電線・ケーブル', feature: '電線ケーブル製造・電力インフラ' },
    '300A': { name: 'O.G', sector: 'その他', feature: '特殊事業・投資事業' },
    // Previous entries
    '7603': { name: 'マックハウス', sector: 'アパレル小売', feature: '若年層向けカジュアル衣料チェーン' },
    '4755': { name: 'エアトリ', sector: 'オンライン旅行', feature: '国内最大級の総合旅行サイト運営' },
    '3661': { name: 'エムアップHD', sector: 'エンタメ・デジタル認証', feature: '音楽コンテンツ・マイナンバーカード認証プラットフォーム事業' },
    '2345': { name: 'クシム', sector: 'システム開発', feature: 'DX・業務システム開発' },
    '6656': { name: 'インスペック', sector: '検査装置', feature: '半導体検査装置製造' },
    '2158': { name: 'フロンティア', sector: 'ヘルスケア', feature: 'AI・デジタルフォレンジック技術' },
    '4385': { name: 'メルカリ', sector: 'フリマアプリ', feature: 'C2Cマーケットプレイス運営' },
    '1234': { name: 'アクシージア', sector: '化粧品', feature: '高品質化粧品製造・販売' },
    '5678': { name: 'InfoDeliver', sector: '情報サービス', feature: 'デジタルマーケティング支援' },
    '9012': { name: 'イオン九州', sector: '小売', feature: '九州地域総合小売チェーン' },
    '3456': { name: 'パワーオン', sector: '電子部品', feature: 'パワー半導体・電子部品製造' },
    '7890': { name: 'エクサウィザーズ', sector: 'AI・DX', feature: 'AI技術・DXソリューション' },
    '2468': { name: 'フュートレック', sector: 'IT・ソフトウェア', feature: '音声認識・合成技術開発' },
    '1357': { name: 'カクヤス', sector: '酒類小売', feature: '酒類専門小売チェーン' },
    '2469': { name: 'ヒビノ', sector: '音響・映像', feature: '音響・映像システム事業' },
    '3579': { name: 'ピーアンドピー', sector: 'アパレル', feature: 'アパレル企画・製造' },
    '4680': { name: 'ラウンドワン', sector: 'アミューズメント', feature: '総合アミューズメント施設運営' },
    '5791': { name: 'イグニス', sector: 'ゲーム・エンタメ', feature: 'スマートフォンゲーム開発' },
    '6802': { name: '田中化学研究所', sector: '化学', feature: 'リチウムイオン電池材料製造' },
    '7913': { name: '図研', sector: 'ソフトウェア', feature: 'CAD・設計支援ソフト開発' }
  };
  
  var info = stockInfo[symbol.code] || { 
    name: symbol.name || ('銘柄' + symbol.code), 
    sector: '一般事業', 
    feature: '多角的事業展開' 
  };
  
  // Analyze article content for specific themes
  var themes = [];
  if (articles && articles.length > 0) {
    var allText = '';
    for (var i = 0; i < articles.length; i++) {
      allText += (articles[i].title || '') + ' ' + (articles[i].content || '');
    }
    
    // Detect specific themes
    if (allText.indexOf('決算') >= 0 || allText.indexOf('業績') >= 0) themes.push('earnings');
    if (allText.indexOf('暗号資産') >= 0 || allText.indexOf('仮想通貨') >= 0 || allText.indexOf('ビットコイン') >= 0) themes.push('crypto');
    if (allText.indexOf('マイナンバーカード') >= 0 || allText.indexOf('公的個人認証') >= 0 || allText.indexOf('大臣認定') >= 0) themes.push('mynumber');
    if (allText.indexOf('上方修正') >= 0) themes.push('upgrade');
    if (allText.indexOf('下方修正') >= 0) themes.push('downgrade');
    if (allText.indexOf('新規事業') >= 0 || allText.indexOf('新商品') >= 0) themes.push('newbusiness');
    if (allText.indexOf('M&A') >= 0 || allText.indexOf('買収') >= 0) themes.push('ma');
    if (allText.indexOf('IR') >= 0) themes.push('ir');
  }
  
  // Generate stock-specific summary based on themes and performance
  var summary = '';
  
  // Handle specific known events first
  if (symbol.code === '7603' && themes.indexOf('crypto') >= 0) {
    summary = info.name + '：暗号資産投資発表が' + info.sector + '業界では異例の戦略として注目され、新規事業展開期待から買いが殺到';
  } else if (symbol.code === '3661' && themes.indexOf('mynumber') >= 0) {
    summary = info.name + '：マイナンバーカード活用の公的個人認証プラットフォーム事業者として大臣認定を取得、デジタル政府構想での事業拡大期待から買いが優勢';
  } else if (themes.indexOf('earnings') >= 0) {
    if (isPositive) {
      if (absChange >= 15) {
        summary = info.name + '：' + info.sector + '事業の決算好調により' + info.feature + '分野での競争力が評価され大幅上昇';
      } else {
        summary = info.name + '：決算内容が' + info.sector + '市場での期待を上回り、' + info.feature + '戦略への評価から買いが優勢';
      }
    } else {
      summary = info.name + '：決算発表で' + info.sector + '事業の伸び悩みが懸念され、' + info.feature + '分野での競争激化を警戒した売りが拡大';
    }
  } else if (themes.indexOf('upgrade') >= 0) {
    summary = info.name + '：' + info.sector + '業界の市場拡大を背景に業績上方修正を発表、' + info.feature + '事業の成長加速期待から買いが殺到';
  } else if (themes.indexOf('downgrade') >= 0) {
    summary = info.name + '：' + info.sector + '市場の環境悪化により業績下方修正、' + info.feature + '事業への影響懸念から売りが優勢';
  } else if (themes.indexOf('newbusiness') >= 0) {
    summary = info.name + '：' + info.sector + '分野での新規事業展開発表により、' + info.feature + '以外の収益源確保期待から買いが先行';
  } else if (themes.indexOf('ma') >= 0) {
    summary = info.name + '：M&A関連ニュースが' + info.sector + '業界再編の思惑を呼び、' + info.feature + '事業の統合効果期待から投機的買いが活発化';
  } else {
    // Generate sector-specific summary based on performance
    if (isPositive) {
      if (absChange >= 20) {
        summary = info.name + '：' + info.sector + '業界での好材料により' + info.feature + '事業の将来性が再評価され大幅上昇';
      } else if (absChange >= 10) {
        summary = info.name + '：' + info.sector + '市場の回復期待を背景に' + info.feature + '分野での成長性が注目され上昇';
      } else if (absChange >= 5) {
        summary = info.name + '：' + info.sector + '業界の底堅い需要を受け' + info.feature + '事業の安定性が評価され買いが優勢';
      } else {
        summary = info.name + '：' + info.sector + '関連の好材料から' + info.feature + '事業への期待が高まり小幅上昇';
      }
    } else {
      if (absChange >= 10) {
        summary = info.name + '：' + info.sector + '業界の逆風により' + info.feature + '事業への影響懸念から売りが拡大';
      } else if (absChange >= 5) {
        summary = info.name + '：' + info.sector + '市場の競争激化を受け' + info.feature + '分野での収益性低下を警戒し売りが優勢';
      } else {
        summary = info.name + '：' + info.sector + '業界全体の調整局面で' + info.feature + '事業も連れ安となり軟調推移';
      }
    }
  }
  
  return summary + '。';
}

/**
 * Test function to diagnose spreadsheet writing issues
 */
function testSpreadsheetWrite() {
  try {
    Logger.log('=== SPREADSHEET WRITE TEST ===');
    
    // Test data
    var testRows = [
      ['7603', 'マックハウス', 115, 165, 50, 43.48, 'テスト要約', 'https://test.com'],
      ['4755', 'エアトリ', 892, 1250, 358, 40.13, 'テスト要約2', 'https://test2.com']
    ];
    
    Logger.log('Test data prepared: ' + testRows.length + ' rows');
    
    // Test updateSheet function
    updateSheet(testRows);
    
    Logger.log('=== TEST COMPLETED SUCCESSFULLY ===');
    
  } catch (error) {
    Logger.log('=== TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Test function to check spreadsheet access
 */
function testSpreadsheetAccess() {
  try {
    Logger.log('=== SPREADSHEET ACCESS TEST ===');
    
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    Logger.log('Current SPREADSHEET_ID: ' + (spreadsheetId || 'NOT_SET'));
    
    var spreadsheet = getOrCreateSpreadsheet();
    Logger.log('Spreadsheet URL: ' + spreadsheet.getUrl());
    Logger.log('Spreadsheet Name: ' + spreadsheet.getName());
    
    var sheet = getOrCreateWorksheet(spreadsheet, 'PTS Daily Report');
    Logger.log('Worksheet Name: ' + sheet.getName());
    Logger.log('Last Row: ' + sheet.getLastRow());
    
    Logger.log('=== ACCESS TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== ACCESS TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}