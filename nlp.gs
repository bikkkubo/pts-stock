// nlp.gs - Natural language processing functions for embeddings, clustering, and summarization

/**
 * Get OpenAI embeddings for an array of text strings
 * @param {Array} textArray - Array of text strings to embed
 * @return {Array} Array of embedding vectors
 */
function getEmbeddings(textArray) {
  try {
    var openAiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!openAiKey || !openAiKey.startsWith('sk-')) {
      console.error('Invalid or missing OPENAI_API_KEY. Expected format: sk-...');
      Logger.log('Error: OPENAI_API_KEY not found or invalid format in Script Properties');
      return [];
    }
    
    var url = 'https://api.openai.com/v1/embeddings';
    
    var payload = {
      'model': 'text-embedding-3-small',
      'input': textArray
    };
    
    var options = {
      'method': 'POST',
      'headers': {
        'Authorization': 'Bearer ' + openAiKey,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error('OpenAI API returned status: ' + response.getResponseCode());
    }
    
    var data = JSON.parse(response.getContentText());
    
    var embeddings = [];
    for (var i = 0; i < data.data.length; i++) {
      embeddings.push(data.data[i].embedding);
    }
    
    Logger.log('Generated ' + embeddings.length + ' embeddings');
    return embeddings;
    
  } catch (error) {
    Logger.log('Error in getEmbeddings(): ' + error.toString());
    return [];
  }
}

/**
 * Cluster articles using K-means clustering on embeddings
 * @param {Array} articles - Array of article objects
 * @return {Array} Array of clusters, each containing articles
 */
function clusterArticles(articles) {
  try {
    if (!articles || articles.length === 0) {
      return [];
    }
    
    // If only 1-2 articles, return single cluster
    if (articles.length <= 2) {
      return [articles];
    }
    
    // Deduplicate articles first
    var uniqueArticles = deduplicateArticles(articles);
    Logger.log('After deduplication: ' + uniqueArticles.length + ' articles');
    
    if (uniqueArticles.length <= 2) {
      return [uniqueArticles];
    }
    
    // Prepare text for embedding
    var texts = [];
    for (var i = 0; i < uniqueArticles.length; i++) {
      var text = (uniqueArticles[i].title || '') + ' ' + (uniqueArticles[i].content || '');
      texts.push(text.substring(0, 1000)); // Limit text length
    }
    
    // Get embeddings
    var embeddings = getEmbeddings(texts);
    if (embeddings.length === 0) {
      return [uniqueArticles];
    }
    
    // Determine number of clusters: k = sqrt(n) rounded
    var k = Math.max(1, Math.round(Math.sqrt(uniqueArticles.length)));
    Logger.log('Clustering ' + uniqueArticles.length + ' articles into ' + k + ' clusters');
    
    // Perform K-means clustering
    var clusters = kMeansClustering(embeddings, k);
    
    // Group articles by cluster
    var articleClusters = [];
    for (var j = 0; j < k; j++) {
      articleClusters.push([]);
    }
    
    for (var l = 0; l < clusters.length; l++) {
      var clusterIndex = clusters[l];
      if (clusterIndex >= 0 && clusterIndex < k) {
        articleClusters[clusterIndex].push(uniqueArticles[l]);
      }
    }
    
    // Remove empty clusters
    var nonEmptyClusters = [];
    for (var m = 0; m < articleClusters.length; m++) {
      if (articleClusters[m].length > 0) {
        nonEmptyClusters.push(articleClusters[m]);
      }
    }
    
    Logger.log('Created ' + nonEmptyClusters.length + ' non-empty clusters');
    return nonEmptyClusters;
    
  } catch (error) {
    Logger.log('Error in clusterArticles(): ' + error.toString());
    return [articles]; // Return original articles as single cluster on error
  }
}

/**
 * Generate enhanced summary using 3-stage Map-Reduce-Narrative approach with KPIs
 * @param {Array} clusters - Array of article clusters
 * @return {Object} Summary object with narrative, metrics CSV, and sources
 */
function summarizeClusters(clusters) {
  try {
    if (!clusters || clusters.length === 0) {
      return { summary: '', sources: [], metrics: '' };
    }
    
    var openAiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!openAiKey || !openAiKey.startsWith('sk-')) {
      console.error('Invalid or missing OPENAI_API_KEY. Expected format: sk-...');
      Logger.log('Error: OPENAI_API_KEY not found or invalid format in Script Properties');
      return { summary: 'サマリー生成エラー: API キー未設定', sources: [], metrics: '' };
    }
    
    // Flatten all articles for fact extraction
    var allArticles = [];
    var allSources = [];
    
    for (var i = 0; i < clusters.length; i++) {
      var cluster = clusters[i];
      for (var j = 0; j < cluster.length; j++) {
        allArticles.push(cluster[j]);
        if (cluster[j].url && allSources.indexOf(cluster[j].url) === -1) {
          allSources.push(cluster[j].url);
        }
      }
    }
    
    // Extract facts and KPIs from all articles
    Logger.log('Extracting facts from ' + allArticles.length + ' articles');
    var facts = extractFacts(allArticles);
    
    // Stage 1: Map - Structure each cluster into JSON
    Logger.log('Stage 1: Map phase - structuring ' + clusters.length + ' clusters');
    var structuredClusters = [];
    
    for (var k = 0; k < clusters.length; k++) {
      var structuredCluster = mapToStructure(clusters[k], openAiKey);
      if (structuredCluster && structuredCluster.title) {
        structuredClusters.push(structuredCluster);
        Logger.log('Mapped cluster ' + (k + 1) + ': ' + structuredCluster.title);
      }
    }
    
    // Stage 2: Reduce - Aggregate KPIs and merge insights
    Logger.log('Stage 2: Reduce phase - aggregating insights');
    var aggregatedInsights = aggregateKPIs(structuredClusters, facts);
    
    // Stage 3: Narrative - Generate final 3-sentence summary
    Logger.log('Stage 3: Narrative generation with KPIs and outlook');
    var narrativeResult = generateNarrativeWithRetry(aggregatedInsights, openAiKey, 3);
    
    // Format metrics as CSV
    var metricsCSV = formatMetricsAsCSV(facts);
    
    // Limit sources to top 3
    var topSources = allSources.slice(0, 3);
    
    Logger.log('3-stage summary completed: ' + narrativeResult.summary.length + ' chars, ' + metricsCSV.length + ' metrics');
    
    return {
      summary: narrativeResult.summary,
      sources: topSources,
      metrics: metricsCSV
    };
    
  } catch (error) {
    Logger.log('Error in summarizeClusters(): ' + error.toString());
    return { summary: 'サマリー生成エラー', sources: [], metrics: '' };
  }
}

/**
 * Stage 1: Map cluster to structured JSON format
 * @param {Array} cluster - Single cluster of articles
 * @param {string} openAiKey - OpenAI API key
 * @return {Object} Structured cluster data
 */
function mapToStructure(cluster, openAiKey) {
  try {
    if (!cluster || cluster.length === 0) {
      return null;
    }
    
    // Prepare cluster text
    var clusterText = '';
    for (var i = 0; i < cluster.length; i++) {
      var article = cluster[i];
      clusterText += 'Title: ' + (article.title || '') + '\nContent: ' + (article.content || '') + '\n\n';
    }
    
    // Limit text length for API efficiency
    if (clusterText.length > 2500) {
      clusterText = clusterText.substring(0, 2500);
    }
    
    var url = 'https://api.openai.com/v1/chat/completions';
    
    var messages = [
      {
        role: 'system',
        content: 'ニュースクラスターを構造化JSONに変換してください。出力形式：{"title":"要因名","kpi":"数値指標","driver":"主要要因","outlook":"将来見通し"}'
      },
      {
        role: 'user',
        content: '以下のニュースクラスターを構造化JSONに変換してください：\n\n' + clusterText
      }
    ];
    
    var payload = {
      model: 'gpt-4o-mini',
      messages: messages,
      temperature: 0.1,
      max_tokens: 200
    };
    
    var options = {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + openAiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error('OpenAI API returned status: ' + response.getResponseCode());
    }
    
    var data = JSON.parse(response.getContentText());
    var jsonText = data.choices[0].message.content || '';
    
    // Try to parse the JSON response
    try {
      return JSON.parse(jsonText);
    } catch (parseError) {
      Logger.log('Failed to parse structured response: ' + jsonText);
      return {
        title: '情報不足',
        kpi: '',
        driver: '詳細情報の不足により要因特定困難',
        outlook: ''
      };
    }
    
  } catch (error) {
    Logger.log('Error in mapToStructure(): ' + error.toString());
    return null;
  }
}

/**
 * Stage 2: Aggregate KPIs and merge insights
 * @param {Array} structuredClusters - Array of structured cluster objects
 * @param {Object} facts - Extracted facts object
 * @return {Object} Aggregated insights
 */
function aggregateKPIs(structuredClusters, facts) {
  try {
    var aggregated = {
      primaryDrivers: [],
      keyMetrics: [],
      outlookStatements: [],
      industryContext: ''
    };
    
    // Collect primary drivers
    for (var i = 0; i < structuredClusters.length; i++) {
      var cluster = structuredClusters[i];
      if (cluster.driver) {
        aggregated.primaryDrivers.push(cluster.driver);
      }
      if (cluster.outlook) {
        aggregated.outlookStatements.push(cluster.outlook);
      }
    }
    
    // Select best numerical metrics
    if (facts.sales.length > 0) {
      var bestSales = facts.sales.reduce(function(max, current) {
        return current.value > max.value ? current : max;
      });
      aggregated.keyMetrics.push('売上高' + bestSales.value + bestSales.unit);
    }
    
    if (facts.profit.length > 0) {
      var bestProfit = facts.profit.reduce(function(max, current) {
        return Math.abs(current.value) > Math.abs(max.value) ? current : max;
      });
      aggregated.keyMetrics.push(
        (bestProfit.type === 'loss' ? '損失' : '利益') + Math.abs(bestProfit.value) + bestProfit.unit
      );
    }
    
    if (facts.yoyGrowth.length > 0) {
      var avgGrowth = facts.yoyGrowth.reduce(function(sum, item) {
        return sum + item.value;
      }, 0) / facts.yoyGrowth.length;
      aggregated.keyMetrics.push(
        '前年比' + (avgGrowth > 0 ? '+' : '') + avgGrowth.toFixed(1) + '%'
      );
    }
    
    // Simple industry context inference
    if (aggregated.primaryDrivers.length > 0) {
      var allDrivers = aggregated.primaryDrivers.join(' ');
      if (allDrivers.indexOf('決算') >= 0 || allDrivers.indexOf('業績') >= 0) {
        aggregated.industryContext = '決算シーズンの業績評価';
      } else if (allDrivers.indexOf('新商品') >= 0 || allDrivers.indexOf('新事業') >= 0) {
        aggregated.industryContext = '事業拡大期の成長期待';
      } else if (allDrivers.indexOf('市場') >= 0) {
        aggregated.industryContext = '市場環境変化への対応';
      } else {
        aggregated.industryContext = '企業固有要因による変動';
      }
    }
    
    Logger.log('Aggregated insights: ' + aggregated.keyMetrics.length + ' metrics, ' + aggregated.primaryDrivers.length + ' drivers');
    return aggregated;
    
  } catch (error) {
    Logger.log('Error in aggregateKPIs(): ' + error.toString());
    return {
      primaryDrivers: ['情報不足'],
      keyMetrics: [],
      outlookStatements: [],
      industryContext: '要因分析困難'
    };
  }
}

/**
 * Stage 3: Generate narrative with retry logic
 * @param {Object} insights - Aggregated insights
 * @param {string} openAiKey - OpenAI API key
 * @param {number} maxRetries - Maximum retry attempts
 * @return {Object} Final narrative result
 */
function generateNarrativeWithRetry(insights, openAiKey, maxRetries) {
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var result = generateNarrative(insights, openAiKey);
      
      // Validate format (3 sentences, ≤400 chars, has Sources)
      if (validateNarrativeFormat(result.summary)) {
        Logger.log('Narrative generated successfully on attempt ' + attempt);
        return result;
      } else {
        Logger.log('Format validation failed on attempt ' + attempt + ', retrying...');
      }
      
    } catch (error) {
      Logger.log('Error generating narrative on attempt ' + attempt + ': ' + error.toString());
    }
  }
  
  // Fallback if all attempts fail
  return {
    summary: '株価変動要因の詳細分析が困難。複数の要因が複合的に影響している可能性。今後の業績発表や市場動向に注目が必要。',
    validated: false
  };
}

/**
 * Generate narrative summary from aggregated insights
 * @param {Object} insights - Aggregated insights
 * @param {string} openAiKey - OpenAI API key
 * @return {Object} Narrative result
 */
function generateNarrative(insights, openAiKey) {
  try {
    var metricsText = insights.keyMetrics.length > 0 ? insights.keyMetrics.join('、') : '数値情報なし';
    var driversText = insights.primaryDrivers.slice(0, 2).join('、');
    var outlookText = insights.outlookStatements.length > 0 ? insights.outlookStatements[0] : '将来見通し未詳';
    var contextText = insights.industryContext;
    
    var url = 'https://api.openai.com/v1/chat/completions';
    
    var messages = [
      {
        role: 'system',
        content: '株価分析の3文要約を作成してください。第1文：具体的数値を含む事実、第2文：将来見通し、第3文：業界背景説明。必ず400文字以内、最後に"Sources:"は不要です。'
      },
      {
        role: 'user',
        content: '数値指標：' + metricsText + '\n要因：' + driversText + '\n見通し：' + outlookText + '\n業界背景：' + contextText + '\n\n上記情報から3文400文字以内で要約してください。'
      }
    ];
    
    var payload = {
      model: 'gpt-4o-mini',
      messages: messages,
      temperature: 0.2,
      max_tokens: 200
    };
    
    var options = {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + openAiKey,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) {
      throw new Error('OpenAI API returned status: ' + response.getResponseCode());
    }
    
    var data = JSON.parse(response.getContentText());
    var summary = data.choices[0].message.content || '';
    
    // Ensure 400 character limit
    if (summary.length > 400) {
      summary = summary.substring(0, 397) + '...';
    }
    
    return {
      summary: summary,
      validated: true
    };
    
  } catch (error) {
    Logger.log('Error in generateNarrative(): ' + error.toString());
    throw error;
  }
}

/**
 * Validate narrative format
 * @param {string} narrative - Generated narrative
 * @return {boolean} True if format is valid
 */
function validateNarrativeFormat(narrative) {
  if (!narrative || narrative.length === 0) return false;
  if (narrative.length > 400) return false;
  
  // Count sentences (split by Japanese sentence endings)
  var sentences = narrative.split(/[。！？]/).filter(function(s) { return s.trim().length > 0; });
  if (sentences.length < 2 || sentences.length > 3) return false;
  
  return true;
}

/**
 * Format extracted facts as CSV string
 * @param {Object} facts - Extracted facts object
 * @return {string} CSV formatted metrics
 */
function formatMetricsAsCSV(facts) {
  var metrics = [];
  
  if (facts.sales.length > 0) {
    facts.sales.forEach(function(sale) {
      metrics.push('売上高:' + sale.value + sale.unit);
    });
  }
  
  if (facts.profit.length > 0) {
    facts.profit.forEach(function(profit) {
      metrics.push((profit.type === 'loss' ? '損失:' : '利益:') + Math.abs(profit.value) + profit.unit);
    });
  }
  
  if (facts.yoyGrowth.length > 0) {
    facts.yoyGrowth.forEach(function(growth) {
      metrics.push('前年比:' + (growth.value > 0 ? '+' : '') + growth.value + '%');
    });
  }
  
  return metrics.join(',');
}

/**
 * Enforce format requirements: 3 sentences, 400 characters max
 * @param {string} summary - Input summary text
 * @return {string} Format-compliant summary
 */
function enforceFormatRequirements(summary) {
  if (!summary) return '';
  
  // Split into sentences
  var sentences = summary.split(/[。！？]/);
  var validSentences = [];
  
  for (var i = 0; i < sentences.length; i++) {
    var sentence = sentences[i].trim();
    if (sentence.length > 0) {
      validSentences.push(sentence);
    }
  }
  
  // Ensure exactly 3 sentences
  if (validSentences.length > 3) {
    validSentences = validSentences.slice(0, 3);
  }
  
  // Rebuild summary with proper punctuation
  var finalSummary = '';
  for (var j = 0; j < validSentences.length; j++) {
    finalSummary += validSentences[j];
    if (j < validSentences.length - 1) {
      finalSummary += '。';
    } else {
      finalSummary += '。';
    }
  }
  
  // Enforce 400 character limit
  if (finalSummary.length > 400) {
    // Find the last complete sentence within 400 chars
    var truncated = finalSummary.substring(0, 397);
    var lastPeriod = truncated.lastIndexOf('。');
    if (lastPeriod > 0) {
      finalSummary = truncated.substring(0, lastPeriod + 1);
    } else {
      finalSummary = truncated + '...';
    }
  }
  
  return finalSummary;
}

/**
 * Extract financial facts and KPIs from article text using regex patterns
 * @param {Array} articles - Array of article objects
 * @return {Object} Extracted facts with sales, profit, YoY percentages
 */
function extractFacts(articles) {
  var facts = {
    sales: [],
    profit: [],
    yoyGrowth: [],
    targets: [],
    outlook: []
  };
  
  try {
    for (var i = 0; i < articles.length; i++) {
      var article = articles[i];
      var text = (article.title || '') + ' ' + (article.content || '');
      
      // Extract sales figures (億円、百万円、千万円)
      var salesPatterns = [
        /売上高?(?:は|が|：|:)?(?:約|おおむね)?([0-9,]+)(?:億|百万|千万)?円/g,
        /営業収益(?:は|が|：|:)?(?:約|おおむね)?([0-9,]+)(?:億|百万|千万)?円/g,
        /売上(?:は|が|：|:)?([0-9,]+)億円/g
      ];
      
      salesPatterns.forEach(function(pattern) {
        var match;
        while ((match = pattern.exec(text)) !== null) {
          facts.sales.push({
            value: parseFloat(match[1].replace(/,/g, '')),
            unit: match[0].indexOf('億') >= 0 ? '億円' : '百万円',
            source: article.title || 'Unknown'
          });
        }
      });
      
      // Extract profit figures
      var profitPatterns = [
        /(?:営業|経常|純)?利益(?:は|が|：|:)?(?:約|おおむね)?([0-9,]+)(?:億|百万|千万)?円/g,
        /(?:営業|経常|純)?損失(?:は|が|：|:)?(?:約|おおむね)?([0-9,]+)(?:億|百万|千万)?円/g
      ];
      
      profitPatterns.forEach(function(pattern) {
        var match;
        while ((match = pattern.exec(text)) !== null) {
          var isLoss = match[0].indexOf('損失') >= 0;
          facts.profit.push({
            value: parseFloat(match[1].replace(/,/g, '')) * (isLoss ? -1 : 1),
            unit: match[0].indexOf('億') >= 0 ? '億円' : '百万円',
            type: isLoss ? 'loss' : 'profit',
            source: article.title || 'Unknown'
          });
        }
      });
      
      // Extract YoY growth percentages
      var yoyPatterns = [
        /前年(?:同期)?比([0-9.]+)%(?:増|減|成長|下落)/g,
        /前年度比([0-9.]+)%(?:増|減|成長|下落)/g,
        /(?:増収|減収|増益|減益)([0-9.]+)%/g
      ];
      
      yoyPatterns.forEach(function(pattern) {
        var match;
        while ((match = pattern.exec(text)) !== null) {
          var isDecrease = match[0].indexOf('減') >= 0 || match[0].indexOf('下落') >= 0;
          facts.yoyGrowth.push({
            value: parseFloat(match[1]) * (isDecrease ? -1 : 1),
            type: isDecrease ? 'decrease' : 'increase',
            source: article.title || 'Unknown'
          });
        }
      });
      
      // Extract outlook/target statements
      var outlookPatterns = [
        /(?:今期|来期|次期|今年度|来年度)(?:は|の)?(?:予想|見通し|目標)(?:は|が|：|:)?([^。]+)/g,
        /(?:通期|今期)(?:見通し|予想)(?:は|を)?([^。]+)/g,
        /業績予想(?:を)?([^。]+)/g
      ];
      
      outlookPatterns.forEach(function(pattern) {
        var match;
        while ((match = pattern.exec(text)) !== null) {
          facts.outlook.push({
            statement: match[1].trim(),
            source: article.title || 'Unknown'
          });
        }
      });
    }
    
    Logger.log('Extracted facts: ' + facts.sales.length + ' sales, ' + facts.profit.length + ' profit, ' + facts.yoyGrowth.length + ' YoY, ' + facts.outlook.length + ' outlook');
    return facts;
    
  } catch (error) {
    Logger.log('Error extracting facts: ' + error.toString());
    return facts;
  }
}

/**
 * Remove duplicate articles based on title similarity
 * @param {Array} articles - Array of article objects
 * @return {Array} Array of unique articles
 */
function deduplicateArticles(articles) {
  if (!articles || articles.length <= 1) {
    return articles;
  }
  
  var unique = [];
  
  for (var i = 0; i < articles.length; i++) {
    var article = articles[i];
    var isDuplicate = false;
    var title1 = (article.title || '').toLowerCase();
    
    for (var j = 0; j < unique.length; j++) {
      var existingTitle = (unique[j].title || '').toLowerCase();
      
      // Simple similarity check - if titles are very similar, consider duplicate
      var similarity = calculateStringSimilarity(title1, existingTitle);
      if (similarity > 0.8) {
        isDuplicate = true;
        break;
      }
    }
    
    if (!isDuplicate) {
      unique.push(article);
    }
  }
  
  Logger.log('Deduplicated ' + articles.length + ' articles to ' + unique.length + ' unique articles');
  return unique;
}

/**
 * Calculate similarity between two strings using Jaccard similarity
 * @param {string} str1 - First string
 * @param {string} str2 - Second string
 * @return {number} Similarity score between 0 and 1
 */
function calculateStringSimilarity(str1, str2) {
  if (!str1 || !str2) return 0;
  
  var words1 = str1.split(/\s+/);
  var words2 = str2.split(/\s+/);
  
  var set1 = {};
  var set2 = {};
  
  for (var i = 0; i < words1.length; i++) {
    set1[words1[i]] = true;
  }
  
  for (var j = 0; j < words2.length; j++) {
    set2[words2[j]] = true;
  }
  
  var intersection = 0;
  var union = 0;
  
  // Count intersection
  for (var word in set1) {
    if (set2[word]) {
      intersection++;
    }
    union++;
  }
  
  // Add words only in set2
  for (var word in set2) {
    if (!set1[word]) {
      union++;
    }
  }
  
  return union === 0 ? 0 : intersection / union;
}

/**
 * Diagnostic function to test OpenAI API connection and key validation
 * Run this function to verify your API key setup
 * @return {Object} Test results with key status and connection info
 */
function testOpenAIConnection() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    var result = {
      keyExists: !!key,
      keyFormatValid: key && key.startsWith('sk-'),
      keyLength: key ? key.length : 0,
      keyPreview: key ? key.substring(0, 7) + '...' : 'NOT_SET',
      connectionTest: 'NOT_TESTED'
    };
    
    console.log('=== OpenAI API Key診断結果 ===');
    console.log('API Key存在:', result.keyExists);
    console.log('APIキー形式正常:', result.keyFormatValid);
    console.log('APIキー長:', result.keyLength);
    console.log('APIキープレビュー:', result.keyPreview);
    
    if (!result.keyExists) {
      console.log('❌ OPENAI_API_KEY が設定されていません');
      console.log('💡 設定方法: PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", "YOUR_API_KEY_HERE");');
      return result;
    }
    
    if (!result.keyFormatValid) {
      console.log('❌ APIキーの形式が無効です (sk- で始まる必要があります)');
      return result;
    }
    
    // Simple connection test with minimal API call
    try {
      var testResponse = UrlFetchApp.fetch('https://api.openai.com/v1/models', {
        method: 'GET',
        headers: {
          'Authorization': 'Bearer ' + key
        },
        muteHttpExceptions: true
      });
      
      result.connectionTest = testResponse.getResponseCode() === 200 ? 'SUCCESS' : 'FAILED';
      result.responseCode = testResponse.getResponseCode();
      
      if (result.connectionTest === 'SUCCESS') {
        console.log('✅ OpenAI API接続テスト成功');
      } else {
        console.log('❌ OpenAI API接続テスト失敗 (レスポンスコード: ' + result.responseCode + ')');
      }
      
    } catch (connError) {
      result.connectionTest = 'ERROR';
      result.connectionError = connError.toString();
      console.log('❌ 接続テストエラー:', connError.toString());
    }
    
    return result;
    
  } catch (error) {
    console.error('testOpenAIConnection エラー:', error.toString());
    return { error: error.toString() };
  }
}

/**
 * Helper function to set up API key (use this as template)
 * Replace 'YOUR_OPENAI_API_KEY_HERE' with your actual API key
 */
function setupOpenAIApiKey() {
  // ⚠️ PLACEHOLDER - Replace with your actual OpenAI API key
  var apiKey = 'YOUR_OPENAI_API_KEY_HERE'; // Format: sk-proj-...
  
  if (apiKey === 'YOUR_OPENAI_API_KEY_HERE') {
    console.log('❌ プレースホルダーです。実際のAPIキーに置き換えてください');
    return false;
  }
  
  try {
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
    console.log('✅ OPENAI_API_KEY が正常に設定されました');
    
    // Test the connection
    var testResult = testOpenAIConnection();
    return testResult.connectionTest === 'SUCCESS';
    
  } catch (error) {
    console.error('API キー設定エラー:', error.toString());
    return false;
  }
}