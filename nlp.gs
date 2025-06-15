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
 * Generate Japanese summary from clustered articles
 * @param {Array} clusters - Array of article clusters
 * @return {Object} Summary object with summary text and sources
 */
function summarizeClusters(clusters) {
  try {
    if (!clusters || clusters.length === 0) {
      return { summary: '', sources: [] };
    }
    
    var openAiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!openAiKey || !openAiKey.startsWith('sk-')) {
      console.error('Invalid or missing OPENAI_API_KEY. Expected format: sk-...');
      Logger.log('Error: OPENAI_API_KEY not found or invalid format in Script Properties');
      return { summary: 'サマリー生成エラー: API キー未設定', sources: [] };
    }
    
    // Prepare input text from all clusters
    var allText = '';
    var allSources = [];
    
    for (var i = 0; i < clusters.length; i++) {
      var cluster = clusters[i];
      for (var j = 0; j < cluster.length; j++) {
        var article = cluster[j];
        allText += (article.title || '') + ' ' + (article.content || '') + '\n';
        if (article.url && allSources.indexOf(article.url) === -1) {
          allSources.push(article.url);
        }
      }
    }
    
    // Limit text length for API
    if (allText.length > 3000) {
      allText = allText.substring(0, 3000);
    }
    
    var url = 'https://api.openai.com/v1/chat/completions';
    
    var messages = [
      {
        role: 'system',
        content: '株式市場のニュースを分析し、日本語で簡潔な要約を作成してください。要約は3文以内、400文字以内で、株価変動の要因を説明してください。'
      },
      {
        role: 'user', 
        content: '以下のニュース記事を分析し、株価変動要因を3文以内の日本語で要約してください：\n\n' + allText
      }
    ];
    
    var payload = {
      model: 'gpt-4o-mini',
      messages: messages,
      temperature: 0.3,
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
    
    // Ensure summary is within character limit
    if (summary.length > 400) {
      summary = summary.substring(0, 397) + '...';
    }
    
    Logger.log('Generated summary: ' + summary.substring(0, 50) + '...');
    
    return {
      summary: summary,
      sources: allSources.slice(0, 5) // Limit to 5 sources
    };
    
  } catch (error) {
    Logger.log('Error in summarizeClusters(): ' + error.toString());
    return { summary: 'サマリー生成エラー', sources: [] };
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