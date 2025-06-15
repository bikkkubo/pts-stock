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
 * Generate Japanese summary from clustered articles using Map-Reduce approach
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
    
    var allSources = [];
    var clusterSummaries = [];
    
    // Map phase: Summarize each cluster independently
    Logger.log('Starting Map phase: processing ' + clusters.length + ' clusters');
    for (var i = 0; i < clusters.length; i++) {
      var cluster = clusters[i];
      var clusterSummary = mapClusterSummary(cluster, openAiKey);
      
      if (clusterSummary && clusterSummary.summary) {
        clusterSummaries.push(clusterSummary.summary);
        Logger.log('Map cluster ' + (i + 1) + ': ' + clusterSummary.summary.substring(0, 50) + '...');
      }
      
      // Collect sources from this cluster
      for (var j = 0; j < cluster.length; j++) {
        var article = cluster[j];
        if (article.url && allSources.indexOf(article.url) === -1) {
          allSources.push(article.url);
        }
      }
    }
    
    // Reduce phase: Combine cluster summaries into final summary
    Logger.log('Starting Reduce phase: combining ' + clusterSummaries.length + ' cluster summaries');
    var finalSummary = '';
    
    if (clusterSummaries.length === 1) {
      // Single cluster - use as-is but ensure format compliance
      finalSummary = enforceFormatRequirements(clusterSummaries[0]);
    } else if (clusterSummaries.length > 1) {
      // Multiple clusters - reduce to single summary
      finalSummary = reduceClusterSummaries(clusterSummaries, openAiKey);
    }
    
    // Ensure final format compliance: 3 sentences, 400 characters max
    finalSummary = enforceFormatRequirements(finalSummary);
    
    // Limit sources to top 3 for URL compliance
    var topSources = allSources.slice(0, 3);
    
    Logger.log('Map-Reduce summary completed: ' + finalSummary.length + ' chars, ' + topSources.length + ' sources');
    
    return {
      summary: finalSummary,
      sources: topSources
    };
    
  } catch (error) {
    Logger.log('Error in summarizeClusters(): ' + error.toString());
    return { summary: 'サマリー生成エラー', sources: [] };
  }
}

/**
 * Map phase: Generate summary for a single cluster
 * @param {Array} cluster - Single cluster of articles
 * @param {string} openAiKey - OpenAI API key
 * @return {Object} Cluster summary object
 */
function mapClusterSummary(cluster, openAiKey) {
  try {
    if (!cluster || cluster.length === 0) {
      return { summary: '' };
    }
    
    // Prepare cluster text
    var clusterText = '';
    for (var i = 0; i < cluster.length; i++) {
      var article = cluster[i];
      clusterText += (article.title || '') + ' ' + (article.content || '') + '\n';
    }
    
    // Limit text length for API efficiency
    if (clusterText.length > 2000) {
      clusterText = clusterText.substring(0, 2000);
    }
    
    var url = 'https://api.openai.com/v1/chat/completions';
    
    var messages = [
      {
        role: 'system',
        content: '株式ニュースクラスターを分析し、1つの要因に絞って簡潔に要約してください。必ず3文以内、200文字以内で、株価変動の具体的要因を説明してください。'
      },
      {
        role: 'user',
        content: '以下のニュース群から株価変動要因を1つに絞って3文以内で要約してください：\n\n' + clusterText
      }
    ];
    
    var payload = {
      model: 'gpt-4o-mini',
      messages: messages,
      temperature: 0.2,
      max_tokens: 150
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
    
    return { summary: summary };
    
  } catch (error) {
    Logger.log('Error in mapClusterSummary(): ' + error.toString());
    return { summary: '' };
  }
}

/**
 * Reduce phase: Combine multiple cluster summaries into final summary
 * @param {Array} clusterSummaries - Array of cluster summary strings
 * @param {string} openAiKey - OpenAI API key
 * @return {string} Final reduced summary
 */
function reduceClusterSummaries(clusterSummaries, openAiKey) {
  try {
    var combinedText = clusterSummaries.join('\n\n');
    
    var url = 'https://api.openai.com/v1/chat/completions';
    
    var messages = [
      {
        role: 'system',
        content: '複数の株価変動要因を統合し、最も重要な要因を特定して3文400文字以内で要約してください。必ず3文で構成し、400文字を厳守してください。'
      },
      {
        role: 'user',
        content: '以下の複数の株価変動要因を統合し、最重要要因を3文400文字以内で要約してください：\n\n' + combinedText
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
    return data.choices[0].message.content || '';
    
  } catch (error) {
    Logger.log('Error in reduceClusterSummaries(): ' + error.toString());
    return clusterSummaries[0] || ''; // Fallback to first summary
  }
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