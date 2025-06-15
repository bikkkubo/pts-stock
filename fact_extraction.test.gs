// fact_extraction.test.gs - Unit tests for KPI extraction functionality

/**
 * Test fact extraction functionality
 */
function testFactExtraction() {
  try {
    Logger.log('=== FACT EXTRACTION TEST ===');
    
    // Mock article data for testing
    var testArticles = [
      {
        title: '決算発表：売上高1200億円で前年比15%増',
        content: '第3四半期の業績は営業利益300億円、純利益250億円となり、前年同期比20%増益を達成しました。来期予想は売上高1400億円を見込んでいます。',
        url: 'https://test1.com'
      },
      {
        title: '業績上方修正のお知らせ',
        content: '今期通期見通しを上方修正し、営業収益800億円、経常利益120億円に引き上げました。前年度比10%の増収を見込みます。',
        url: 'https://test2.com'
      }
    ];
    
    // Test fact extraction
    var facts = extractFacts(testArticles);
    
    Logger.log('Sales extracted: ' + facts.sales.length);
    for (var i = 0; i < facts.sales.length; i++) {
      Logger.log('  Sale: ' + facts.sales[i].value + facts.sales[i].unit + ' from ' + facts.sales[i].source);
    }
    
    Logger.log('Profits extracted: ' + facts.profit.length);
    for (var j = 0; j < facts.profit.length; j++) {
      Logger.log('  Profit: ' + facts.profit[j].value + facts.profit[j].unit + ' (' + facts.profit[j].type + ') from ' + facts.profit[j].source);
    }
    
    Logger.log('YoY Growth extracted: ' + facts.yoyGrowth.length);
    for (var k = 0; k < facts.yoyGrowth.length; k++) {
      Logger.log('  Growth: ' + facts.yoyGrowth[k].value + '% (' + facts.yoyGrowth[k].type + ') from ' + facts.yoyGrowth[k].source);
    }
    
    Logger.log('Outlook statements extracted: ' + facts.outlook.length);
    for (var l = 0; l < facts.outlook.length; l++) {
      Logger.log('  Outlook: ' + facts.outlook[l].statement + ' from ' + facts.outlook[l].source);
    }
    
    // Test CSV formatting
    var metricsCSV = formatMetricsAsCSV(facts);
    Logger.log('Formatted CSV: ' + metricsCSV);
    
    Logger.log('=== FACT EXTRACTION TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== FACT EXTRACTION TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test summary format validation
 */
function testSummaryValidation() {
  try {
    Logger.log('=== SUMMARY VALIDATION TEST ===');
    
    var testCases = [
      {
        narrative: '売上高1200億円で前年比15%増加を達成。今後は新商品投入により更なる成長を目指す。業界全体の好調な市場環境も追い風となる。',
        expected: true
      },
      {
        narrative: '短すぎる文章。',
        expected: false
      },
      {
        narrative: '非常に長い文章で400文字を超える場合のテストです。'.repeat(20),
        expected: false
      },
      {
        narrative: '売上増加。利益向上。業績好調。市場拡大。',
        expected: false  // 4 sentences
      }
    ];
    
    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];
      var result = validateNarrativeFormat(testCase.narrative);
      var status = result === testCase.expected ? 'PASS' : 'FAIL';
      Logger.log('Test ' + (i + 1) + ': ' + status + ' - Expected: ' + testCase.expected + ', Got: ' + result);
      Logger.log('  Text length: ' + testCase.narrative.length + ' chars');
    }
    
    Logger.log('=== SUMMARY VALIDATION TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== SUMMARY VALIDATION TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test complete enhanced summarization pipeline
 */
function testEnhancedSummarization() {
  try {
    Logger.log('=== ENHANCED SUMMARIZATION TEST ===');
    
    // Create mock clusters
    var mockClusters = [
      [
        {
          title: '決算発表：売上高1200億円達成',
          content: '第3四半期決算で売上高1200億円、営業利益300億円を達成。前年同期比15%増収、20%増益となりました。',
          url: 'https://kabutan.jp/news/1'
        }
      ],
      [
        {
          title: '業績予想を上方修正',
          content: '通期業績予想を上方修正し、売上高1400億円、営業利益350億円に引き上げました。',
          url: 'https://tdnet.jp/news/1'
        }
      ]
    ];
    
    Logger.log('Testing enhanced summarization with ' + mockClusters.length + ' clusters...');
    
    // This would normally call the OpenAI API, so we'll just test the structure
    Logger.log('Mock clusters created successfully');
    Logger.log('Cluster 1: ' + mockClusters[0].length + ' articles');
    Logger.log('Cluster 2: ' + mockClusters[1].length + ' articles');
    
    // Test fact extraction on flattened articles
    var allArticles = [];
    for (var i = 0; i < mockClusters.length; i++) {
      for (var j = 0; j < mockClusters[i].length; j++) {
        allArticles.push(mockClusters[i][j]);
      }
    }
    
    var facts = extractFacts(allArticles);
    Logger.log('Facts extracted from mock data:');
    Logger.log('  Sales: ' + facts.sales.length);
    Logger.log('  Profits: ' + facts.profit.length);
    Logger.log('  YoY Growth: ' + facts.yoyGrowth.length);
    
    Logger.log('=== ENHANCED SUMMARIZATION TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== ENHANCED SUMMARIZATION TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}