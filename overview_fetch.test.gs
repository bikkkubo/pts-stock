// overview_fetch.test.gs - Tests for company overview fetching functionality

/**
 * Test company overview fetching from Kabutan
 */
function testOverviewFetch() {
  try {
    Logger.log('=== COMPANY OVERVIEW FETCH TEST ===');
    
    // Test with known stock codes
    var testCodes = ['7603', '4755', '3661', '2345'];
    
    for (var i = 0; i < testCodes.length; i++) {
      var code = testCodes[i];
      Logger.log('Testing overview fetch for: ' + code);
      
      var overview = getCompanyOverview(code);
      
      if (overview && overview.length > 0) {
        Logger.log('✅ SUCCESS - ' + code + ': ' + overview.substring(0, 100) + '...');
        Logger.log('  Full length: ' + overview.length + ' characters');
      } else {
        Logger.log('⚠️ NO OVERVIEW - ' + code + ': No overview text found');
      }
      
      // Wait 1 second between requests to avoid rate limiting
      Utilities.sleep(1000);
    }
    
    Logger.log('=== COMPANY OVERVIEW FETCH TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== COMPANY OVERVIEW FETCH TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test company name prefix removal
 */
function testCompanyNamePrefixRemoval() {
  try {
    Logger.log('=== COMPANY NAME PREFIX REMOVAL TEST ===');
    
    var testCases = [
      {
        input: 'マックハウス：売上高1200億円で前年比15%増加を達成。',
        companyName: 'マックハウス',
        expected: '売上高1200億円で前年比15%増加を達成。'
      },
      {
        input: 'エアトリ:オンライン旅行事業が好調で業績上方修正。',
        companyName: 'エアトリ',
        expected: 'オンライン旅行事業が好調で業績上方修正。'
      },
      {
        input: '決算発表で売上高が前年比20%増となった。',
        companyName: 'テスト会社',
        expected: '決算発表で売上高が前年比20%増となった。'
      }
    ];
    
    for (var i = 0; i < testCases.length; i++) {
      var testCase = testCases[i];
      var result = removeCompanyNamePrefix(testCase.input, testCase.companyName);
      var status = result === testCase.expected ? 'PASS' : 'FAIL';
      
      Logger.log('Test ' + (i + 1) + ': ' + status);
      Logger.log('  Input: ' + testCase.input);
      Logger.log('  Expected: ' + testCase.expected);
      Logger.log('  Got: ' + result);
      Logger.log('');
    }
    
    Logger.log('=== COMPANY NAME PREFIX REMOVAL TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== COMPANY NAME PREFIX REMOVAL TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test complete enhanced summarization with overview
 */
function testSummarizationWithOverview() {
  try {
    Logger.log('=== SUMMARIZATION WITH OVERVIEW TEST ===');
    
    // Mock data
    var mockSymbol = {
      code: '7603',
      name: 'マックハウス',
      open: 115,
      close: 165,
      diff: 50,
      diffPercent: 43.48
    };
    
    // Mock clusters
    var mockClusters = [
      [
        {
          title: '決算発表：売上高1200億円達成',
          content: '第3四半期決算で売上高1200億円、営業利益300億円を達成。前年同期比15%増収となりました。',
          url: 'https://kabutan.jp/news/1'
        }
      ]
    ];
    
    // Get real overview
    var overview = getCompanyOverview(mockSymbol.code);
    Logger.log('Retrieved overview: ' + (overview ? overview.substring(0, 100) + '...' : 'No overview'));
    
    if (overview) {
      Logger.log('Overview integration test would use this data:');
      Logger.log('  Symbol: ' + mockSymbol.code + ' (' + mockSymbol.name + ')');
      Logger.log('  Clusters: ' + mockClusters.length);
      Logger.log('  Overview: ' + overview.length + ' chars');
    } else {
      Logger.log('⚠️ No overview available for testing');
    }
    
    Logger.log('=== SUMMARIZATION WITH OVERVIEW TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== SUMMARIZATION WITH OVERVIEW TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}