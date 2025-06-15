// pts_ranking_fix.test.gs - Tests for A plan implementation with fetchNightPts

/**
 * Test fetchNightPts function for gainers
 */
function testFetchNightPtsGainers() {
  try {
    Logger.log('=== FETCH NIGHT PTS GAINERS TEST ===');
    
    var gainers = fetchNightPts('increase');
    
    if (gainers && gainers.length > 0) {
      Logger.log('✅ SUCCESS - Fetched ' + gainers.length + ' gainers');
      
      // Check for expected top stocks
      var expectedCodes = ['7603', '3726', '212A', '4075'];
      var foundExpected = 0;
      
      for (var i = 0; i < Math.min(4, gainers.length); i++) {
        var stock = gainers[i];
        Logger.log('  #' + (i+1) + ': ' + stock.code + ' (' + stock.name + ') +' + stock.diffPercent + '%');
        
        if (expectedCodes.indexOf(stock.code) > -1) {
          foundExpected++;
        }
      }
      
      if (foundExpected >= 2) {
        Logger.log('✅ EXPECTED STOCKS - Found ' + foundExpected + '/4 expected stocks in top 4');
      } else {
        Logger.log('⚠️ EXPECTED STOCKS - Only found ' + foundExpected + '/4 expected stocks');
      }
      
      // Validate data structure
      var sample = gainers[0];
      var requiredFields = ['code', 'name', 'open', 'close', 'diff', 'diffPercent'];
      var validStructure = true;
      
      for (var i = 0; i < requiredFields.length; i++) {
        if (typeof sample[requiredFields[i]] === 'undefined') {
          validStructure = false;
          break;
        }
      }
      
      if (validStructure) {
        Logger.log('✅ DATA STRUCTURE - All required fields present');
      } else {
        Logger.log('❌ DATA STRUCTURE - Some required fields missing');
      }
      
    } else {
      Logger.log('❌ FAIL - No gainers data retrieved');
    }
    
    Logger.log('=== FETCH NIGHT PTS GAINERS TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== FETCH NIGHT PTS GAINERS TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test fetchNightPts function for decliners
 */
function testFetchNightPtsDecliners() {
  try {
    Logger.log('=== FETCH NIGHT PTS DECLINERS TEST ===');
    
    var decliners = fetchNightPts('decrease');
    
    if (decliners && decliners.length > 0) {
      Logger.log('✅ SUCCESS - Fetched ' + decliners.length + ' decliners');
      
      // Log top decliners
      for (var i = 0; i < Math.min(3, decliners.length); i++) {
        var stock = decliners[i];
        Logger.log('  #' + (i+1) + ': ' + stock.code + ' (' + stock.name + ') ' + stock.diffPercent + '%');
      }
      
      // Check that percentages are negative
      var hasNegativePercent = false;
      for (var i = 0; i < decliners.length; i++) {
        if (decliners[i].diffPercent < 0) {
          hasNegativePercent = true;
          break;
        }
      }
      
      if (hasNegativePercent) {
        Logger.log('✅ DECLINE VALIDATION - Found negative percentages as expected');
      } else {
        Logger.log('⚠️ DECLINE VALIDATION - No negative percentages found');
      }
      
    } else {
      Logger.log('❌ FAIL - No decliners data retrieved');
    }
    
    Logger.log('=== FETCH NIGHT PTS DECLINERS TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== FETCH NIGHT PTS DECLINERS TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test complete fetchPts pipeline with new method
 */
function testNewFetchPtsPipeline() {
  try {
    Logger.log('=== NEW FETCH PTS PIPELINE TEST ===');
    
    var results = fetchPts('2025-06-15');
    
    if (results && results.length > 0) {
      Logger.log('✅ SUCCESS - Pipeline returned ' + results.length + ' total stocks');
      
      // Check for mix of gainers and decliners
      var posCount = 0;
      var negCount = 0;
      
      for (var i = 0; i < results.length; i++) {
        if (results[i].diffPercent > 0) posCount++;
        if (results[i].diffPercent < 0) negCount++;
      }
      
      Logger.log('  Positive changes: ' + posCount);
      Logger.log('  Negative changes: ' + negCount);
      
      if (posCount > 0 && negCount > 0) {
        Logger.log('✅ BALANCE - Good mix of gainers and decliners');
      } else {
        Logger.log('⚠️ BALANCE - Imbalanced data (all positive or all negative)');
      }
      
      // Log sample of results
      Logger.log('Sample results:');
      for (var i = 0; i < Math.min(5, results.length); i++) {
        var stock = results[i];
        Logger.log('  ' + stock.code + ' (' + stock.name + ') ' + stock.diffPercent + '%');
      }
      
    } else {
      Logger.log('❌ FAIL - Pipeline returned no data');
    }
    
    Logger.log('=== NEW FETCH PTS PIPELINE TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== NEW FETCH PTS PIPELINE TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test safeFetchPts wrapper function
 */
function testSafeFetchPts() {
  try {
    Logger.log('=== SAFE FETCH PTS TEST ===');
    
    var results = safeFetchPts('2025-06-15');
    
    if (results && results.length > 0) {
      Logger.log('✅ SUCCESS - SafeFetchPts returned ' + results.length + ' stocks');
      
      // Verify first few results
      for (var i = 0; i < Math.min(3, results.length); i++) {
        var stock = results[i];
        Logger.log('  #' + (i+1) + ': ' + stock.code + ' (' + stock.name + ') ' + stock.diffPercent + '%');
      }
      
    } else {
      Logger.log('❌ FAIL - SafeFetchPts returned no data');
    }
    
    Logger.log('=== SAFE FETCH PTS TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== SAFE FETCH PTS TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Run all PTS ranking fix tests
 */
function runAllPtsRankingFixTests() {
  Logger.log('=== RUNNING ALL PTS RANKING FIX TESTS ===');
  
  testFetchNightPtsGainers();
  Logger.log('');
  
  testFetchNightPtsDecliners();
  Logger.log('');
  
  testNewFetchPtsPipeline();
  Logger.log('');
  
  testSafeFetchPts();
  
  Logger.log('=== ALL PTS RANKING FIX TESTS COMPLETED ===');
}