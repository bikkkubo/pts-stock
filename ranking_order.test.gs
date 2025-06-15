// ranking_order.test.gs - Tests for PTS ranking alignment with Kabutan screen display

/**
 * Test PTS ranking alignment with mode parameters
 */
function testPtsRankingAlignment() {
  try {
    Logger.log('=== PTS RANKING ALIGNMENT TEST ===');
    
    // Expected ranking from user: 7603マックハウス, 3726フォーシーズ, 212AＦＥＡＳＹ, 4075ブレインズ
    var expectedTop4 = ['7603', '3726', '212A', '4075'];
    
    // Test gainers with mode=1
    Logger.log('Testing gainers ranking (mode=1)...');
    var gainers = fetchKabutanPtsGainers();
    
    if (gainers && gainers.length > 0) {
      Logger.log('✅ SUCCESS - Fetched ' + gainers.length + ' gainers');
      
      // Verify ranking order (should be descending by diffPercent)
      var isProperOrder = true;
      for (var i = 1; i < gainers.length; i++) {
        if (gainers[i].diffPercent > gainers[i-1].diffPercent) {
          isProperOrder = false;
          break;
        }
      }
      
      if (isProperOrder) {
        Logger.log('✅ GAINERS RANKING - Proper descending order confirmed');
      } else {
        Logger.log('⚠️ GAINERS RANKING - Order may not be properly sorted');
      }
      
      // Check if top 4 matches expected ranking
      var matchesExpected = true;
      for (var i = 0; i < Math.min(4, gainers.length); i++) {
        if (gainers[i].code !== expectedTop4[i]) {
          matchesExpected = false;
          break;
        }
      }
      
      if (matchesExpected) {
        Logger.log('✅ EXPECTED RANKING - Matches user-provided ranking');
      } else {
        Logger.log('⚠️ EXPECTED RANKING - Does not match expected order');
        Logger.log('Expected: ' + expectedTop4.join(', '));
        Logger.log('Actual: ' + gainers.slice(0, 4).map(function(s) { return s.code; }).join(', '));
      }
      
      // Log top 4 for verification
      for (var i = 0; i < Math.min(4, gainers.length); i++) {
        var stock = gainers[i];
        Logger.log('  #' + (i+1) + ': ' + stock.code + ' (' + stock.name + ') +' + stock.diffPercent + '%');
      }
    } else {
      Logger.log('❌ FAIL - No gainers data retrieved');
    }
    
    Logger.log('');
    
    // Test losers with mode=2
    Logger.log('Testing losers ranking (mode=2)...');
    var losers = fetchKabutanPtsLosers();
    
    if (losers && losers.length > 0) {
      Logger.log('✅ SUCCESS - Fetched ' + losers.length + ' losers');
      
      // Verify ranking order (should be ascending by diffPercent, most negative first)
      var isProperOrder = true;
      for (var i = 1; i < losers.length; i++) {
        if (losers[i].diffPercent < losers[i-1].diffPercent) {
          isProperOrder = false;
          break;
        }
      }
      
      if (isProperOrder) {
        Logger.log('✅ LOSERS RANKING - Proper ascending order confirmed');
      } else {
        Logger.log('⚠️ LOSERS RANKING - Order may not be properly sorted');
      }
      
      // Log top 3 for verification
      for (var i = 0; i < Math.min(3, losers.length); i++) {
        var stock = losers[i];
        Logger.log('  #' + (i+1) + ': ' + stock.code + ' (' + stock.name + ') ' + stock.diffPercent + '%');
      }
    } else {
      Logger.log('❌ FAIL - No losers data retrieved');
    }
    
    Logger.log('=== PTS RANKING ALIGNMENT TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== PTS RANKING ALIGNMENT TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test URL parameter inclusion
 */
function testModeParameterUrls() {
  try {
    Logger.log('=== MODE PARAMETER URL TEST ===');
    
    // This is a conceptual test since we can't easily verify URLs
    // In real implementation, the URLs should be:
    // Gainers: https://kabutan.jp/warning/pts_night_price_increase?mode=1
    // Losers: https://kabutan.jp/warning/pts_night_price_decrease?mode=2
    
    Logger.log('Expected URLs:');
    Logger.log('  Gainers: https://kabutan.jp/warning/pts_night_price_increase?mode=1');
    Logger.log('  Losers: https://kabutan.jp/warning/pts_night_price_decrease?mode=2');
    
    // Test that functions can be called without errors
    Logger.log('Testing function availability...');
    
    if (typeof fetchKabutanPtsGainers === 'function') {
      Logger.log('✅ fetchKabutanPtsGainers function available');
    } else {
      Logger.log('❌ fetchKabutanPtsGainers function not found');
    }
    
    if (typeof fetchKabutanPtsLosers === 'function') {
      Logger.log('✅ fetchKabutanPtsLosers function available');
    } else {
      Logger.log('❌ fetchKabutanPtsLosers function not found');
    }
    
    Logger.log('=== MODE PARAMETER URL TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== MODE PARAMETER URL TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test data structure consistency
 */
function testPtsDataStructure() {
  try {
    Logger.log('=== PTS DATA STRUCTURE TEST ===');
    
    var gainers = fetchKabutanPtsGainers();
    
    if (gainers && gainers.length > 0) {
      var stock = gainers[0];
      var requiredFields = ['code', 'name', 'open', 'close', 'diff', 'diffPercent'];
      var validStructure = true;
      
      for (var i = 0; i < requiredFields.length; i++) {
        var field = requiredFields[i];
        if (typeof stock[field] === 'undefined') {
          Logger.log('❌ Missing field: ' + field);
          validStructure = false;
        }
      }
      
      if (validStructure) {
        Logger.log('✅ DATA STRUCTURE - All required fields present');
        Logger.log('  Sample: ' + JSON.stringify(stock));
      } else {
        Logger.log('⚠️ DATA STRUCTURE - Some fields missing');
      }
    } else {
      Logger.log('⚠️ No data available for structure test');
    }
    
    Logger.log('=== PTS DATA STRUCTURE TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== PTS DATA STRUCTURE TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}