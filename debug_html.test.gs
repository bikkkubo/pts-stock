// debug_html.test.gs - Debug HTML parsing for PTS ranking

/**
 * Debug HTML content from Kabutan PTS page
 */
function debugKabutanHtml() {
  try {
    Logger.log('=== KABUTAN HTML DEBUG TEST ===');
    
    var url = 'https://kabutan.jp/warning/pts_night_price_increase?mode=1';
    
    var response = UrlFetchApp.fetch(url, {
      'method': 'GET',
      'headers': {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'
      }
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log('❌ Failed to fetch URL: ' + response.getResponseCode());
      return;
    }
    
    var html = response.getContentText();
    Logger.log('✅ HTML fetched successfully, length: ' + html.length);
    
    // Look for the expected stock codes in order
    var expectedCodes = ['7603', '3726', '212A', '4075'];
    
    for (var i = 0; i < expectedCodes.length; i++) {
      var code = expectedCodes[i];
      var position = html.indexOf(code);
      if (position > -1) {
        Logger.log('Found ' + code + ' at position: ' + position);
        
        // Show context around the code
        var start = Math.max(0, position - 200);
        var end = Math.min(html.length, position + 300);
        var context = html.substring(start, end);
        Logger.log('Context for ' + code + ': ' + context.substring(0, 500));
        Logger.log('---');
      } else {
        Logger.log('❌ Code ' + code + ' not found in HTML');
      }
    }
    
    // Look for table structure
    var tableStart = html.indexOf('<table');
    var tableEnd = html.indexOf('</table>');
    
    if (tableStart > -1 && tableEnd > -1) {
      Logger.log('✅ Table found from position ' + tableStart + ' to ' + tableEnd);
      var tableHtml = html.substring(tableStart, tableEnd + 8);
      Logger.log('Table HTML length: ' + tableHtml.length);
      
      // Count rows
      var rowCount = (tableHtml.match(/<tr/g) || []).length;
      Logger.log('Number of table rows: ' + rowCount);
    } else {
      Logger.log('⚠️ No table structure found, looking for other patterns');
      
      // Look for div-based structure
      var divCount = (html.match(/<div/g) || []).length;
      Logger.log('Number of div elements: ' + divCount);
    }
    
    // Test current parsing logic
    Logger.log('Testing current parseKabutanPtsData function:');
    var results = parseKabutanPtsData(html, 'gainers');
    Logger.log('Parsed results count: ' + results.length);
    
    for (var i = 0; i < Math.min(4, results.length); i++) {
      var stock = results[i];
      Logger.log('  Result ' + (i+1) + ': ' + stock.code + ' (' + stock.name + ') ' + stock.diffPercent + '%');
    }
    
    Logger.log('=== KABUTAN HTML DEBUG TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== DEBUG TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Test alternative parsing strategies
 */
function testAlternativeParsing() {
  try {
    Logger.log('=== ALTERNATIVE PARSING TEST ===');
    
    var url = 'https://kabutan.jp/warning/pts_night_price_increase?mode=1';
    var response = UrlFetchApp.fetch(url);
    var html = response.getContentText();
    
    // Strategy 1: Simple text search for codes in order
    Logger.log('Strategy 1: Simple text search');
    var expectedCodes = ['7603', '3726', '212A', '4075'];
    var foundPositions = [];
    
    for (var i = 0; i < expectedCodes.length; i++) {
      var code = expectedCodes[i];
      var pos = html.indexOf(code);
      if (pos > -1) {
        foundPositions.push({ code: code, position: pos });
      }
    }
    
    // Sort by position to see actual order in HTML
    foundPositions.sort(function(a, b) { return a.position - b.position; });
    
    Logger.log('Codes found in HTML order:');
    for (var i = 0; i < foundPositions.length; i++) {
      Logger.log('  ' + (i+1) + ': ' + foundPositions[i].code + ' at position ' + foundPositions[i].position);
    }
    
    // Strategy 2: Look for stock code pattern with href
    Logger.log('Strategy 2: href pattern search');
    var hrefPattern = /href="\/stock\/\?code=(\d{4}[A-Z]?)"/g;
    var hrefMatches = [];
    var match;
    
    while ((match = hrefPattern.exec(html)) !== null) {
      hrefMatches.push({ code: match[1], position: match.index });
    }
    
    Logger.log('Found ' + hrefMatches.length + ' href stock code links');
    for (var i = 0; i < Math.min(10, hrefMatches.length); i++) {
      Logger.log('  ' + (i+1) + ': ' + hrefMatches[i].code + ' at position ' + hrefMatches[i].position);
    }
    
    Logger.log('=== ALTERNATIVE PARSING TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== ALTERNATIVE PARSING TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}