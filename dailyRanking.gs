// dailyRanking.gs - Daily PTS ranking update functions

/**
 * Daily PTS ranking update function - triggered at 8:40 AM
 * This will fetch top 10 gainers and losers and add to the ranking sheet
 */
function updateDailyPtsRanking() {
  try {
    Logger.log('Starting daily PTS ranking update...');
    
    // Get current date
    var today = new Date();
    var dateStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
    
    Logger.log('Processing PTS ranking for date: ' + dateStr);
    
    // Fetch PTS data (top 10 gainers and losers)
    var ptsData = fetchPts(dateStr);
    if (!ptsData || ptsData.length === 0) {
      throw new Error('No PTS ranking data retrieved for ' + dateStr);
    }
    
    Logger.log('Retrieved ' + ptsData.length + ' PTS ranking items');
    
    // Process the ranking data into rows
    var rankingRows = [];
    
    // Separate gainers and losers
    var gainers = [];
    var losers = [];
    
    for (var i = 0; i < ptsData.length; i++) {
      var stock = ptsData[i];
      if (stock.diffPercent > 0) {
        gainers.push(stock);
      } else {
        losers.push(stock);
      }
    }
    
    // Sort gainers by percentage (descending) and losers by percentage (ascending)
    gainers.sort(function(a, b) { return b.diffPercent - a.diffPercent; });
    losers.sort(function(a, b) { return a.diffPercent - b.diffPercent; });
    
    // Add gainers (top 10)
    for (var j = 0; j < Math.min(10, gainers.length); j++) {
      var gainer = gainers[j];
      rankingRows.push([
        dateStr,              // A: Date
        'TOP10上昇',           // B: Category
        j + 1,                // C: Rank
        gainer.code,          // D: Code
        gainer.name || '',    // E: Name
        gainer.open,          // F: Open
        gainer.close,         // G: Close
        gainer.diff,          // H: Diff
        gainer.diffPercent,   // I: Diff%
        'https://kabutan.jp/stock/?code=' + gainer.code  // J: URL
      ]);
    }
    
    // Add losers (top 10)
    for (var k = 0; k < Math.min(10, losers.length); k++) {
      var loser = losers[k];
      rankingRows.push([
        dateStr,              // A: Date
        'TOP10下落',          // B: Category
        k + 1,                // C: Rank
        loser.code,           // D: Code
        loser.name || '',     // E: Name
        loser.open,           // F: Open
        loser.close,          // G: Close
        loser.diff,           // H: Diff
        loser.diffPercent,    // I: Diff%
        'https://kabutan.jp/stock/?code=' + loser.code   // J: URL
      ]);
    }
    
    // Write to ranking sheet
    Logger.log('Preparing to write ' + rankingRows.length + ' ranking rows to sheet');
    try {
      updateRankingSheet(rankingRows);
      Logger.log('Successfully wrote ' + rankingRows.length + ' ranking rows to sheet');
    } catch (sheetError) {
      Logger.log('Error writing to ranking sheet: ' + sheetError.toString());
      throw sheetError;
    }
    
    Logger.log('Daily PTS ranking update completed successfully for ' + dateStr);
    
  } catch (error) {
    Logger.log('Error in updateDailyPtsRanking(): ' + error.toString());
    throw error;
  }
}

/**
 * Update the PTS Daily Ranking sheet with new data
 * @param {Array} rankingRows - Array of ranking row data
 */
function updateRankingSheet(rankingRows) {
  try {
    Logger.log('Starting ranking sheet update...');
    
    // Get or create spreadsheet
    var spreadsheet = getOrCreateSpreadsheet();
    Logger.log('Using spreadsheet: ' + spreadsheet.getName());
    
    // Get or create ranking worksheet  
    var sheet = getOrCreateRankingWorksheet(spreadsheet, 'PTS Daily Ranking');
    Logger.log('Using worksheet: ' + sheet.getName());
    
    // Check if headers exist, if not create them
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      // Create headers
      var headers = [
        '日付',        // A: Date
        '分類',        // B: Category (TOP10上昇/TOP10下落)
        '順位',        // C: Rank
        'コード',      // D: Code
        '銘柄名',      // E: Name
        '始値',        // F: Open
        '終値',        // G: Close
        '差額',        // H: Diff
        '騰落率(%)',   // I: Diff%
        'URL'         // J: URL
      ];
      
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // Format headers
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('#FFFFFF');
      
      lastRow = 1;
      Logger.log('Created headers in ranking sheet');
    }
    
    // Insert new rows at the top (after headers)
    if (rankingRows.length > 0) {
      // Insert empty rows at row 2 (right after headers)
      sheet.insertRows(2, rankingRows.length);
      
      // Write ranking data to the newly inserted rows
      var dataRange = sheet.getRange(2, 1, rankingRows.length, rankingRows[0].length);
      dataRange.setValues(rankingRows);
      
      // Format the new data
      formatRankingData(sheet, 2, rankingRows.length);
      
      Logger.log('Inserted ' + rankingRows.length + ' new ranking rows at the top');
    }
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, 10);
    
    Logger.log('Ranking sheet update completed successfully');
    
  } catch (error) {
    Logger.log('Error in updateRankingSheet(): ' + error.toString());
    throw error;
  }
}

/**
 * Get or create the PTS Daily Ranking worksheet
 * @param {Spreadsheet} spreadsheet - The spreadsheet object
 * @param {string} sheetName - Name of the worksheet
 * @return {Sheet} The worksheet object
 */
function getOrCreateRankingWorksheet(spreadsheet, sheetName) {
  try {
    // Try to get existing sheet
    var sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      // Create new sheet
      sheet = spreadsheet.insertSheet(sheetName);
      Logger.log('Created new ranking worksheet: ' + sheetName);
    } else {
      Logger.log('Using existing ranking worksheet: ' + sheetName);
    }
    
    return sheet;
    
  } catch (error) {
    Logger.log('Error in getOrCreateRankingWorksheet(): ' + error.toString());
    throw error;
  }
}

/**
 * Format ranking data with conditional formatting
 * @param {Sheet} sheet - The worksheet
 * @param {number} startRow - Starting row number
 * @param {number} numRows - Number of rows to format
 */
function formatRankingData(sheet, startRow, numRows) {
  try {
    // Get the percentage change column (column I)
    var percentRange = sheet.getRange(startRow, 9, numRows, 1);
    var values = percentRange.getValues();
    
    // Apply conditional formatting based on positive/negative values
    for (var i = 0; i < values.length; i++) {
      var rowNum = startRow + i;
      var percentValue = values[i][0];
      
      if (percentValue > 0) {
        // Positive change - green background
        sheet.getRange(rowNum, 1, 1, 10).setBackground('#E8F5E8');
        sheet.getRange(rowNum, 9, 1, 1).setFontColor('#0F7B0F');
      } else if (percentValue < 0) {
        // Negative change - red background  
        sheet.getRange(rowNum, 1, 1, 10).setBackground('#FFEBE9');
        sheet.getRange(rowNum, 9, 1, 1).setFontColor('#D50000');
      }
    }
    
    // Format percentage columns
    sheet.getRange(startRow, 9, numRows, 1).setNumberFormat('0.00"%"');
    
    // Format price columns (F, G, H)
    sheet.getRange(startRow, 6, numRows, 3).setNumberFormat('#,##0');
    
    Logger.log('Applied formatting to ranking data');
    
  } catch (error) {
    Logger.log('Error in formatRankingData(): ' + error.toString());
  }
}

/**
 * Install daily ranking trigger for 8:40 AM JST
 * Run this function once manually to set up the daily ranking automation
 */
function installDailyRankingTrigger() {
  try {
    // Delete existing ranking triggers first
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'updateDailyPtsRanking') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Create new trigger for weekdays at 08:40 JST
    ScriptApp.newTrigger('updateDailyPtsRanking')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .nearMinute(40)
      .inTimezone('Asia/Tokyo')
      .create();
    
    Logger.log('Daily ranking trigger installed successfully - will run daily at 08:40 JST');
    
  } catch (error) {
    Logger.log('Error installing daily ranking trigger: ' + error.toString());
    throw error;
  }
}

/**
 * Remove daily ranking triggers (for cleanup)
 */
function removeDailyRankingTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'updateDailyPtsRanking') {
        ScriptApp.deleteTrigger(triggers[i]);
        removed++;
      }
    }
    
    Logger.log('Removed ' + removed + ' daily ranking triggers');
    
  } catch (error) {
    Logger.log('Error removing daily ranking triggers: ' + error.toString());
    throw error;
  }
}

/**
 * Test function for daily ranking update
 */
function testDailyRankingUpdate() {
  try {
    Logger.log('=== DAILY RANKING UPDATE TEST ===');
    
    // Test the daily ranking update function
    updateDailyPtsRanking();
    
    Logger.log('=== TEST COMPLETED SUCCESSFULLY ===');
    
  } catch (error) {
    Logger.log('=== TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
  }
}

/**
 * Manual trigger to update ranking for a specific date
 * @param {string} dateStr - Date in YYYY-MM-DD format
 */
function manualRankingUpdate(dateStr) {
  try {
    if (!dateStr) {
      dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    }
    
    Logger.log('Manual ranking update for date: ' + dateStr);
    
    // Temporarily override the date for testing
    var originalFetchPts = this.fetchPts;
    this.fetchPts = function(date) {
      return originalFetchPts.call(this, dateStr);
    };
    
    updateDailyPtsRanking();
    
    // Restore original function
    this.fetchPts = originalFetchPts;
    
    Logger.log('Manual ranking update completed');
    
  } catch (error) {
    Logger.log('Error in manual ranking update: ' + error.toString());
    throw error;
  }
}