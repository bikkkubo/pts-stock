// ===== writer.gs  mapping part =====
/**
 * Update spreadsheet with PTS data
 * @param {Array} allRows - Array of stock data rows
 */
function updateSheet(allRows) {
  try {
    var spreadsheet = getOrCreateSpreadsheet();
    var sheet = getOrCreateWorksheet(spreadsheet, 'PTS Daily Report');
    
    // Set headers (A-I columns)
    var headers = [
      'コード',      // A: Stock code
      '銘柄名',      // B: Stock name
      '始値',        // C: Open price (previous close)
      '終値',        // D: Close price (PTS price)
      '差額',        // E: Price difference
      '騰落率(%)',   // F: Percentage change
      'AI要約',      // G: AI summary
      'メトリクス',  // H: Metrics
      '情報源'       // I: Source URLs
    ];
    
    // Clear existing data
    if (sheet.getLastRow() > 2) {
      sheet.getRange(3, 1, sheet.getLastRow() - 2, 9).clear();
    }
    
    // Set headers
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    
    if (allRows && allRows.length > 0) {
      // Map data to match A-I column structure
      var values = allRows.map(function(row) {
        return [
          row.code || '',                           // A: コード
          row.name || '',                           // B: 銘柄名
          row.open || 0,                           // C: 始値 (previous close)
          row.close || 0,                          // D: 終値 (PTS price)
          row.diff || 0,                           // E: 差額
          (row.diffPercent || 0) + '%',            // F: 騰落率(%)
          row.summary || '',                       // G: AI要約
          row.metrics || '',                       // H: メトリクス
          (row.sources || []).join('\n')           // I: 情報源
        ];
      });
      
      // Write data to sheet
      sheet.getRange(3, 1, values.length, 9).setValues(values);
      
      // Format percentage column
      sheet.getRange(3, 6, values.length, 1).setNumberFormat('0.00%');
      
      Logger.log('Successfully updated sheet with ' + values.length + ' rows');
    }
    
    // Set column widths for better display
    sheet.setColumnWidth(1, 80);   // Code
    sheet.setColumnWidth(2, 150);  // Name
    sheet.setColumnWidth(3, 80);   // Open
    sheet.setColumnWidth(4, 80);   // Close
    sheet.setColumnWidth(5, 80);   // Diff
    sheet.setColumnWidth(6, 100);  // Percent
    sheet.setColumnWidth(7, 300);  // Summary
    sheet.setColumnWidth(8, 150);  // Metrics
    sheet.setColumnWidth(9, 200);  // Sources
    
  } catch (error) {
    Logger.log('Error in updateSheet(): ' + error.toString());
    throw error;
  }
}

/**
 * Get or create main spreadsheet
 * @return {Spreadsheet} Google Spreadsheet object
 */
function getOrCreateSpreadsheet() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (spreadsheetId) {
    try {
      return SpreadsheetApp.openById(spreadsheetId);
    } catch (error) {
      Logger.log('Spreadsheet not found with ID: ' + spreadsheetId + ', creating new one');
    }
  }
  
  // Create new spreadsheet
  var spreadsheet = SpreadsheetApp.create('PTS Daily Report - ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'));
  var newId = spreadsheet.getId();
  
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', newId);
  Logger.log('Created new spreadsheet with ID: ' + newId);
  
  return spreadsheet;
}

/**
 * Get or create worksheet
 * @param {Spreadsheet} spreadsheet - Google Spreadsheet object
 * @param {string} sheetName - Name of the worksheet
 * @return {Sheet} Google Sheet object
 */
function getOrCreateWorksheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log('Created new worksheet: ' + sheetName);
  }
  
  return sheet;
}