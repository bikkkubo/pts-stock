// writer.gs - Google Sheets integration for writing PTS report data

/**
 * Update Google Spreadsheet with PTS report data
 * @param {Array} rows - Array of row data (each row has 7 columns A:G)
 */
function updateSheet(rows) {
  try {
    Logger.log('updateSheet called with ' + (rows ? rows.length : 0) + ' rows');
    
    if (!rows || rows.length === 0) {
      Logger.log('No data to write to sheet');
      return;
    }
    
    // Get the active spreadsheet or create one if needed
    Logger.log('Getting or creating spreadsheet...');
    var spreadsheet = getOrCreateSpreadsheet();
    Logger.log('Getting or creating worksheet...');
    var sheet = getOrCreateWorksheet(spreadsheet, 'PTS Daily Report');
    
    // Clear existing data (except headers)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 9).clear();
    }
    
    // Set headers if not present
    setupHeaders(sheet);
    
    // Add timestamp to data
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var timestampRow = ['レポート生成日時: ' + timestamp, '', '', '', '', '', '', '', ''];
    
    // Write timestamp row
    sheet.getRange(2, 1, 1, 9).setValues([timestampRow]);
    
    // Write data rows starting from row 3
    if (rows.length > 0) {
      var startRow = 3;
      var range = sheet.getRange(startRow, 1, rows.length, 9);
      range.setValues(rows);
      
      // Format the data
      formatDataRange(sheet, startRow, rows.length);
    }
    
    Logger.log('Successfully wrote ' + rows.length + ' rows to spreadsheet');
    
  } catch (error) {
    Logger.log('Error in updateSheet(): ' + error.toString());
    throw error;
  }
}

/**
 * Get existing spreadsheet or create new one
 * @return {Spreadsheet} Google Spreadsheet object
 */
function getOrCreateSpreadsheet() {
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  Logger.log('Stored SPREADSHEET_ID: ' + (spreadsheetId || 'NOT_SET'));
  
  if (spreadsheetId) {
    try {
      Logger.log('Attempting to open existing spreadsheet with ID: ' + spreadsheetId);
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      Logger.log('Successfully opened existing spreadsheet: ' + spreadsheet.getName());
      return spreadsheet;
    } catch (error) {
      Logger.log('Could not open existing spreadsheet: ' + error.toString());
      Logger.log('Will create new spreadsheet instead');
    }
  }
  
  // Create new spreadsheet
  Logger.log('Creating new spreadsheet...');
  var spreadsheet = SpreadsheetApp.create('PTS Daily Report - ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd'));
  
  // Save the ID for future use
  var newSpreadsheetId = spreadsheet.getId();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', newSpreadsheetId);
  Logger.log('Saved new SPREADSHEET_ID: ' + newSpreadsheetId);
  
  Logger.log('Created new spreadsheet: ' + spreadsheet.getUrl());
  return spreadsheet;
}

/**
 * Get existing worksheet or create new one
 * @param {Spreadsheet} spreadsheet - Spreadsheet object
 * @param {string} sheetName - Name of the worksheet
 * @return {Sheet} Worksheet object
 */
function getOrCreateWorksheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log('Created new worksheet: ' + sheetName);
  }
  
  return sheet;
}

/**
 * Setup column headers in the worksheet
 * @param {Sheet} sheet - Worksheet object
 */
function setupHeaders(sheet) {
  var headers = [
    'コード',        // A: Stock code
    '銘柄名',        // B: Stock name
    '始値',          // C: Open price
    '終値',          // D: Close price
    '差額',          // E: Price difference
    '騰落率(%)',     // F: Percentage change
    'AI要約',        // G: AI-generated summary
    'メトリクス',    // H: MetricsCSV (NEW)
    '情報源'         // I: Source URLs
  ];
  
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // Format headers
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  headerRange.setHorizontalAlignment('center');
}

/**
 * Format the data range for better readability
 * @param {Sheet} sheet - Worksheet object
 * @param {number} startRow - Starting row number
 * @param {number} numRows - Number of rows to format
 */
function formatDataRange(sheet, startRow, numRows) {
  try {
    var dataRange = sheet.getRange(startRow, 1, numRows, 9);
    
    // Set alternating row colors
    for (var i = 0; i < numRows; i++) {
      var rowRange = sheet.getRange(startRow + i, 1, 1, 9);
      if (i % 2 === 0) {
        rowRange.setBackground('#f8f9fa');
      } else {
        rowRange.setBackground('#ffffff');
      }
    }
    
    // Format specific columns
    
    // Column A (Code): Center alignment
    sheet.getRange(startRow, 1, numRows, 1).setHorizontalAlignment('center');
    
    // Column B (Name): Center alignment
    sheet.getRange(startRow, 2, numRows, 1).setHorizontalAlignment('center');
    
    // Columns C, D, E (Prices): Number format with commas
    sheet.getRange(startRow, 3, numRows, 3).setNumberFormat('#,##0.00');
    
    // Column F (Percentage): Percentage format (already in percent, so show as number with %)
    var percentRange = sheet.getRange(startRow, 6, numRows, 1);
    percentRange.setNumberFormat('0.00"%"');
    
    // Color-code percentage changes
    for (var j = 0; j < numRows; j++) {
      var cellRange = sheet.getRange(startRow + j, 6, 1, 1);
      var value = cellRange.getValue();
      if (typeof value === 'number') {
        if (value > 0) {
          cellRange.setFontColor('#0f9d58'); // Green for positive
        } else if (value < 0) {
          cellRange.setFontColor('#ea4335'); // Red for negative
        }
      }
    }
    
    // Column G (Summary): Wrap text
    sheet.getRange(startRow, 7, numRows, 1).setWrap(true);
    
    // Column H (Metrics): Wrap text, smaller font
    var metricsRange = sheet.getRange(startRow, 8, numRows, 1);
    metricsRange.setWrap(true);
    metricsRange.setFontSize(9);
    metricsRange.setHorizontalAlignment('left');
    
    // Column I (Sources): Wrap text, smaller font
    var sourcesRange = sheet.getRange(startRow, 9, numRows, 1);
    sourcesRange.setWrap(true);
    sourcesRange.setFontSize(8);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, 9);
    
    // Set minimum widths for readability
    sheet.setColumnWidth(1, 80);  // Code
    sheet.setColumnWidth(2, 120); // Name
    sheet.setColumnWidth(7, 350); // Summary
    sheet.setColumnWidth(8, 200); // Metrics
    sheet.setColumnWidth(9, 150); // Sources
    
  } catch (error) {
    Logger.log('Error in formatDataRange(): ' + error.toString());
  }
}

/**
 * Get spreadsheet URL for sharing
 * @return {string} Spreadsheet URL
 */
function getSpreadsheetUrl() {
  try {
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (spreadsheetId) {
      return 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/edit';
    }
  } catch (error) {
    Logger.log('Error getting spreadsheet URL: ' + error.toString());
  }
  return '';
}

/**
 * Create a backup copy of the current spreadsheet
 * @return {string} Backup spreadsheet URL
 */
function createBackup() {
  try {
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!spreadsheetId) {
      Logger.log('No spreadsheet ID found for backup');
      return '';
    }
    
    var originalSheet = SpreadsheetApp.openById(spreadsheetId);
    var backupName = 'PTS Daily Report Backup - ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm-ss');
    
    var backup = originalSheet.copy(backupName);
    Logger.log('Created backup: ' + backup.getUrl());
    
    return backup.getUrl();
    
  } catch (error) {
    Logger.log('Error creating backup: ' + error.toString());
    return '';
  }
}