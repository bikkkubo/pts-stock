// setup.gs - Setup and initialization functions

/**
 * Complete setup for PTS system including both main report and daily ranking
 * Run this function once to set up all triggers and test the system
 */
function setupPtsSystem() {
  try {
    Logger.log('=== PTS SYSTEM SETUP STARTING ===');
    
    // 1. Install main report trigger (6:45 AM)
    Logger.log('Installing main report trigger...');
    installTriggers();
    
    // 2. Install daily ranking trigger (8:40 AM)
    Logger.log('Installing daily ranking trigger...');
    installDailyRankingTrigger();
    
    // 3. Test spreadsheet access
    Logger.log('Testing spreadsheet access...');
    testSpreadsheetAccess();
    
    // 4. Test daily ranking functionality
    Logger.log('Testing daily ranking update...');
    testDailyRankingUpdate();
    
    Logger.log('=== PTS SYSTEM SETUP COMPLETED SUCCESSFULLY ===');
    Logger.log('Main report will run daily at 06:45 JST');
    Logger.log('Daily ranking will run daily at 08:40 JST');
    
  } catch (error) {
    Logger.log('=== PTS SYSTEM SETUP FAILED ===');
    Logger.log('Error: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    throw error;
  }
}

/**
 * Clean up all PTS triggers
 */
function cleanupPtsSystem() {
  try {
    Logger.log('=== CLEANING UP PTS SYSTEM ===');
    
    // Remove main report triggers
    removeTriggers();
    
    // Remove daily ranking triggers
    removeDailyRankingTriggers();
    
    Logger.log('=== PTS SYSTEM CLEANUP COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== PTS SYSTEM CLEANUP FAILED ===');
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Show current system status
 */
function showPtsSystemStatus() {
  try {
    Logger.log('=== PTS SYSTEM STATUS ===');
    
    // Check triggers
    var triggers = ScriptApp.getProjectTriggers();
    var mainTriggers = 0;
    var rankingTriggers = 0;
    
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      if (trigger.getHandlerFunction() === 'main') {
        mainTriggers++;
        Logger.log('Main report trigger: ' + trigger.getTriggerSource() + ' - ' + trigger.getEventType());
      } else if (trigger.getHandlerFunction() === 'updateDailyPtsRanking') {
        rankingTriggers++;
        Logger.log('Daily ranking trigger: ' + trigger.getTriggerSource() + ' - ' + trigger.getEventType());
      }
    }
    
    Logger.log('Main report triggers: ' + mainTriggers);
    Logger.log('Daily ranking triggers: ' + rankingTriggers);
    
    // Check spreadsheet access
    var spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    Logger.log('Spreadsheet ID: ' + (spreadsheetId || 'NOT_SET'));
    
    // Check API keys
    var openAiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    var nikkeiKey = PropertiesService.getScriptProperties().getProperty('NIKKEI_API_KEY');
    
    Logger.log('OpenAI API Key: ' + (openAiKey ? 'SET' : 'NOT_SET'));
    Logger.log('Nikkei API Key: ' + (nikkeiKey ? 'SET' : 'NOT_SET'));
    
    Logger.log('=== STATUS CHECK COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== STATUS CHECK FAILED ===');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Quick test of the entire system
 */
function quickSystemTest() {
  try {
    Logger.log('=== QUICK SYSTEM TEST ===');
    
    // Test spreadsheet access
    Logger.log('Testing spreadsheet access...');
    var spreadsheet = getOrCreateSpreadsheet();
    Logger.log('Spreadsheet: ' + spreadsheet.getName());
    
    // Test main report sheet
    Logger.log('Testing main report sheet...');
    var mainSheet = getOrCreateWorksheet(spreadsheet, 'PTS Daily Report');
    Logger.log('Main sheet: ' + mainSheet.getName());
    
    // Test ranking sheet
    Logger.log('Testing ranking sheet...');
    var rankingSheet = getOrCreateRankingWorksheet(spreadsheet, 'PTS Daily Ranking');
    Logger.log('Ranking sheet: ' + rankingSheet.getName());
    
    // Test data fetching
    Logger.log('Testing data fetching...');
    var testData = getMockPtsData();
    Logger.log('Mock data: ' + testData.length + ' items');
    
    Logger.log('=== QUICK SYSTEM TEST COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== QUICK SYSTEM TEST FAILED ===');
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}