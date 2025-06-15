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
    
    // 5. Check API key status
    Logger.log('Checking API key configuration...');
    checkApiKeyStatus();
    
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

/**
 * Setup Nikkei API Key in environment (optional but recommended)
 * @param {string} apiKey - Your Nikkei API Key (optional)
 */
function setupNikkeiApiKey(apiKey) {
  try {
    Logger.log('=== NIKKEI API KEY SETUP ===');
    
    if (apiKey && apiKey !== 'YOUR_NIKKEI_API_KEY_HERE') {
      // Set the provided API key
      PropertiesService.getScriptProperties().setProperty('NIKKEI_API_KEY', apiKey);
      Logger.log('✅ NIKKEI_API_KEY has been set successfully');
      
      // Test the key by attempting a simple API call
      try {
        var testResponse = UrlFetchApp.fetch('https://api.nikkei.com/news/search?q=test&limit=1', {
          method: 'GET',
          headers: {
            'X-API-Key': apiKey,
            'Content-Type': 'application/json'
          },
          muteHttpExceptions: true
        });
        
        if (testResponse.getResponseCode() === 200) {
          Logger.log('✅ Nikkei API connection test successful');
        } else {
          Logger.log('⚠️ Nikkei API key set but connection test failed (Status: ' + testResponse.getResponseCode() + ')');
        }
      } catch (testError) {
        Logger.log('⚠️ Nikkei API key set but connection test failed: ' + testError.toString());
      }
      
    } else {
      // No API key provided - show current status and instructions
      var existingKey = PropertiesService.getScriptProperties().getProperty('NIKKEI_API_KEY');
      
      if (existingKey) {
        Logger.log('✅ NIKKEI_API_KEY is already configured');
        Logger.log('Current key preview: ' + existingKey.substring(0, 8) + '...');
      } else {
        Logger.log('⚠️ NIKKEI_API_KEY is not configured');
        Logger.log('💡 High-quality Nikkei news will be skipped during analysis');
        Logger.log('💡 To set up: setupNikkeiApiKey("YOUR_ACTUAL_API_KEY")');
        Logger.log('💡 Or manually: PropertiesService.getScriptProperties().setProperty("NIKKEI_API_KEY", "YOUR_API_KEY");');
      }
    }
    
    Logger.log('=== NIKKEI API KEY SETUP COMPLETED ===');
    
  } catch (error) {
    Logger.log('=== NIKKEI API KEY SETUP FAILED ===');
    Logger.log('Error: ' + error.toString());
    throw error;
  }
}

/**
 * Check status of all API keys
 */
function checkApiKeyStatus() {
  try {
    Logger.log('=== API KEY STATUS CHECK ===');
    
    // Check OpenAI API Key
    var openaiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (openaiKey && openaiKey.startsWith('sk-')) {
      Logger.log('✅ OPENAI_API_KEY: Configured (' + openaiKey.substring(0, 7) + '...)');
    } else {
      Logger.log('❌ OPENAI_API_KEY: Missing or invalid format');
    }
    
    // Check Nikkei API Key
    var nikkeiKey = PropertiesService.getScriptProperties().getProperty('NIKKEI_API_KEY');
    if (nikkeiKey) {
      Logger.log('✅ NIKKEI_API_KEY: Configured (' + nikkeiKey.substring(0, 8) + '...)');
    } else {
      Logger.log('⚠️ NIKKEI_API_KEY: Not configured (optional - high-quality news will be skipped)');
    }
    
    Logger.log('=== API KEY STATUS CHECK COMPLETED ===');
    
  } catch (error) {
    Logger.log('Error checking API key status: ' + error.toString());
  }
}