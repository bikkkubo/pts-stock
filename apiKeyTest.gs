// apiKeyTest.gs - API Key testing and setup functions

/**
 * Test OpenAI API connection with current key
 */
function testCurrentOpenAIKey() {
  try {
    Logger.log('=== TESTING CURRENT OPENAI API KEY ===');
    
    var openAiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    if (!openAiKey) {
      Logger.log('ERROR: OPENAI_API_KEY is not set');
      return false;
    }
    
    Logger.log('API Key found (length: ' + openAiKey.length + ')');
    Logger.log('API Key prefix: ' + openAiKey.substring(0, 10) + '...');
    
    // Test with a simple request
    var url = 'https://api.openai.com/v1/chat/completions';
    
    var payload = {
      'model': 'gpt-3.5-turbo',
      'messages': [
        {
          'role': 'user',
          'content': 'Hello, this is a test message. Please respond with "API connection successful".'
        }
      ],
      'max_tokens': 50
    };
    
    var options = {
      'method': 'POST',
      'headers': {
        'Authorization': 'Bearer ' + openAiKey,
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    Logger.log('Sending test request to OpenAI...');
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseText = response.getContentText();
    
    Logger.log('Response Code: ' + responseCode);
    
    if (responseCode === 200) {
      Logger.log('SUCCESS: OpenAI API connection successful');
      var data = JSON.parse(responseText);
      if (data.choices && data.choices.length > 0) {
        Logger.log('Response: ' + data.choices[0].message.content);
      }
      return true;
    } else {
      Logger.log('ERROR: OpenAI API connection failed');
      Logger.log('Response: ' + responseText);
      return false;
    }
    
  } catch (error) {
    Logger.log('ERROR: Exception during API test');
    Logger.log('Error: ' + error.toString());
    return false;
  }
}

/**
 * Setup OpenAI API Key - YOU NEED TO EDIT THIS WITH YOUR ACTUAL KEY
 */
function setupOpenAIApiKey() {
  try {
    Logger.log('=== SETTING UP OPENAI API KEY ===');
    
    // ⚠️ IMPORTANT: Replace the placeholder below with your actual OpenAI API key
    var newApiKey = 'sk-proj-YOUR_ACTUAL_API_KEY_HERE';
    
    if (newApiKey === 'sk-proj-YOUR_ACTUAL_API_KEY_HERE') {
      Logger.log('ERROR: You need to edit this function with your actual API key');
      Logger.log('Please edit setupOpenAIApiKey() function in apiKeyTest.gs');
      return false;
    }
    
    // Validate key format
    if (!newApiKey.startsWith('sk-')) {
      Logger.log('ERROR: Invalid API key format. OpenAI keys should start with "sk-"');
      return false;
    }
    
    if (newApiKey.length < 20) {
      Logger.log('ERROR: API key appears too short');
      return false;
    }
    
    // Set the API key in script properties
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', newApiKey);
    
    Logger.log('SUCCESS: OpenAI API key has been set');
    Logger.log('Key prefix: ' + newApiKey.substring(0, 10) + '...');
    
    // Test the new key
    Logger.log('Testing the new API key...');
    var testResult = testCurrentOpenAIKey();
    
    if (testResult) {
      Logger.log('=== OPENAI API KEY SETUP COMPLETED SUCCESSFULLY ===');
      return true;
    } else {
      Logger.log('=== OPENAI API KEY SETUP FAILED - KEY TEST FAILED ===');
      return false;
    }
    
  } catch (error) {
    Logger.log('ERROR: Exception during API key setup');
    Logger.log('Error: ' + error.toString());
    return false;
  }
}

/**
 * Show current API key status
 */
function showApiKeyStatus() {
  try {
    Logger.log('=== API KEY STATUS ===');
    
    // OpenAI API Key
    var openAiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (openAiKey) {
      Logger.log('OpenAI API Key: SET (length: ' + openAiKey.length + ')');
      Logger.log('Key prefix: ' + openAiKey.substring(0, 10) + '...');
    } else {
      Logger.log('OpenAI API Key: NOT SET');
    }
    
    // Nikkei API Key
    var nikkeiKey = PropertiesService.getScriptProperties().getProperty('NIKKEI_API_KEY');
    if (nikkeiKey) {
      Logger.log('Nikkei API Key: SET (length: ' + nikkeiKey.length + ')');
    } else {
      Logger.log('Nikkei API Key: NOT SET (optional)');
    }
    
    Logger.log('=== STATUS CHECK COMPLETED ===');
    
  } catch (error) {
    Logger.log('ERROR: Exception during status check');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Clear all API keys (for cleanup)
 */
function clearAllApiKeys() {
  try {
    Logger.log('=== CLEARING ALL API KEYS ===');
    
    PropertiesService.getScriptProperties().deleteProperty('OPENAI_API_KEY');
    PropertiesService.getScriptProperties().deleteProperty('NIKKEI_API_KEY');
    
    Logger.log('All API keys have been cleared');
    
  } catch (error) {
    Logger.log('ERROR: Exception during API key cleanup');
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Manual API key input function (alternative to setupOpenAIApiKey)
 * Use this if you prefer to input the key manually each time
 */
function inputOpenAIKeyManually() {
  try {
    Logger.log('=== MANUAL API KEY INPUT ===');
    Logger.log('To set your OpenAI API key manually:');
    Logger.log('1. Go to GAS Editor > Settings (gear icon)');
    Logger.log('2. Click "Script properties" tab');
    Logger.log('3. Add property: OPENAI_API_KEY');
    Logger.log('4. Set value to your OpenAI API key (starts with sk-)');
    Logger.log('5. Click "Save script properties"');
    Logger.log('6. Run testCurrentOpenAIKey() to verify');
    
  } catch (error) {
    Logger.log('ERROR: Exception during manual input guide');
    Logger.log('Error: ' + error.toString());
  }
}