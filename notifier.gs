// notifier.gs - Slack webhook notification functions

/**
 * Send notification to Slack via webhook
 * @param {string} message - Message to send
 * @param {boolean} isError - Whether this is an error notification
 */
function notifySlack(message, isError) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
    if (!webhookUrl) {
      Logger.log('SLACK_WEBHOOK not found in Script Properties - skipping notification');
      return;
    }
    
    var color = isError ? '#ea4335' : '#0f9d58'; // Red for error, green for success
    var emoji = isError ? ':x:' : ':white_check_mark:';
    var title = isError ? 'PTS Daily Report - Error' : 'PTS Daily Report - Success';
    
    // Get current timestamp in JST
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss JST');
    
    // Prepare Slack payload
    var payload = {
      'username': 'PTS Report Bot',
      'icon_emoji': ':chart_with_upwards_trend:',
      'attachments': [
        {
          'color': color,
          'title': title,
          'text': message,
          'fields': [
            {
              'title': 'Timestamp',
              'value': timestamp,
              'short': true
            }
          ],
          'footer': 'Google Apps Script',
          'footer_icon': 'https://developers.google.com/apps-script/images/apps-script-logo.png'
        }
      ]
    };
    
    // Add spreadsheet link for success notifications
    if (!isError) {
      var spreadsheetUrl = getSpreadsheetUrl();
      if (spreadsheetUrl) {
        payload.attachments[0].fields.push({
          'title': 'Spreadsheet',
          'value': '<' + spreadsheetUrl + '|View Report>',
          'short': true
        });
      }
    }
    
    var options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(webhookUrl, options);
    
    if (response.getResponseCode() === 200) {
      Logger.log('Slack notification sent successfully');
    } else {
      Logger.log('Slack notification failed with status: ' + response.getResponseCode());
    }
    
  } catch (error) {
    Logger.log('Error in notifySlack(): ' + error.toString());
  }
}

/**
 * Send a simple text message to Slack
 * @param {string} text - Simple text message
 */
function notifySlackSimple(text) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
    if (!webhookUrl) {
      Logger.log('SLACK_WEBHOOK not found - skipping simple notification');
      return;
    }
    
    var payload = {
      'text': text,
      'username': 'PTS Report Bot',
      'icon_emoji': ':chart_with_upwards_trend:'
    };
    
    var options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(webhookUrl, options);
    
    if (response.getResponseCode() === 200) {
      Logger.log('Simple Slack notification sent successfully');
    } else {
      Logger.log('Simple Slack notification failed with status: ' + response.getResponseCode());
    }
    
  } catch (error) {
    Logger.log('Error in notifySlackSimple(): ' + error.toString());
  }
}

/**
 * Send notification with detailed report statistics
 * @param {Object} stats - Statistics object with counts and metrics
 */
function notifySlackWithStats(stats) {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
    if (!webhookUrl) {
      Logger.log('SLACK_WEBHOOK not found - skipping stats notification');
      return;
    }
    
    var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss JST');
    var spreadsheetUrl = getSpreadsheetUrl();
    
    var fields = [
      {
        'title': 'Processed Symbols',
        'value': stats.symbolCount || 0,
        'short': true
      },
      {
        'title': 'Total Articles',
        'value': stats.articleCount || 0,
        'short': true
      },
      {
        'title': 'Clusters Created',
        'value': stats.clusterCount || 0,
        'short': true
      },
      {
        'title': 'Processing Time',
        'value': stats.processingTime || 'N/A',
        'short': true
      },
      {
        'title': 'Timestamp',
        'value': timestamp,
        'short': false
      }
    ];
    
    if (spreadsheetUrl) {
      fields.push({
        'title': 'Report',
        'value': '<' + spreadsheetUrl + '|View Spreadsheet>',
        'short': false
      });
    }
    
    var payload = {
      'username': 'PTS Report Bot',
      'icon_emoji': ':chart_with_upwards_trend:',
      'attachments': [
        {
          'color': '#0f9d58',
          'title': 'PTS Daily Report - Completed with Statistics',
          'fields': fields,
          'footer': 'Google Apps Script',
          'footer_icon': 'https://developers.google.com/apps-script/images/apps-script-logo.png'
        }
      ]
    };
    
    var options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(webhookUrl, options);
    
    if (response.getResponseCode() === 200) {
      Logger.log('Slack stats notification sent successfully');
    } else {
      Logger.log('Slack stats notification failed with status: ' + response.getResponseCode());
    }
    
  } catch (error) {
    Logger.log('Error in notifySlackWithStats(): ' + error.toString());
  }
}

/**
 * Test Slack webhook connection
 * @return {boolean} True if webhook is working
 */
function testSlackWebhook() {
  try {
    var webhookUrl = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK');
    if (!webhookUrl) {
      Logger.log('SLACK_WEBHOOK not found in Script Properties');
      return false;
    }
    
    var testMessage = 'PTS Report Bot - Connection Test at ' + 
                     Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss JST');
    
    var payload = {
      'text': testMessage,
      'username': 'PTS Report Bot',
      'icon_emoji': ':test_tube:'
    };
    
    var options = {
      'method': 'POST',
      'headers': {
        'Content-Type': 'application/json'
      },
      'payload': JSON.stringify(payload)
    };
    
    var response = UrlFetchApp.fetch(webhookUrl, options);
    
    if (response.getResponseCode() === 200) {
      Logger.log('Slack webhook test successful');
      return true;
    } else {
      Logger.log('Slack webhook test failed with status: ' + response.getResponseCode());
      return false;
    }
    
  } catch (error) {
    Logger.log('Error in testSlackWebhook(): ' + error.toString());
    return false;
  }
}

/**
 * Send notification about trigger installation
 */
function notifyTriggerInstalled() {
  var message = 'PTS Daily Report trigger has been installed successfully. The report will run automatically every weekday at 06:45 JST.';
  notifySlackSimple(message);
}

/**
 * Send notification about trigger removal
 */
function notifyTriggerRemoved() {
  var message = 'PTS Daily Report trigger has been removed. Automatic reports are now disabled.';
  notifySlackSimple(message);
}