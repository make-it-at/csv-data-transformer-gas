// ChartURL.gs
// Version 1.2
// Last updated: 2025-03-02

function getImageUrl() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName("【TOP】");
    
    if (!sheet) {
      throw new Error("Sheet '【TOP】' not found. Please check the sheet name.");
    }
    
    var sourceUrl = sheet.getRange("B35").getValue();
    
    if (!sourceUrl) {
      throw new Error("No URL found in cell B35. Please check the cell content.");
    }
    
    var imageUrl = extractImageUrl(sourceUrl);
    sheet.getRange("B36").setValue(imageUrl);
    
    // Add timestamp for logging
    sheet.getRange("C36").setValue(new Date());
    
    LoggerManager.info("Image URL extracted successfully", { url: imageUrl });
  } catch (error) {
    LoggerManager.error("Error in getImageUrl", error);
    // Optionally, write the error to a cell in the spreadsheet
    var errorSheet = spreadsheet.getSheetByName("Errors") || spreadsheet.insertSheet("Errors");
    errorSheet.appendRow([new Date(), "getImageUrl", error.message]);
  }
}

function extractImageUrl(url) {
  try {
    var html = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true
    }).getContentText();
    
    // Extract image URL from the specific structure
    var regex = /\<p class\=\"graph\"\>\<img src\=\"([^\"]+)\"/i;
    var match = regex.exec(html);
    
    if (match && match[1]) {
      LoggerManager.debug("Image URL found", { url: match[1] });
      return match[1];
    }
    
    throw new Error("Image URL pattern not found in the HTML content.");
  } catch (error) {
    LoggerManager.error("Error in extractImageUrl", error);
    return "Error: " + error.message;
  }
}

function createDailyTrigger() {
  try {
    deleteTriggers();
    
    ScriptApp.newTrigger('getImageUrl')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
    
    LoggerManager.info("Daily trigger set for 9:00 AM");
  } catch (error) {
    LoggerManager.error("Error creating daily trigger", error);
    throw error;
  }
}

function deleteTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'getImageUrl') {
        ScriptApp.deleteTrigger(trigger);
        LoggerManager.debug("Trigger deleted", { 
          triggerId: trigger.getUniqueId(),
          handlerFunction: trigger.getHandlerFunction()
        });
      }
    });
  } catch (error) {
    LoggerManager.error("Error deleting triggers", error);
    throw error;
  }
}

function setupTrigger() {
  try {
    createDailyTrigger();
    LoggerManager.info("Trigger setup completed");
  } catch (error) {
    LoggerManager.error("Error in trigger setup", error);
    throw error;
  }
}

// Run this function to test the script manually
function testScript() {
  LoggerManager.info("Starting test script");
  try {
    getImageUrl();
    LoggerManager.info("Test script completed successfully");
  } catch (error) {
    LoggerManager.error("Test script failed", error);
    throw error;
  }
}

function getImageUrlChart() {
  try {
    return getImageUrl();
  } catch (error) {
    LoggerManager.error("Error in getImageUrlChart", error);
    throw error;
  }
}