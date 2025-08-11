/**
 * Pushes initial goals data from the spreadsheet to documents via the WEPhelper library
 * Retrieves data from the 'forGoals' sheet starting at cell L1
 */
function pushInitialGoals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forGoals');
    if (!ss) {
      throw new Error("Sheet 'forGoals' not found");
    }
    
    const info = ss.getRange("L1").getDataRegion().getValues();
    if (!info || info.length < 2) {
      throw new Error("No data found or insufficient data in range");
    }
    
    const msg = "Successfully appended initial goals to documents!";
    appendTableInfo(info, msg);
    
  } catch (err) {
    console.error("Error in pushInitialGoals:", err);
    throw err;
  }
}

/**
 * Handles the table appending process with error handling
 * @param {Array<Array>} info - 2D array containing document IDs and table data
 * @param {string} message - Success message to display
 */
function appendTableInfo(info, message) {
  try {
    const result = nsAppendGoals.appendTable(info);
    
    if (result.done) {
      console.info(message);
      logIt({level: "info", theMsg: message});
    } else {
      throw new Error(result.info || "Unknown error occurred while appending tables");
    }
    
  } catch (err) {
    const errorMsg = `Error appending to documents: ${err.message}`;
    console.error(errorMsg);
    logIt({level: "error", theMsg: errorMsg});
    throw err;
  }
}