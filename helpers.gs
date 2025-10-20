/**
 * @fileoverview Helper functions for WEPhelper library
 * Contains utility functions for document ID retrieval and management
 */

/**
 * Helper function to generate and populate a list of document IDs from a folder.
 * This function retrieves the docIDs tab and MainWEPfolderID setting, then calls
 * getListOfDocIds to populate the tab with document information.
 * 
 * Provides user feedback via toast messages for success and error states.
 * 
 * @returns {boolean} True if successful, false if an error occurred
 * @throws Will log errors to the logging system but won't throw exceptions
 * 
 * @example
 * // Called from menu or another function
 * helper_getListOfDocIds();
 */
function helper_getListOfDocIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Show initial toast to indicate process has started
    ss.toast("Retrieving document list from folder...", "Working", -1);
    
    // Get the docIDs tab
    const docIdTab = ss.getSheetByName("docIDs");
    if (!docIdTab) {
      const errorMsg = "DocID tab is missing. Please create a tab named 'docIDs'.";
      ss.toast(errorMsg, "Error", 5);
      logIt({
        level: "severe",
        theMsg: errorMsg,
        error: new Error(errorMsg)
      });
      return false;
    }
    
    // Get the folder ID from settings
    const docFolderId = getSettingValue("MainWEPfolderID");
    if (!docFolderId || docFolderId.toString().trim() === '') {
      const errorMsg = "WEP Doc folder ID is missing or blank in settings. Please configure MainWEPfolderID.";
      ss.toast(errorMsg, "Error", 5);
      logIt({
        level: "severe",
        theMsg: errorMsg,
        error: new Error(errorMsg)
      });
      return false;
    }
    
    // Attempt to retrieve and populate the document list
    const result = getListOfDocIds(docFolderId, docIdTab);
    
    if (result) {
      // Success - show completion message with document count
      const docCount = docIdTab.getLastRow() - 1; // Subtract header row
      ss.toast(
        `Successfully retrieved ${docCount} document${docCount !== 1 ? 's' : ''} from folder.`, 
        "Complete", 
        3
      );
      
      logIt({
        level: "info",
        theMsg: `Document list retrieved successfully. Found ${docCount} documents.`
      });
      
      return true;
    } else {
      // getListOfDocIds returned false
      const errorMsg = "Failed to retrieve document list. Check folder permissions and ID.";
      ss.toast(errorMsg, "Error", 5);
      return false;
    }
    
  } catch (err) {
    // Catch any unexpected errors
    const errorMsg = `Unexpected error retrieving document list: ${err.message}`;
    ss.toast(errorMsg, "Error", 5);
    
    logIt({
      level: "severe",
      theMsg: "Error in helper_getListOfDocIds",
      error: err
    });
    
    return false;
  }
}


/**
 * Generates a list of document IDs and titles from a specified Google Drive folder.
 * Populates the provided sheet with two columns: document name and document ID.
 * Results are sorted alphabetically by document name.
 * 
 * @param {string} folderId - The ID of the Google Drive folder containing documents
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet where results will be written
 * @returns {boolean} True if successful, false if an error occurred
 * 
 * @example
 * const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("docIDs");
 * const folderId = "1ABC...XYZ";
 * const success = getListOfDocIds(folderId, sheet);
 */
function getListOfDocIds(folderId, sheet) {
  try {
    // Validate inputs
    if (!folderId || typeof folderId !== 'string') {
      throw new Error("Invalid folder ID provided");
    }
    
    if (!sheet) {
      throw new Error("Invalid sheet object provided");
    }
    
    // Attempt to access the folder
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (err) {
      logIt({
        level: "severe",
        theMsg: `Cannot access folder with ID: ${folderId}. Check permissions and ID validity.`,
        error: err
      });
      return false;
    }
    
    // Get all files in the folder
    const files = folder.getFiles();
    const results = [];
    
    // Collect file names and IDs
    while (files.hasNext()) {
      const file = files.next();
      results.push([file.getName(), file.getId()]);
    }
    
    // Check if any files were found
    if (results.length === 0) {
      logIt({
        level: "warning",
        theMsg: `No documents found in folder: ${folder.getName()}`
      });
      
      // Still write headers even if no documents found
      sheet.clear();
      sheet.getRange(1, 1, 1, 2).setValues([['StudLastFirst', 'DocID']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
      
      return true;
    }
    
    // Sort results alphabetically by document name
    results.sort((a, b) => a[0].localeCompare(b[0]));
    
    // Clear the sheet before writing new data
    sheet.clear();
    
    // Write the data to the sheet
    // Start at column determined by getLastColumn() + 1, or column 1 if sheet is empty
    const startCol = sheet.getLastColumn() > 0 ? sheet.getLastColumn() + 1 : 1;
    const dataRange = sheet.getRange(2, startCol, results.length, 2);
    dataRange.setValues(results);
    
    // Set headers
    const headerRange = sheet.getRange(1, startCol, 1, 2);
    headerRange.setValues([['StudLastFirst', 'DocID']]);
    headerRange.setFontWeight('bold');
    
    // Auto-resize columns for better readability
    sheet.autoResizeColumns(startCol, 2);
    
    // Log success
    logIt({
      level: "info",
      theMsg: `Successfully retrieved ${results.length} documents from folder: ${folder.getName()}`
    });
    
    return true;
    
  } catch (err) {
    // Log the error and return false
    logIt({
      level: "severe",
      theMsg: "Error in getListOfDocIds",
      error: err
    });
    return false;
  }
}


/**
 * Checks to see if an argument is missing (undefined or null)
 * 
 * @param {*} arg - The argument to check
 * @returns {boolean} True if the argument is missing (undefined or null), false otherwise
 * 
 * @example
 * if (isMissing(myVariable)) {
 *   console.log("Variable is missing!");
 * }
 */
function isMissing(arg) {
  Logger.log(typeof arg);
  if (arg === undefined || arg === null) {
    return true;
  } else {
    return false;
  }
}


/**
 * Checks if the script execution time has exceeded the limit.
 * Google allows GSuite accounts scripts to run for 30 minutes.
 * Free accounts are limited to 4-6 minutes.
 * 
 * This function checks against a 25-minute limit (1,500,000 milliseconds)
 * to provide a safety buffer before the actual timeout.
 * 
 * @param {Date} start - The start time of the script execution
 * @returns {boolean} True if time limit is exceeded, false otherwise
 * 
 * @example
 * const start = new Date();
 * // ... do work ...
 * if (isTimeUp_(start)) {
 *   console.log("Time limit reached!");
 *   break;
 * }
 */
function isTimeUp_(start) {
  const now = new Date();
  const elapsed = now.getTime() - start.getTime();
  
  // 5 min = 300000, 25 min = 1500000
  return elapsed > 1500000;
}


/**
 * Displays a prompt dialog and returns the user's input.
 * Handles OK, CANCEL, and CLOSE button responses.
 * 
 * @param {GoogleAppsScript.Base.Ui} ui - The UI object from SpreadsheetApp.getUi()
 * @param {string} title - The title of the prompt dialog
 * @param {string} message - The message/question to display in the dialog
 * @returns {string|undefined} The user's response text if OK is clicked with valid input,
 *                             undefined if cancelled, closed, or no input provided
 * 
 * @example
 * const ui = SpreadsheetApp.getUi();
 * const response = getAnswer_(ui, "Enter Name", "Please enter your name:");
 * if (response) {
 *   Logger.log("User entered: " + response);
 * }
 */
function getAnswer_(ui, title, message) {
  const result = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() === ui.Button.OK && result.getResponseText() !== '') {
    return result.getResponseText();
  } else if (result.getSelectedButton() === ui.Button.OK && result.getResponseText() === '') {
    ui.alert("Please enter a value.");
    return undefined;
  } else if (result.getSelectedButton() === ui.Button.CANCEL || 
             result.getSelectedButton() === ui.Button.CLOSE) {
    return undefined;
  }
}