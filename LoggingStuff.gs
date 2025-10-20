//Updated logging to include custom headers on 10/19/2025

/**
 * @fileoverview Logging utilities for WEPhelper library
 * Provides centralized logging with automatic spreadsheet context
 * Uses MyBBLog library for logging to a central log sheet
 */

/**
 * Logs messages with automatic spreadsheet context identification.
 * Automatically prepends the spreadsheet name and ID to each log message,
 * making it easy to identify which of 85+ spreadsheets generated each log entry.
 * 
 * @param {Object} info - The logging information object
 * @param {string} info.level - The log level: "severe", "warning", "info", "config", "fine", "finer", "finest"
 * @param {string} info.theMsg - The message to be logged
 * @param {Error} [info.error] - Optional error object. If provided, the stack trace will be logged as severe
 * @param {boolean} [info.includeContext=true] - Whether to include spreadsheet context (name/ID). Set to false to disable.
 * 
 * @example
 * // Basic error logging
 * logIt({
 *   level: "severe",
 *   theMsg: "Failed to create WEP documents",
 *   error: err
 * });
 * 
 * @example
 * // Info logging
 * logIt({
 *   level: "info",
 *   theMsg: "Successfully processed 45 documents"
 * });
 * 
 * @example
 * // Logging without context (rare use case)
 * logIt({
 *   level: "warning",
 *   theMsg: "Generic warning message",
 *   includeContext: false
 * });
 */
let logIt = (info) => {
  let log = MyBBLog.getLog({
    sheetName: 'LogHeaders',
    level: 'FINEST', 
    sheetId: '1UNOvucvbduwRzpgWPj0xxzBW7Knh-gkrgqORFuW2VHw',
    displayUserId: 'EMAIL_FULL',
    customFields: {
      spreadsheetName: SpreadsheetApp.getActiveSpreadsheet().getName(),
      spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId()
    }

  });
  
  let message = info.theMsg;
  
  // Add spreadsheet context unless explicitly disabled
  // This helps identify which of 85+ spreadsheets generated this log entry
  if (info.includeContext !== false) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const spreadsheetName = ss.getName();
      const spreadsheetId = ss.getId();
      
      // Prepend context in searchable format: [Name | ID] Message
      message = `[${spreadsheetName} | ${spreadsheetId}] ${info.theMsg}`;
      
    } catch (contextError) {
      // If we can't get spreadsheet context (rare), log anyway with a note
      message = `[Context unavailable] ${info.theMsg}`;
      console.error("Could not retrieve spreadsheet context:", contextError);
    }
  }
  
  // Log the message at the specified level
  log[info.level](message);
  
  // If an error object is provided, log its stack trace as severe
  if (info.error) {
    log.severe(info.error.stack);
  }
};


/**
 * Alias for logIt to support legacy code that uses logError.
 * Scripts from sheets can log to error log sheet using either name.
 * 
 * @function
 * @see logIt
 */
var logError = logIt;


/**
 * Test function to verify logging is working correctly.
 * Logs a test info message with spreadsheet context.
 * Can be called from Apps Script editor to test logging setup.
 * 
 * @example
 * // Run from Apps Script editor to test
 * testLog();
 */
const testLog = () => {
  logIt({
    level: "info",
    theMsg: "Test log entry from testLog function",
  });
};


/**
 * Alias for testLog to support menu/trigger access.
 * @function
 * @see testLog
 */
var testLog1 = testLog;


/**
 * Test function demonstrating error logging with error object.
 * Shows how to log errors with stack traces.
 * 
 * @example
 * // Run from Apps Script editor to test error logging
 * testLogWithError();
 */
const testLogWithError = () => {
  try {
    // Intentionally cause an error for testing
    throw new Error("This is a test error with stack trace");
  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Test error log entry",
      error: err
    });
  }
};


/**
 * Test function demonstrating logging without context.
 * Shows how to disable automatic spreadsheet context.
 * 
 * @example
 * // Run from Apps Script editor to test logging without context
 * testLogWithoutContext();
 */
const testLogWithoutContext = () => {
  logIt({
    level: "info",
    theMsg: "This message has no spreadsheet context",
    includeContext: false
  });
};

