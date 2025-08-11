/**
 * @fileoverview This script contains functions for creating WEP documents from a
 * template using data from a spreadsheet. It is designed to be used as a
 * library in another Google Apps Script project.
 * 
 * Performance improvements include:
 * - Batch text replacements for better efficiency
 * - Robust error handling that continues processing on failures
 * - Detailed validation of required fields and headers
 * - Comprehensive error reporting with Failed_Documents sheet
 * - Proper document lifecycle management
 */

// An Immediately Invoked Function Expression (IIFE) to create a private scope
// and expose public functions. This is a robust way to structure a library.
const WEPs = (() => {

  /**
   * Required column headers for WEP document creation
   * @constant {Array<string>}
   * @private
   */
  const REQUIRED_HEADERS = [
    'filename',
    'studlastfirst', 
    'grade',
    'studfirst',
    'studentnumber',
    'giftedarea',
    'schoolcode'
  ];

  /**
   * Placeholder mappings for text replacement in template documents
   * @constant {Object}
   * @private
   */
  const PLACEHOLDER_MAP = {
    '<<StudLastFirst>>': 'studlastfirst',
    '<<Grade>>': 'grade',
    '<<StudFirst>>': 'studfirst',
    '<<StudentNumber>>': 'studentnumber',
    '<<GiftedArea>>': 'giftedarea',
    '<<SchoolCode>>': 'schoolcode'
  };

  /**
   * Main entry point for creating initial WEP documents.
   * This function retrieves settings from a configuration sheet, validates them,
   * and then calls the core document creation function. It also handles
   * time limits and progress tracking with robust error handling.
   * @returns {void}
   */
  const createInitialWEPs_ = () => {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast("Creating Initial WEPs!", "Started!", -1);

      // Get the active sheet and configuration settings.
      const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forDocs');
      const settings = getSettingsInstance(CONFIG.SHEET_NAME);

      // Retrieve and validate the necessary settings.
      const fileID = settings.getSetting(CONFIG.KEYS.WEP_TEMPLATE_ID);
      const folderID = settings.getSetting(CONFIG.KEYS.MAIN_WEP_FOLDER_ID);
      const startRow = settings.getSetting(CONFIG.KEYS.ROW_NUM) || 0;

      // Check if critical settings are missing.
      if (!fileID) {
        const errorMsg = "ERROR: WEP template ID is not set. Please configure it in the settings sheet.";
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Configuration Error", -1);
        logIt({ level: "severe", theMsg: errorMsg });
        return;
      }

      if (!folderID) {
        const errorMsg = "ERROR: Main WEP folder ID is not set. Please configure it in the settings sheet.";
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Configuration Error", -1);
        logIt({ level: "severe", theMsg: errorMsg });
        return;
      }
      
      const infoObj = {
        sheet: ss,
        theData: ss.getDataRange().getValues(),
        fileID,
        folderID,
        startRow
      };

      const theRslts = createWEPs_(infoObj);
      console.log(theRslts);

      // Handle results with improved messaging
      switch (theRslts.done) {
        case "true":
          const successMsg = `All documents created successfully! ${theRslts.totalProcessed} documents processed.`;
          SpreadsheetApp.getActiveSpreadsheet().toast(successMsg, "All done!", -1);
          logIt({ level: "info", theMsg: successMsg });
          settings.setSetting(CONFIG.KEYS.ROW_NUM, 0);
          break;
        case "false":
          const partialMsg = `Time limit exceeded! ${theRslts.totalProcessed} documents processed. Please run createDocs again to continue.`;
          SpreadsheetApp.getActiveSpreadsheet().toast(partialMsg, "Partial completion!", -1);
          logIt({ level: "warn", theMsg: partialMsg });
          settings.setSetting(CONFIG.KEYS.ROW_NUM, theRslts.rowNum);
          break;
        case "partial":
          const mixedMsg = `Process completed with some failures. ${theRslts.successful} successful, ${theRslts.failed} failed. Check Failed_Documents sheet for details.`;
          SpreadsheetApp.getActiveSpreadsheet().toast(mixedMsg, "Completed with errors!", -1);
          logIt({ level: "warn", theMsg: mixedMsg });
          settings.setSetting(CONFIG.KEYS.ROW_NUM, 0);
          break;
        case "error":
          // The error was already logged and toasted in the createWEPs_ function,
          // so we just stop here.
          break;
      }
    } catch (err) {
      const errorMsg = `An unexpected error occurred in createInitialWEPs: ${err.message}`;
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
      logIt({ level: "severe", theMsg: errorMsg, error: err });
    }
  };

  /**
   * Validates that all required headers are present in the spreadsheet
   * @param {Array<string>} headers - The header row from the spreadsheet
   * @returns {Object} Validation result with isValid boolean and missing headers array
   * @private
   */
  const validateHeaders_ = (headers) => {
    const lowerCaseHeaders = headers.map(h => h.toString().toLowerCase().trim());
    const missingHeaders = [];
    
    for (const requiredHeader of REQUIRED_HEADERS) {
      if (!lowerCaseHeaders.includes(requiredHeader.toLowerCase())) {
        missingHeaders.push(requiredHeader);
      }
    }
    
    return {
      isValid: missingHeaders.length === 0,
      missingHeaders: missingHeaders
    };
  };

  /**
   * Validates that a data row has all required fields populated
   * @param {Object} rowData - The row data object
   * @param {number} rowIndex - The row index for error reporting
   * @returns {Object} Validation result with isValid boolean and missing fields array
   * @private
   */
  const validateRowData_ = (rowData, rowIndex) => {
    const missingFields = [];
    
    for (const field of REQUIRED_HEADERS) {
      const value = rowData[field];
      if (!value || value.toString().trim() === '') {
        missingFields.push(field);
      }
    }
    
    return {
      isValid: missingFields.length === 0,
      missingFields: missingFields,
      rowIndex: rowIndex
    };
  };

  /**
   * Creates or updates the Failed_Documents sheet with error information
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The active spreadsheet
   * @param {Array<Object>} failedDocuments - Array of failed document information
   * @private
   */
  const createFailedDocumentsSheet_ = (spreadsheet, failedDocuments) => {
    try {
      // Check if Failed_Documents sheet already exists
      let failedSheet = spreadsheet.getSheetByName('Failed_Documents');
      
      // If sheet exists, clear it; otherwise create it
      if (failedSheet) {
        failedSheet.clear();
      } else {
        failedSheet = spreadsheet.insertSheet('Failed_Documents');
      }
      
      // Set up headers
      const headers = ['Row Number', 'Filename', 'Student Name', 'Error Type', 'Error Details', 'Timestamp'];
      failedSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      failedSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      
      // Add failed document data
      if (failedDocuments.length > 0) {
        const failedData = failedDocuments.map(doc => [
          doc.rowNumber,
          doc.filename || 'N/A',
          doc.studentName || 'N/A',
          doc.errorType,
          doc.errorDetails,
          new Date().toLocaleString()
        ]);
        
        failedSheet.getRange(2, 1, failedData.length, headers.length).setValues(failedData);
      }
      
      // Auto-resize columns
      failedSheet.autoResizeColumns(1, headers.length);
      
      logIt({ level: "info", theMsg: `Failed_Documents sheet updated with ${failedDocuments.length} failed records.` });
      
    } catch (err) {
      logIt({ level: "error", theMsg: `Failed to create/update Failed_Documents sheet: ${err.message}` });
    }
  };

  /**
   * Performs batch text replacement in a document body
   * @param {GoogleAppsScript.Document.Body} body - The document body
   * @param {Object} rowData - The data to replace placeholders with
   * @private
   */
  const performTextReplacements_ = (body, rowData) => {
    // Perform all text replacements in a single pass through the document
    for (const [placeholder, fieldName] of Object.entries(PLACEHOLDER_MAP)) {
      const value = rowData[fieldName] || '';
      body.replaceText(placeholder, value.toString());
    }
  };

  /**
   * The core function to create new documents by copying a template and
   * replacing placeholder text with data from a spreadsheet.
   * Now includes robust error handling, validation, and detailed reporting.
   * 
   * @param {object} theInfo - An object containing all necessary data.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} theInfo.sheet - The sheet containing the data.
   * @param {Array<Array<any>>} theInfo.theData - The raw data range values.
   * @param {string} theInfo.fileID - The ID of the template document.
   * @param {string} theInfo.folderID - The ID of the destination folder.
   * @param {number} theInfo.startRow - The starting row number for processing.
   * @returns {object} An object with detailed results including success/failure counts and status
   */
  const createWEPs_ = (theInfo) => {
    const { sheet, theData, fileID, folderID, startRow } = theInfo;
    
    // Initialize results tracking
    const results = {
      done: "error",
      successful: 0,
      failed: 0,
      totalProcessed: 0,
      rowNum: startRow,
      failedDocuments: []
    };

    try {
      // Validate template file and destination folder
      let file, folder;
      
      try {
        file = DriveApp.getFileById(fileID);
      } catch (err) {
        const errorMsg = `ERROR: Template file with ID '${fileID}' not found. Please check your settings.`;
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
        logIt({ level: "severe", theMsg: errorMsg });
        return results;
      }

      try {
        folder = DriveApp.getFolderById(folderID);
      } catch (err) {
        const errorMsg = `ERROR: Destination folder with ID '${folderID}' not found. Please check your settings.`;
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
        logIt({ level: "severe", theMsg: errorMsg });
        return results;
      }

      // Validate headers before processing
      if (theData.length < 2) {
        const errorMsg = "ERROR: No data found in spreadsheet or insufficient data.";
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Data Error", -1);
        logIt({ level: "severe", theMsg: errorMsg });
        return results;
      }

      const headerValidation = validateHeaders_(theData[0]);
      if (!headerValidation.isValid) {
        const errorMsg = `ERROR: Missing required headers: ${headerValidation.missingHeaders.join(', ')}. Please check your spreadsheet headers.`;
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Header Validation Error", -1);
        logIt({ level: "severe", theMsg: errorMsg });
        return results;
      }
      
      // Convert the data range to an array of objects
      let data = ObjApp.rangeToObjects(theData);
      logIt({ level: "info", theMsg: `Found template file: ${file.getName()} and destination folder: ${folder.getName()}. Starting document creation for ${data.length} records.` });

      let docIDs = [];
      let firstRow = startRow + 2; // ObjApp removes header, so add 2 to get to correct spreadsheet row
      let start = new Date();

      // Process each row with robust error handling
      data.every((row, index) => {
        const currentRowNumber = index + firstRow;
        results.totalProcessed++;
        
        try {
          // Validate row data before processing
          const rowValidation = validateRowData_(row, currentRowNumber);
          if (!rowValidation.isValid) {
            results.failed++;
            results.failedDocuments.push({
              rowNumber: currentRowNumber,
              filename: row.filename || 'Unknown',
              studentName: row.studlastfirst || 'Unknown',
              errorType: 'Validation Error',
              errorDetails: `Missing required fields: ${rowValidation.missingFields.join(', ')}`
            });
            
            logIt({ 
              level: "warn", 
              theMsg: `Row ${currentRowNumber}: Validation failed for ${row.studlastfirst || 'Unknown'} - Missing: ${rowValidation.missingFields.join(', ')}` 
            });
            
            // Continue to next row
            return true;
          }
          
          // Create a copy of the template file in the destination folder
          let newFile;
          try {
            newFile = file.makeCopy(row.filename, folder);
          } catch (err) {
            results.failed++;
            results.failedDocuments.push({
              rowNumber: currentRowNumber,
              filename: row.filename,
              studentName: row.studlastfirst,
              errorType: 'File Creation Error',
              errorDetails: `Failed to create copy: ${err.message}`
            });
            
            logIt({ 
              level: "error", 
              theMsg: `Row ${currentRowNumber}: Failed to create document for ${row.studlastfirst} - ${err.message}` 
            });
            
            return true; // Continue processing
          }
          
          // Open the new document and perform text replacements
          let doc;
          try {
            doc = DocumentApp.openById(newFile.getId());
            const body = doc.getBody();
            
            // Perform all text replacements efficiently
            performTextReplacements_(body, row);
            
            // Properly save and close the document
            doc.saveAndClose();
            
            // Track successful creation
            docIDs.push([newFile.getId()]);
            results.successful++;
            
          } catch (err) {
            results.failed++;
            results.failedDocuments.push({
              rowNumber: currentRowNumber,
              filename: row.filename,
              studentName: row.studlastfirst,
              errorType: 'Text Replacement Error',
              errorDetails: `Failed to update document content: ${err.message}`
            });
            
            logIt({ 
              level: "error", 
              theMsg: `Row ${currentRowNumber}: Failed to update content for ${row.studlastfirst} - ${err.message}` 
            });
            
            // Clean up the created file if text replacement failed
            try {
              if (newFile) {
                DriveApp.getFileById(newFile.getId()).setTrashed(true);
              }
            } catch (cleanupErr) {
              logIt({ level: "warn", theMsg: `Failed to cleanup broken document: ${cleanupErr.message}` });
            }
            
            return true; // Continue processing
          }
          
        } catch (err) {
          results.failed++;
          results.failedDocuments.push({
            rowNumber: currentRowNumber,
            filename: row.filename || 'Unknown',
            studentName: row.studlastfirst || 'Unknown',
            errorType: 'Unexpected Error',
            errorDetails: err.message
          });
          
          logIt({ 
            level: "error", 
            theMsg: `Row ${currentRowNumber}: Unexpected error processing ${row.studlastfirst || 'Unknown'} - ${err.message}` 
          });
          
          return true; // Continue processing
        }

        // Check for script execution time limit
        if (isTimeUp_(start)) {
          results.done = "false";
          results.rowNum = index + 1;
          return false; // Exit the loop
        }
        
        return true; // Continue processing
      });

      // Write successful document IDs back to the spreadsheet
      if (docIDs.length > 0) {
        try {
          // Add a new column for Doc IDs if it doesn't exist
          let lcol = sheet.getLastColumn();
          if (!(sheet.getRange(1, lcol).getValue() == 'DocID')) {
            sheet.getRange(1, lcol + 1).setValue('DocID');
            lcol += 1;
          }

          // Write the new Doc IDs back to the spreadsheet
          sheet.getRange(firstRow, lcol, docIDs.length).setValues(docIDs);
        } catch (err) {
          logIt({ level: "error", theMsg: `Failed to write DocIDs to spreadsheet: ${err.message}` });
        }
      }

      // Create Failed_Documents sheet if there were failures
      if (results.failedDocuments.length > 0) {
        createFailedDocumentsSheet_(SpreadsheetApp.getActiveSpreadsheet(), results.failedDocuments);
      }

      // Determine final status
      if (results.done !== "false") { // Not stopped due to time limit
        if (results.failed === 0) {
          results.done = "true"; // All successful
        } else if (results.successful > 0) {
          results.done = "partial"; // Mixed results
        } else {
          results.done = "error"; // All failed
        }
      }

      // Log final summary
      const summaryMsg = `Document creation completed. Total: ${results.totalProcessed}, Successful: ${results.successful}, Failed: ${results.failed}`;
      logIt({ level: "info", theMsg: summaryMsg });

    } catch (err) {
      const errorMsg = `Critical error in createWEPs_: ${err.message}`;
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
      logIt({ level: "severe", theMsg: errorMsg, error: err });
      results.done = "error";
    }

    return results;
  };

  return {
    createInitialWEPs: createInitialWEPs_,
    create: createWEPs_
  };
})();

// Global variables for the library's public functions.
var createWEPdocs = WEPs;
var createWEPs = WEPs;