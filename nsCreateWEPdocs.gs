/**
 * @fileoverview This script contains functions for creating WEP documents from a
 * template using data from a spreadsheet. It is designed to be used as a
 * library in another Google Apps Script project.
 */

// An Immediately Invoked Function Expression (IIFE) to create a private scope
// and expose public functions. This is a robust way to structure a library.
const WEPs = (() => {

  /**
   * Main entry point for creating initial WEP documents.
   * This function retrieves settings from a configuration sheet, validates them,
   * and then calls the core document creation function. It also handles
   * time limits and progress tracking.
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
        logIt({ level: "sever", theMsg: errorMsg });
        return;
      }

      if (!folderID) {
        const errorMsg = "ERROR: Main WEP folder ID is not set. Please configure it in the settings sheet.";
        SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Configuration Error", -1);
        logIt({ level: "sever", theMsg: errorMsg });
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

      switch (theRslts.done) {
        case "true":
          SpreadsheetApp.getActiveSpreadsheet().toast("Got done in time!", "All done!", -1);
          settings.setSetting(CONFIG.KEYS.ROW_NUM, 0);
          break;
        case "false":
          SpreadsheetApp.getActiveSpreadsheet().toast("Exceeded time limit! Please run createDocs again.", "NOT done yet!", -1);
          settings.setSetting(CONFIG.KEYS.ROW_NUM, theRslts.rowNum);
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
   * The core function to create new documents by copying a template and
   * replacing placeholder text with data from a spreadsheet.
   * @param {object} theInfo - An object containing all necessary data.
   * @param {GoogleAppsScript.Spreadsheet.Sheet} theInfo.sheet - The sheet containing the data.
   * @param {Array<Array<any>>} theInfo.theData - The raw data range values.
   * @param {string} theInfo.fileID - The ID of the template document.
   * @param {string} theInfo.folderID - The ID of the destination folder.
   * @param {number} theInfo.startRow - The starting row number for processing.
   * @returns {object} An object with the status ("true", "false", or "error") and row number.
   */
  const createWEPs_ = (theInfo) => {
    const { sheet, theData, fileID, folderID, startRow } = theInfo;

    // Use null checks to handle invalid file/folder IDs gracefully.
    const file = DriveApp.getFileById(fileID);
    const folder = DriveApp.getFolderById(folderID);

    // If the template file is not found, log the error and notify the user.
    if (!file) {
      const errorMsg = `ERROR: Template file with ID '${fileID}' not found. Please check your settings.`;
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
      logIt({ level: "sever", theMsg: errorMsg });
      return { done: "error" };
    }

    // If the destination folder is not found, log the error and notify the user.
    if (!folder) {
      const errorMsg = `ERROR: Destination folder with ID '${folderID}' not found. Please check your settings.`;
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
      logIt({ level: "sever", theMsg: errorMsg });
      return { done: "error" };
    }
    
    // Convert the data range to an array of objects.
    let data = ObjApp.rangeToObjects(theData);
    logIt({ level: "info", theMsg: `Found template file: ${file.getName()} and destination folder: ${folder.getName()}` });

    let docIDs = [];
    let firstRow = startRow + 2; // ObjApp removes header, so add 2 to get to correct spreadsheet row.
    let start = new Date();
    let rslts = {};

    // Use a try/catch block for the main loop to catch any other runtime errors.
    try {
      data.every((row, index) => {
        // Ensure the row has a valid filename before proceeding.
        if (!row.filename) {
          const skipMsg = `Skipping row ${index + 1} due to missing filename.`;
          logIt({ level: "warn", theMsg: skipMsg });
          // Returning true continues the loop.
          return true;
        }
        
        // Create a copy of the template file in the destination folder.
        const newFile = file.makeCopy(row.filename, folder);
        
        // Open the new document to get its body for text replacement.
        const body = DocumentApp.openById(newFile.getId()).getBody();

        // Replace placeholders with data from the row object.
        body.replaceText('<<StudLastFirst>>', row.studlastfirst);
        body.replaceText('<<Grade>>', row.grade);
        body.replaceText('<<StudFirst>>', row.studfirst);
        body.replaceText('<<StudentNumber>>', row.studentnumber);
        body.replaceText('<<GiftedArea>>', row.giftedarea);
        body.replaceText('<<SchoolCode>>', row.schoolcode);
        
        docIDs.push([newFile.getId()]);

        // Check for script execution time limit.
        if (isTimeUp_(start)) {
          rslts = { done: "false", rowNum: index + 1 };
          return false; // Exit the loop.
        }
        return true;
      });
    } catch (err) {
      const errorMsg = `An error occurred during document creation/replacement: ${err.message}`;
      SpreadsheetApp.getActiveSpreadsheet().toast(errorMsg, "Script Error", -1);
      logIt({ level: "sever", theMsg: errorMsg, error: err });
      return { done: "error" };
    }

    // Add a new column for Doc IDs if it doesn't exist.
    let lcol = sheet.getLastColumn();
    if (!(sheet.getRange(1, lcol).getValue() == 'DocID')) {
      sheet.getRange(1, lcol + 1).setValue('DocID');
      lcol += 1;
    }

    // Write the new Doc IDs back to the spreadsheet.
    sheet.getRange(firstRow, lcol, docIDs.length).setValues(docIDs);

    // Final check to see if all documents were processed.
    if (docIDs.length >= data.length) {
      rslts = { done: "true", rowNum: "0" };
    }

    return rslts;
  };

  return {
    createInitialWEPs: createInitialWEPs_,
    create: createWEPs_
  };
})();

// Global variables for the library's public functions.
var createWEPdocs = WEPs;
var createWEPs = WEPs;
