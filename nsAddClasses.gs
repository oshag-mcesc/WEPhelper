/**
 * Reusable Error Tracker Class
 * Tracks processing errors and manages error reporting to a sheet
 */
class DocumentProcessingTracker {
  constructor(errorSheetName = "Errors") {
    this.errorSheetName = errorSheetName;
    this.errors = [];
    this.successCount = 0;
    this.processedDocs = new Set();
  }

  /**
   * Record a successful document processing
   * @param {string} docId - The document ID that was processed
   */
  recordSuccess(docId) {
    if (!this.processedDocs.has(docId)) {
      this.processedDocs.add(docId);
      this.successCount++;
    }
  }

  /**
   * Record an error during processing
   * @param {number} rowIndex - The row number in the sheet (1-based)
   * @param {string} docId - The document ID that failed
   * @param {string} errorMsg - The error message
   * @param {Object} rowData - Optional full row data for context
   */
  recordError(rowIndex, docId, errorMsg, rowData = null) {
    this.errors.push({
      rowIndex: rowIndex,
      docId: docId,
      error: errorMsg,
      timestamp: new Date(),
      rowData: rowData
    });
  }

  /**
   * Check if there were any errors
   * @returns {boolean}
   */
  hasErrors() {
    return this.errors.length > 0;
  }

  /**
   * Get error count
   * @returns {number}
   */
  getErrorCount() {
    return this.errors.length;
  }

  /**
   * Get success count
   * @returns {number}
   */
  getSuccessCount() {
    return this.successCount;
  }

  /**
   * Write errors to the error sheet
   * Creates the sheet if it doesn't exist
   */
  writeErrorsToSheet() {
    if (this.errors.length === 0) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName(this.errorSheetName);

    // Create sheet if it doesn't exist
    if (!errorSheet) {
      errorSheet = ss.insertSheet(this.errorSheetName);
      // Add headers
      errorSheet.getRange(1, 1, 1, 5).setValues([[
        "Timestamp", "Row Number", "Document ID", "Error Message", "Status"
      ]]).setFontWeight("bold");
    }

    // Prepare error data
    const errorData = this.errors.map(err => [
      err.timestamp,
      err.rowIndex + 2, // +2 because row 1 is header, rowIndex is 0-based
      err.docId,
      err.error,
      "Failed"
    ]);

    // Append to sheet
    const lastRow = errorSheet.getLastRow();
    errorSheet.getRange(lastRow + 1, 1, errorData.length, 5).setValues(errorData);

    // Format timestamp column
    errorSheet.getRange(lastRow + 1, 1, errorData.length, 1)
      .setNumberFormat("yyyy-mm-dd hh:mm:ss");
  }

  /**
   * Show a modal dialog with processing results
   * @param {string} processType - Type of process (e.g., "Classes", "Goals")
   */
  showResultsDialog(processType = "Documents") {
    const totalProcessed = this.successCount + this.errors.length;
    
    let html = `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            body {
              font-family: Arial, sans-serif;
              padding: 20px;
              margin: 0;
            }
            .success {
              color: #0f9d58;
              font-size: 18px;
              font-weight: bold;
              margin-bottom: 15px;
            }
            .error {
              color: #d93025;
              font-size: 18px;
              font-weight: bold;
              margin-bottom: 15px;
            }
            .summary {
              background-color: #f1f3f4;
              padding: 15px;
              border-radius: 5px;
              margin-bottom: 15px;
            }
            .summary-item {
              margin: 8px 0;
              font-size: 14px;
            }
            .error-list {
              max-height: 200px;
              overflow-y: auto;
              background-color: #fef7e0;
              padding: 10px;
              border-radius: 5px;
              margin-top: 10px;
            }
            .error-item {
              margin: 5px 0;
              font-size: 12px;
              padding: 5px;
              border-bottom: 1px solid #e8eaed;
            }
            button {
              background-color: #1a73e8;
              color: white;
              border: none;
              padding: 10px 20px;
              font-size: 14px;
              border-radius: 4px;
              cursor: pointer;
              margin-top: 15px;
            }
            button:hover {
              background-color: #1557b0;
            }
          </style>
        </head>
        <body>
    `;

    if (this.errors.length === 0) {
      html += `
          <div class="success">✓ All ${processType} Processed Successfully!</div>
          <div class="summary">
            <div class="summary-item"><strong>Total Documents:</strong> ${totalProcessed}</div>
            <div class="summary-item"><strong>Successful:</strong> ${this.successCount}</div>
            <div class="summary-item"><strong>Failed:</strong> 0</div>
          </div>
      `;
    } else {
      html += `
          <div class="error">⚠ Processing Completed with Errors</div>
          <div class="summary">
            <div class="summary-item"><strong>Total Documents:</strong> ${totalProcessed}</div>
            <div class="summary-item"><strong>Successful:</strong> ${this.successCount}</div>
            <div class="summary-item"><strong>Failed:</strong> ${this.errors.length}</div>
          </div>
          <p style="font-size: 14px;">The following documents could not be processed:</p>
          <div class="error-list">
      `;

      this.errors.forEach(err => {
        html += `
            <div class="error-item">
              <strong>Row ${err.rowIndex + 2}:</strong> Doc ID: ${err.docId}<br>
              <span style="color: #5f6368;">Error: ${err.error}</span>
            </div>
        `;
      });

      html += `
          </div>
          <p style="font-size: 13px; color: #5f6368; margin-top: 15px;">
            Error details have been saved to the "${this.errorSheetName}" tab.
          </p>
      `;
    }

    html += `
          <button onclick="google.script.host.close()">Close</button>
        </body>
      </html>
    `;

    const htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(500)
      .setHeight(400);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `${processType} Processing Results`);
  }
}


/**
 * Top-level function to trigger class-adding process.
 * Can be called from menu (WEPhelper.addClassesToTable).
 * @param {Array} [classData] - Optional array of data, else pulls from sheet.
 */
function addClassesToTable(classData) {
  const settings = getSettingsInstance("config");
  const tracker = new DocumentProcessingTracker("Errors");
  
  try {
    let theData;

    if (classData) {
      theData = classData;
    } else {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("forClasses");
      theData = sheet.getRange(2, 9, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    }
    
    const rowNum = parseInt(settings.getSetting("classRowNum") || "0");
    const results = nsAddClasses.addClasses(theData, rowNum, tracker);
    settings.setSetting("classRowNum", results.classRowNum);

    // Write errors to sheet if any occurred
    if (tracker.hasErrors()) {
      tracker.writeErrorsToSheet();
    }

    // Show results dialog
    tracker.showResultsDialog("Classes");

    // Also show toast for quick feedback
    const ui = SpreadsheetApp.getActiveSpreadsheet();
    if (parseInt(results.classRowNum) === 0) {
      if (tracker.hasErrors()) {
        ui.toast(`Completed with ${tracker.getErrorCount()} error(s). Check the Errors tab.`, "Completed with Errors", 5);
      } else {
        ui.toast("All documents processed successfully!", "All done!", 3);
      }
    } else {
      ui.toast("Exceeded time limit! Please run addClasses again.", "NOT done yet!", -1);
    }

  } catch (err) {
    logIt({
      level: "severe",
      theMsg: "Error in addClassesToTable",
      error: err.toString()
    });
    SpreadsheetApp.getUi().alert("Critical Error", 
      "An unexpected error occurred: " + err.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/** 
 * Namespace for functions related to adding classes to WEP documents 
 */
const nsAddClasses = (() => {

  /**
   * Add student classes to the table on their WEP document.
   * Uses timing logic to avoid timeouts.
   * 
   * @param {Array} theData - Array of student/class data.
   * @param {number} rowNum - Starting index to process.
   * @param {DocumentProcessingTracker} tracker - Error tracking instance.
   * @return {Object} - { classRowNum: string }
   */
  function addClasses(theData, rowNum, tracker) {
    const start = new Date();
    const headerStyle = { [DocumentApp.Attribute.BOLD]: true };
    const rowStyle = { [DocumentApp.Attribute.BOLD]: false };
    const length = theData.length;

    let results = {};

    try {
      for (let i = rowNum; i < length;) {
        const currentDocId = theData[i][0];
        let tbl = null;
        
        // Try to get the table, catch errors for this specific document
        try {
          tbl = getTheTable_(currentDocId);
        } catch (err) {
          // Record error and skip all rows for this document
          tracker.recordError(i, currentDocId, err.message, theData[i]);
          
          logIt({
            level: "warning",
            theMsg: `Skipping document due to error: ${currentDocId}`,
            error: err.toString()
          });
          
          // Skip all rows with the same docId
          do {
            i++;
          } while (i < length && theData[i][0] === currentDocId);
          
          continue; // Move to next document
        }

        // Process all rows for this document
        const docStartIndex = i;
        do {
          try {
            const tr = tbl.appendTableRow().setAttributes(rowStyle);
            for (let j = 1; j < 4; j++) {
              tr.appendTableCell(theData[i][j]);
            }
            i++;
            if (i === length) break;
          } catch (err) {
            // Error adding specific row
            tracker.recordError(i, currentDocId, `Error adding row: ${err.message}`, theData[i]);
            logIt({
              level: "warning",
              theMsg: `Error adding row ${i} to doc ${currentDocId}`,
              error: err.toString()
            });
            i++;
            if (i === length) break;
          }
        } while (i < length && theData[i - 1][0] === theData[i][0]); // same doc ID

        // Set header style for this document's table
        try {
          tbl.getRow(0).setAttributes(headerStyle);
          tracker.recordSuccess(currentDocId);
        } catch (err) {
          logIt({
            level: "warning",
            theMsg: `Error setting header style for doc ${currentDocId}`,
            error: err.toString()
          });
        }

        // Check time limit
        if (isTimeUp_(start)) {
          results = { classRowNum: i.toString() }; // Save progress
          logIt({
            level: "warning",
            theMsg: "Execution time limit reached in addClasses()",
            error: `Stopped at row ${i}`
          });
          break;
        }
      }

      // If all done
      if (!results.classRowNum) {
        results = { classRowNum: "0" };
      }

      return results;

    } catch (err) {
      logIt({
        level: "severe",
        theMsg: "Error in addClasses()",
        error: err.toString()
      });
      return { classRowNum: rowNum.toString() }; // return where it left off
    }
  }

  /**
   * Finds the target table in the WEP document.
   * @param {string} docId - The ID of the Google Doc.
   * @returns {GoogleAppsScript.Document.Table} - The matching table.
   */
  function getTheTable_(docId) {
    try {
      const doc = DocumentApp.openById(docId);
      const tables = doc.getBody().getTables();

      for (let i = 0; i < tables.length; i++) {
        const cell = tables[i].getCell(0, 0).getText();
        if (cell.trim() === "Service Area") {
          return tables[i];
        }
      }

      throw new Error(`No matching table found in doc: ${docId}`);
    } catch (err) {
      // Re-throw with more specific message
      if (err.message.includes("perhaps it was deleted")) {
        throw new Error("Document not found or inaccessible");
      } else if (err.message.includes("No matching table")) {
        throw new Error("Required table not found in document");
      } else {
        throw new Error(`Unable to access document: ${err.message}`);
      }
    }
  }

  return {
    addClasses
  };

})();


// // Example usage for Goals (demonstrating reusability)
// /**
//  * Example function for adding goals - follows same pattern
//  * @param {Array} [goalData] - Optional array of goal data
//  */
// function addGoalsToTable(goalData) {
//   const settings = getSettingsInstance("config");
//   const tracker = new DocumentProcessingTracker("Errors");
  
//   try {
//     let theData;
//     if (goalData) {
//       theData = goalData;
//     } else {
//       const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("forGoals");
//       theData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
//     }
    
//     const rowNum = parseInt(settings.getSetting("goalRowNum") || "0");
//     const results = nsAddGoals.addGoals(theData, rowNum, tracker);
//     settings.setSetting("goalRowNum", results.goalRowNum);

//     if (tracker.hasErrors()) {
//       tracker.writeErrorsToSheet();
//     }

//     tracker.showResultsDialog("Goals");

//   } catch (err) {
//     logIt({
//       level: "severe",
//       theMsg: "Error in addGoalsToTable",
//       error: err.toString()
//     });
//     SpreadsheetApp.getUi().alert("Critical Error", 
//       "An unexpected error occurred: " + err.toString(), 
//       SpreadsheetApp.getUi().ButtonSet.OK);
//   }
// }