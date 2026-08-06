
/**
 * Pushes initial goals data from the spreadsheet to documents
 * Retrieves data from the 'forGoals' sheet starting at cell L1
 */
function pushInitialGoals() {
  const tracker = new DocumentProcessingTracker("Errors");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forGoals');
    if (!ss) {
      throw new Error("Sheet 'forGoals' not found");
    }
    
    const info = ss.getRange("L1").getDataRegion().getValues();
    if (!info || info.length < 2) {
      throw new Error("No data found or insufficient data in range");
    }
    
    // Process the goals with error tracking
    nsAppendGoals.appendTable(info, tracker);
    
    // Write errors to sheet if any occurred
    if (tracker.hasErrors()) {
      tracker.writeErrorsToSheet();
    }
    
    // Show results dialog
    tracker.showResultsDialog("Goals");
    
    // Also log success
    if (!tracker.hasErrors()) {
      logIt({
        level: "info", 
        theMsg: `Successfully appended initial goals to ${tracker.getSuccessCount()} documents!`
      });
    }
    
  } catch (err) {
    const errorMsg = `Error in pushInitialGoals: ${err.message}`;
    console.error(errorMsg);
    logIt({level: "severe", theMsg: errorMsg, error: err.toString()});
    SpreadsheetApp.getUi().alert("Critical Error", 
      "An unexpected error occurred: " + err.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Namespace for appending goals functionality
 * @namespace nsAppendGoals
 */
var nsAppendGoals = (function() {
  
  /**
   * Appends information to multiple documents as tables with batch processing.
   * Each row in the data array creates a table in the corresponding document.
   * 
   * @param {Array<Array>} info - 2D array where:
   *   - First row contains headers: [docID, header1, header2, ...]
   *   - Subsequent rows contain: [docID, value1, value2, ...]
   * @param {DocumentProcessingTracker} tracker - Error tracking instance
   * @return {Object} Result object with properties:
   *   - {boolean} done - Whether the operation completed successfully
   *   - {string} info - Error message if done is false
   * @memberof nsAppendGoals
   */
  function appendTable(info, tracker) {
    const result = {
      done: false,
      info: ""
    };
    
    try {
      // Validate input
      if (!info || !Array.isArray(info) || info.length < 2) {
        throw new Error("Invalid input: info must be a 2D array with at least 2 rows");
      }
      
      // Extract headers (skip docID column)
      const headers = info[0].slice(1);
      const dataRows = info.slice(1);
      
      if (headers.length === 0) {
        throw new Error("No table headers found");
      }
      
      // Group data by document ID for batch processing
      const docDataMap = groupDataByDocId(dataRows);
      
      // Process each document
      const docIds = Object.keys(docDataMap);
      logIt({
        level: "info", 
        theMsg: `Starting to process ${docIds.length} documents with ${dataRows.length} total rows`
      });
      
      let rowIndex = 1; // Start at 1 because row 0 is headers
      
      for (const docId of docIds) {
        try {
          processDocument(docId, docDataMap[docId], headers);
          tracker.recordSuccess(docId);
          
        } catch (docError) {
          // Find the first row index for this docId in the original data
          const firstRowForDoc = dataRows.findIndex(row => row[0] === docId);
          const actualRowIndex = firstRowForDoc >= 0 ? firstRowForDoc : rowIndex;
          
          tracker.recordError(
            actualRowIndex,
            docId,
            docError.message,
            dataRows[firstRowForDoc]
          );
          
          logIt({
            level: "warning", 
            theMsg: `Failed to process document ${docId}: ${docError.message}`
          });
        }
        
        rowIndex++;
      }
      
      // Determine overall success
      if (!tracker.hasErrors()) {
        result.done = true;
        logIt({
          level: "info", 
          theMsg: `Successfully processed all ${tracker.getSuccessCount()} documents`
        });
      } else {
        result.done = false;
        result.info = `Partial success: ${tracker.getSuccessCount()} succeeded, ${tracker.getErrorCount()} failed`;
        logIt({level: "warning", theMsg: result.info});
      }
      
    } catch (err) {
      console.error("Error in appendTable:", err);
      result.done = false;
      result.info = err.message;
      logIt({level: "severe", theMsg: `Critical error in appendTable: ${err.message}`});
    }
    
    return result;
  }

  /**
   * Groups data rows by document ID to enable batch processing
   * @param {Array<Array>} dataRows - Array of data rows where first column is docID
   * @return {Object} Map of docID -> array of data rows for that document
   * @private
   */
  function groupDataByDocId(dataRows) {
    const docDataMap = {};
    
    for (const row of dataRows) {
      const docId = row[0];
      if (!docId) {
        console.warn("Skipping row with empty document ID");
        continue;
      }
      
      if (!docDataMap[docId]) {
        docDataMap[docId] = [];
      }
      
      // Store the data portion (exclude docID)
      docDataMap[docId].push(row.slice(1));
    }
    
    return docDataMap;
  }

  /**
   * Processes a single document by adding all its associated tables
   * @param {string} docId - Google Document ID
   * @param {Array<Array>} docData - Array of data rows for this document
   * @param {Array<string>} headers - Column headers for the table
   * @private
   */
  function processDocument(docId, docData, headers) {
    let doc;
    
    try {
      // Open document once for all operations
      doc = DocumentApp.openById(docId);
      const body = doc.getBody();
      
      // Process each data row as a separate table
      for (const dataRow of docData) {
        createAndFormatTable(body, headers, dataRow);
      }
      
    } catch (err) {
      // Provide more specific error messages
      if (err.message.includes("perhaps it was deleted")) {
        throw new Error("Document not found or inaccessible");
      } else if (err.message.includes("Permission denied")) {
        throw new Error("Permission denied - cannot access document");
      } else {
        throw new Error(`Document processing failed: ${err.message}`);
      }
    } finally {
      // Ensure document is saved and closed
      if (doc) {
        try {
          doc.saveAndClose();
        } catch (saveErr) {
          console.warn(`Warning: Could not save/close document ${docId}:`, saveErr);
        }
      }
    }
  }

  /**
   * Creates and formats a single table in the document body
   * @param {GoogleAppsScript.Document.Body} body - Document body element
   * @param {Array<string>} headers - Column headers
   * @param {Array} dataRow - Data values for the table
   * @private
   */
  function createAndFormatTable(body, headers, dataRow) {
    // Create table data with proper structure
    const tableData = [];
    
    // Ensure we don't exceed available data
    const maxColumns = Math.min(headers.length, dataRow.length);
    
    for (let i = 0; i < maxColumns; i++) {
      tableData.push([
        headers[i] || '',  // Header cell
        dataRow[i] || ''   // Data cell
      ]);
    }
    
    // Skip empty tables
    if (tableData.length === 0) {
      console.warn("Skipping empty table creation");
      return;
    }
    
    // Create and format the table
    const table = body.appendTable(tableData);
    
    // Apply formatting
    table.setColumnWidth(0, 75);
    
    // Set header column background color
    const numRows = table.getNumRows();
    for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
      const headerCell = table.getCell(rowIndex, 0);
      headerCell.setBackgroundColor("#E0E0E0");
    }
  }
  
  // Return public interface
  return {
    appendTable: appendTable
  };
  
})();