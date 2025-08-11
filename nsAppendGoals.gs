
/**
 * Namespace for appending goals functionality
 * @namespace nsAppendGoals
 */
var nsAppendGoals = (function() {
  
  /**
   * Appends information to multiple documents as tables with batch processing for improved performance.
   * Each row in the data array creates a table in the corresponding document.
   * 
   * @param {Array<Array>} info - 2D array where:
   *   - First row contains headers: [docID, header1, header2, ...]
   *   - Subsequent rows contain: [docID, value1, value2, ...]
   * @return {Object} Result object with properties:
   *   - {boolean} done - Whether the operation completed successfully
   *   - {string} info - Error message if done is false
   *   - {number} processedDocs - Number of documents successfully processed
   *   - {Array<string>} failedDocs - Array of document IDs that failed processing
   * @memberof nsAppendGoals
   */
  function appendTable(info) {
  const result = {
    done: false,
    info: "",
    processedDocs: 0,
    failedDocs: []
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
    logIt({level: "info", theMsg: `Starting to process ${docIds.length} documents with ${dataRows.length} total rows`});
    
    for (const docId of docIds) {
      try {
        processDocument(docId, docDataMap[docId], headers);
        result.processedDocs++;
      } catch (docError) {
        console.error(`Failed to process document ${docId}:`, docError);
        result.failedDocs.push(docId);
        logIt({level: "error", theMsg: `Failed to process document ${docId}: ${docError.message}`});
      }
    }
    
    // Determine overall success
    if (result.failedDocs.length === 0) {
      result.done = true;
      logIt({level: "info", theMsg: `Successfully processed all ${result.processedDocs} documents`});
    } else if (result.processedDocs > 0) {
      result.done = false;
      result.info = `Partial success: ${result.processedDocs} succeeded, ${result.failedDocs.length} failed`;
      logIt({level: "warn", theMsg: result.info});
    } else {
      result.done = false;
      result.info = "All documents failed to process";
      logIt({level: "error", theMsg: result.info});
    }
    
  } catch (err) {
    console.error("Error in appendTable:", err);
    result.done = false;
    result.info = err.message;
    logIt({level: "error", theMsg: `Critical error in appendTable: ${err.message}`});
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
    throw new Error(`Document processing failed: ${err.message}`);
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
  }  // Close createAndFormatTable function
  
  // Return public interface
  return {
    appendTable: appendTable
  };
  
})();  // Close the IIFE