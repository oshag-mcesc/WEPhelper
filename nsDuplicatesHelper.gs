/**
 * @file This script provides utility functions for Google Sheets,
 * specifically for highlighting duplicate rows based on specified criteria.
 * @namespace SheetUtils
 */
var SheetUtils = SheetUtils || {}; // Ensure SheetUtils namespace exists

/**
 * @constant {Array<string>} REQUIRED_HEADERS
 * @description The column headers required for identifying duplicate rows.
 * These headers must exist in the sheet's first row.
 */
SheetUtils.REQUIRED_HEADERS = ["Student ID", "Course Name", "Teacher Name"];

/**
 * Finds the column indices for the required headers in the sheet.
 * 
 * @memberof SheetUtils
 * @param {Array<Array>} headerRow - The first row of the sheet containing column headers.
 * @returns {Object|null} An object mapping header names to their 0-indexed column positions,
 *                        or null if any required headers are missing.
 * @private
 * @example
 * // Returns: { "Student ID": 2, "Course Name": 3, "Teacher Name": 4 }
 */
SheetUtils.findHeaderIndices = (headerRow) => {
  const indices = {};
  const normalizedHeaders = headerRow.map(h => String(h).trim().toLowerCase());
  
  for (const requiredHeader of SheetUtils.REQUIRED_HEADERS) {
    const normalizedRequired = requiredHeader.trim().toLowerCase();
    const index = normalizedHeaders.indexOf(normalizedRequired);
    
    if (index === -1) {
      return null;
    }
    
    indices[requiredHeader] = index;
  }
  
  return indices;
};

/**
 * Highlights duplicate rows in the active sheet based on a combination of
 * "Student ID", "Course Name", and "Teacher Name".
 * Duplicate rows (including the first occurrence of a duplicate key) will be
 * highlighted in a light orange color (`#FFD966`).
 *
 * The function dynamically locates the required columns by searching for their
 * header names in the first row. This allows the columns to be in any order
 * and in any position within the sheet.
 *
 * This function clears any existing background colors in the data range
 * before applying new highlights.
 *
 * @memberof SheetUtils
 * @returns {void}
 * @example
 * // To call this function from a sheet script after adding the library:
 * SheetUtils.highlightDuplicateRows();
 */
SheetUtils.highlightDuplicateRows = () => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();
  
  try {
    const range = sheet.getDataRange();
    const values = range.getValues();
    
    // If there's no data at all, log error and exit
    if (values.length === 0) {
      logIt({
        level: "severe",
        theMsg: "No data found in sheet",
        error: new Error(`Sheet "${sheetName}" contains no data`)
      });
      return;
    }
    
    // Find the column indices based on header names
    const headerRow = values[0];
    const headerIndices = SheetUtils.findHeaderIndices(headerRow);
    
    if (!headerIndices) {
      const missingHeaders = SheetUtils.REQUIRED_HEADERS.join('", "');
      const errorMsg = `Required column headers not found in sheet "${sheetName}". Please ensure the sheet has columns named: "${missingHeaders}".`;
      logIt({
        level: "severe",
        theMsg: "Required column headers not found",
        error: new Error(errorMsg)
      });
      SpreadsheetApp.getUi().alert(errorMsg);
      return;
    }
    
    // Clear existing background colors to ensure a clean highlight
    range.setBackground(null);
    
    // If there's only a header row, exit (not an error condition)
    if (values.length <= 1) {
      return;
    }
    
    // A Map to store unique keys and the row number (1-indexed) of their first occurrence
    const seenKeys = new Map();
    // A Set to store row numbers (1-indexed) of all rows identified as duplicates
    const duplicateRows = new Set();
    
    // Iterate over each row, starting from the second row (index 1) to skip the header
    values.forEach((row, index) => {
      const currentRowNumber = index + 1;
      if (index === 0) return; // Skip the header row (index 0)
      
      // Extract the values for the unique key combination using the dynamically found column indices
      const studentId = row[headerIndices["Student ID"]];
      const courseName = row[headerIndices["Course Name"]];
      const teacherName = row[headerIndices["Teacher Name"]];
      
      // Create a unique key by concatenating the relevant fields
      const uniqueKey = `${studentId}|~|${courseName}|~|${teacherName}`;
      
      // Check if the unique key has been seen before
      if (seenKeys.has(uniqueKey)) {
        // If the key exists, it means we've found a duplicate
        duplicateRows.add(currentRowNumber);
        duplicateRows.add(seenKeys.get(uniqueKey));
      } else {
        // If the key is new, add it to the map with its corresponding row number
        seenKeys.set(uniqueKey, currentRowNumber);
      }
    });
    
    // Apply highlighting to all identified duplicate rows
    if (duplicateRows.size > 0) {
      const highlightColor = "#FFD966"; // Light orange color for duplicates
      duplicateRows.forEach(rowNum => {
        sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).setBackground(highlightColor);
      });
      Logger.log(`Highlighted ${duplicateRows.size} duplicate rows.`);
    } else {
      // No duplicates found - notify the user
      SpreadsheetApp.getUi().alert(
        "No Duplicates Found",
        `No duplicate rows were found in sheet "${sheetName}".`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    // Log any errors that occur during script execution
    logIt({
      level: "severe",
      theMsg: "Failed to highlight duplicate rows",
      error: error
    });
    SpreadsheetApp.getUi().alert("An error occurred while highlighting duplicates. Please check the error logs for details.");
  }
};